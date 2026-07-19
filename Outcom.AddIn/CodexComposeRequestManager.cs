using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outcom.AddIn
{
    /// <summary>
    /// Exécute les propositions de réponse sans fenêtre modale. Les objets Outlook restent
    /// cantonnés au thread principal et ne sont utilisés qu'au moment de l'insertion finale.
    /// </summary>
    internal sealed class CodexComposeRequestManager : IDisposable
    {
        private readonly ThisAddIn _addIn;
        private readonly CodexOperationTracker _tracker;
        private readonly Control _dispatcher;
        private readonly Dictionary<long, ComposeRequest> _requests =
            new Dictionary<long, ComposeRequest>();
        private int _disposeState;

        internal CodexComposeRequestManager(
            ThisAddIn addIn,
            CodexOperationTracker tracker)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
            _tracker = tracker ?? throw new ArgumentNullException(nameof(tracker));
            _dispatcher = new Control();
            _dispatcher.CreateControl();
            IntPtr unusedHandle = _dispatcher.Handle;
        }

        internal void Start(
            Outlook.Inspector inspector,
            Outlook.MailItem draft,
            OutlookComposeContext context)
        {
            if (inspector == null)
            {
                throw new ArgumentNullException(nameof(inspector));
            }

            StartCore(inspector, null, draft, context);
        }

        internal void Start(
            Outlook.Explorer explorer,
            Outlook.MailItem draft,
            OutlookComposeContext context)
        {
            if (explorer == null)
            {
                throw new ArgumentNullException(nameof(explorer));
            }

            StartCore(null, explorer, draft, context);
        }

        private void StartCore(
            Outlook.Inspector inspector,
            Outlook.Explorer explorer,
            Outlook.MailItem draft,
            OutlookComposeContext context)
        {
            ThrowIfDisposed();

            if (draft == null)
            {
                throw new ArgumentNullException(nameof(draft));
            }

            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            if (!context.HasReplySource &&
                string.IsNullOrWhiteSpace(context.DraftBody))
            {
                throw new InvalidOperationException(
                    "Outcom n'a détecté ni message source ni orientation dans le brouillon. " +
                    "La proposition n'a pas été lancée afin d'éviter une réponse hors contexte.");
            }

            string prompt = context.BuildPrompt();
            if (string.IsNullOrWhiteSpace(prompt))
            {
                throw new ArgumentException("La demande Codex est vide.", nameof(prompt));
            }

            long identity = GetComIdentity(draft);
            if (_requests.ContainsKey(identity))
            {
                throw new InvalidOperationException(
                    "Une proposition Codex est déjà en cours pour ce message.");
            }

            bool insertAtBeginningOnConflict = _addIn
                .GetCodexService()
                .GetGlobalContext()
                .InsertProposalAtBeginningOnConflict;
            var request = new ComposeRequest(
                this,
                identity,
                inspector,
                explorer,
                draft,
                prompt,
                BuildDescription(context.Subject, context.IsForward),
                context.ReplaceResponseSection,
                context.ExpectedResponseDraft,
                insertAtBeginningOnConflict);
            _requests.Add(identity, request);
            try
            {
                request.Start();
            }
            catch
            {
                _requests.Remove(identity);
                request.ReturnOwnershipAfterFailedStart();
                throw;
            }
        }

        private static string BuildDescription(string subject, bool isForward)
        {
            string value = string.IsNullOrWhiteSpace(subject)
                ? "message sans objet"
                : subject.Trim().Replace('\r', ' ').Replace('\n', ' ');
            if (value.Length > 70)
            {
                value = value.Substring(0, 67) + "…";
            }

            return (isForward
                ? "Message d'accompagnement — "
                : "Proposition de réponse — ") + value;
        }

        private void DispatchCompletion(
            ComposeRequest request,
            Task<CodexTurnResult> generationTask)
        {
            if (Volatile.Read(ref _disposeState) != 0)
            {
                request.DisposeCancellation();
                return;
            }

            try
            {
                _dispatcher.BeginInvoke(new Action(
                    () => CompleteOnOutlookThread(request, generationTask)));
            }
            catch (InvalidOperationException)
            {
                request.DisposeCancellation();
            }
        }

        private void CompleteOnOutlookThread(
            ComposeRequest request,
            Task<CodexTurnResult> generationTask)
        {
            if (!_requests.Remove(request.Identity))
            {
                request.DisposeCancellation();
                return;
            }

            try
            {
                RequestStopReason stopReason = request.StopReason;
                if (stopReason == RequestStopReason.MessageSent)
                {
                    request.Operation.Skip("le message a été envoyé avant la réponse");
                    LocalLogger.Info(
                        "Proposition Codex ignorée : le message a été envoyé pendant la génération.");
                    return;
                }

                if (stopReason == RequestStopReason.InspectorClosed)
                {
                    request.Operation.Skip("la fenêtre de rédaction a été fermée");
                    return;
                }

                if (stopReason == RequestStopReason.UserCanceled || generationTask.IsCanceled)
                {
                    request.Operation.Cancel("demande annulée");
                    return;
                }

                if (generationTask.IsFaulted)
                {
                    Exception error = generationTask.Exception == null
                        ? null
                        : generationTask.Exception.GetBaseException();
                    request.Operation.Fail(GetUserFacingFailure(error));
                    LocalLogger.Error(
                        "Échec d'une proposition Codex asynchrone (" +
                        (error == null ? "Unknown" : error.GetType().Name) + ").");
                    return;
                }

                CodexTurnResult result = generationTask.Result;
                if (result == null || string.IsNullOrWhiteSpace(result.Text))
                {
                    request.Operation.Fail("Codex n'a retourné aucun texte");
                    return;
                }

                bool sent;
                try
                {
                    sent = request.Draft.Sent;
                }
                catch (COMException)
                {
                    request.Operation.Skip("le message n'est plus accessible dans Outlook");
                    return;
                }

                if (sent)
                {
                    request.Operation.Skip("le message a déjà été envoyé");
                    LocalLogger.Info(
                        "Proposition Codex ignorée : Outlook indique que le message est envoyé.");
                    return;
                }

                try
                {
                    bool insertedAtBeginning;
                    if (request.Inspector != null)
                    {
                        insertedAtBeginning = OutlookComposeContext.ApplyProposal(
                            request.Inspector,
                            request.Draft,
                            result.Text,
                            request.ReplaceResponseDraft,
                            request.ExpectedResponseDraft,
                            request.InsertAtBeginningOnConflict);
                    }
                    else
                    {
                        insertedAtBeginning = OutlookComposeContext.ApplyProposal(
                            request.Explorer,
                            request.Draft,
                            result.Text,
                            request.ReplaceResponseDraft,
                            request.ExpectedResponseDraft,
                            request.InsertAtBeginningOnConflict);
                    }

                    request.Operation.Complete(insertedAtBeginning
                        ? "proposition insérée au début du message"
                        : "projet remplacé par le message complet");
                    LocalLogger.Info(
                        insertedAtBeginning
                            ? "Proposition Codex insérée au début du message."
                            : "Projet de réponse remplacé par le message complet généré.");
                }
                catch (InvalidOperationException exception)
                {
                    request.Operation.Skip(exception.Message);
                }
            }
            catch (Exception exception)
            {
                request.Operation.Fail(GetUserFacingFailure(exception));
                LocalLogger.Error(
                    "Impossible de finaliser une proposition Codex (" +
                    exception.GetType().Name + ").");
            }
            finally
            {
                request.ReleaseOutlookObjects();
                request.DisposeCancellation();
            }
        }

        private static string GetUserFacingFailure(Exception exception)
        {
            if (exception == null)
            {
                return "erreur inconnue";
            }

            if (exception is CodexAppServerException ||
                exception is InvalidOperationException ||
                exception is ArgumentException)
            {
                return exception.Message;
            }

            return "consultez le journal Outcom";
        }

        private static long GetComIdentity(object value)
        {
            IntPtr identity = IntPtr.Zero;
            try
            {
                identity = Marshal.GetIUnknownForObject(value);
                return identity.ToInt64();
            }
            finally
            {
                if (identity != IntPtr.Zero)
                {
                    Marshal.Release(identity);
                }
            }
        }

        private void ThrowIfDisposed()
        {
            if (Volatile.Read(ref _disposeState) != 0)
            {
                throw new ObjectDisposedException(nameof(CodexComposeRequestManager));
            }
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposeState, 1) != 0)
            {
                return;
            }

            var requests = new List<ComposeRequest>(_requests.Values);
            _requests.Clear();
            foreach (ComposeRequest request in requests)
            {
                request.AbortForShutdown();
            }

            _dispatcher.Dispose();
        }

        private enum RequestStopReason
        {
            None,
            UserCanceled,
            MessageSent,
            InspectorClosed,
            Shutdown
        }

        private sealed class ComposeRequest
        {
            private readonly CodexComposeRequestManager _manager;
            private readonly string _prompt;
            private readonly string _description;
            private readonly CancellationTokenSource _cancellation =
                new CancellationTokenSource();
            private readonly object _stateLock = new object();
            private Outlook.ItemEvents_10_Event _draftEvents;
            private Outlook.InspectorEvents_10_Event _inspectorEvents;
            private RequestStopReason _stopReason;
            private int _outlookObjectsReleased;
            private int _cancellationDisposed;

            internal ComposeRequest(
                CodexComposeRequestManager manager,
                long identity,
                Outlook.Inspector inspector,
                Outlook.Explorer explorer,
                Outlook.MailItem draft,
                string prompt,
                string description,
                bool replaceResponseDraft,
                string expectedResponseDraft,
                bool insertAtBeginningOnConflict)
            {
                _manager = manager;
                Identity = identity;
                Inspector = inspector;
                Explorer = explorer;
                Draft = draft;
                _prompt = prompt;
                _description = description;
                ReplaceResponseDraft = replaceResponseDraft;
                ExpectedResponseDraft = expectedResponseDraft;
                InsertAtBeginningOnConflict = insertAtBeginningOnConflict;
            }

            internal long Identity { get; private set; }

            internal Outlook.Inspector Inspector { get; private set; }

            internal Outlook.Explorer Explorer { get; private set; }

            internal Outlook.MailItem Draft { get; private set; }

            internal CodexOperationHandle Operation { get; private set; }

            internal bool ReplaceResponseDraft { get; private set; }

            internal string ExpectedResponseDraft { get; private set; }

            internal bool InsertAtBeginningOnConflict { get; private set; }

            internal RequestStopReason StopReason
            {
                get
                {
                    lock (_stateLock)
                    {
                        return _stopReason;
                    }
                }
            }

            internal void Start()
            {
                _draftEvents = (Outlook.ItemEvents_10_Event)Draft;
                _draftEvents.Send += Draft_Send;
                if (Inspector != null)
                {
                    _inspectorEvents = (Outlook.InspectorEvents_10_Event)Inspector;
                    _inspectorEvents.Close += Inspector_Close;
                }
                Operation = _manager._tracker.Begin(_description, CancelByUser);

                Task<CodexTurnResult> task = _manager._addIn
                    .GetCodexService()
                    .GenerateComposeReplyAsync(_prompt, null, _cancellation.Token);
                task.ContinueWith(
                    completedTask => _manager.DispatchCompletion(this, completedTask),
                    CancellationToken.None,
                    TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
            }

            private void Draft_Send(ref bool cancel)
            {
                RequestStop(RequestStopReason.MessageSent);
            }

            private void Inspector_Close()
            {
                RequestStop(RequestStopReason.InspectorClosed);
            }

            private void CancelByUser()
            {
                RequestStop(RequestStopReason.UserCanceled);
            }

            private void RequestStop(RequestStopReason reason)
            {
                lock (_stateLock)
                {
                    if (_stopReason != RequestStopReason.None)
                    {
                        return;
                    }

                    _stopReason = reason;
                }

                try
                {
                    _cancellation.Cancel();
                }
                catch (ObjectDisposedException)
                {
                }
            }

            internal void ReturnOwnershipAfterFailedStart()
            {
                RequestStop(RequestStopReason.UserCanceled);
                if (Operation != null)
                {
                    Operation.Fail("la demande n'a pas pu démarrer");
                }

                DetachEvents();
                Draft = null;
                Inspector = null;
                Explorer = null;
                DisposeCancellation();
            }

            internal void AbortForShutdown()
            {
                RequestStop(RequestStopReason.Shutdown);
                if (Operation != null)
                {
                    Operation.Cancel("arrêt d'Outlook");
                }

                ReleaseOutlookObjects();
            }

            internal void ReleaseOutlookObjects()
            {
                if (Interlocked.Exchange(ref _outlookObjectsReleased, 1) != 0)
                {
                    return;
                }

                DetachEvents();
                ReleaseComObject(Draft);
                ReleaseComObject(Inspector);
                ReleaseComObject(Explorer);
                Draft = null;
                Inspector = null;
                Explorer = null;
            }

            private void DetachEvents()
            {
                if (_draftEvents != null)
                {
                    try
                    {
                        _draftEvents.Send -= Draft_Send;
                    }
                    catch (COMException)
                    {
                    }
                }

                if (_inspectorEvents != null)
                {
                    try
                    {
                        _inspectorEvents.Close -= Inspector_Close;
                    }
                    catch (COMException)
                    {
                    }
                }

                _draftEvents = null;
                _inspectorEvents = null;
            }

            internal void DisposeCancellation()
            {
                if (Interlocked.Exchange(ref _cancellationDisposed, 1) == 0)
                {
                    _cancellation.Dispose();
                }
            }

            private static void ReleaseComObject(object value)
            {
                if (value == null || !Marshal.IsComObject(value))
                {
                    return;
                }

                try
                {
                    Marshal.ReleaseComObject(value);
                }
                catch (Exception)
                {
                }
            }
        }
    }
}
