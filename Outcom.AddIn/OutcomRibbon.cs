using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outcom.AddIn
{
    [ComVisible(true)]
    public sealed class OutcomRibbon : Office.IRibbonExtensibility
    {
        private const string ExplorerRibbonId = "Microsoft.Outlook.Explorer";
        private const string ReadRibbonId = "Microsoft.Outlook.Mail.Read";
        private const string ComposeRibbonId = "Microsoft.Outlook.Mail.Compose";
        private const string RibbonResourceName = "Outcom.AddIn.OutcomRibbon.xml";
        private const string ReadRibbonResourceName =
            "Outcom.AddIn.OutcomReadRibbon.xml";
        private const string ComposeRibbonResourceName =
            "Outcom.AddIn.OutcomComposeRibbon.xml";
        private int operationInProgress;

        public string GetCustomUI(string ribbonId)
        {
            string resourceName;
            if (string.Equals(ribbonId, ExplorerRibbonId, StringComparison.Ordinal))
            {
                resourceName = RibbonResourceName;
            }
            else if (string.Equals(ribbonId, ReadRibbonId, StringComparison.Ordinal))
            {
                resourceName = ReadRibbonResourceName;
            }
            else if (string.Equals(ribbonId, ComposeRibbonId, StringComparison.Ordinal))
            {
                resourceName = ComposeRibbonResourceName;
            }
            else
            {
                return null;
            }

            Assembly assembly = typeof(OutcomRibbon).Assembly;
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException(
                        "La ressource Ribbon Outcom est introuvable : " + resourceName);
                }

                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        public void ProposeComposeReply(Office.IRibbonControl control)
        {
            Outlook.Inspector inspector = null;
            Outlook.Explorer explorer = null;
            Outlook.MailItem draft = null;
            object currentItem = null;
            object wordEditor = null;
            IWin32Window owner = null;
            try
            {
                object ribbonContext = control == null ? null : control.Context;
                inspector = ribbonContext as Outlook.Inspector;
                if (inspector != null)
                {
                    owner = GetOutlookWindowOwner(inspector);
                    currentItem = inspector.CurrentItem;
                    draft = currentItem as Outlook.MailItem;
                    if (draft == null)
                    {
                        throw new InvalidOperationException(
                            "Le message Outlook en cours de rédaction est introuvable.");
                    }

                    currentItem = null;
                    if (!inspector.IsWordMail())
                    {
                        throw new InvalidOperationException(
                            "L'éditeur Word d'Outlook est nécessaire pour analyser la zone de réponse.");
                    }

                    wordEditor = inspector.WordEditor;
                    if (wordEditor == null)
                    {
                        throw new InvalidOperationException(
                            "Outlook n'a pas rendu l'éditeur du message disponible.");
                    }

                    OutlookComposeContext inspectorContext =
                        OutlookComposeContext.Capture(draft, wordEditor);
                    ReleaseComObject(wordEditor);
                    wordEditor = null;
                    Globals.ThisAddIn.StartComposeReplyRequest(
                        inspector,
                        draft,
                        inspectorContext);
                    inspector = null;
                    draft = null;
                }
                else
                {
                    explorer = ribbonContext as Outlook.Explorer;
                    if (explorer == null)
                    {
                        ReleaseComObject(ribbonContext);
                        throw new InvalidOperationException(
                            "Aucune fenêtre Outlook compatible n'est active.");
                    }

                    owner = GetOutlookWindowOwner(explorer);
                    currentItem = explorer.ActiveInlineResponse;
                    draft = currentItem as Outlook.MailItem;
                    if (draft == null)
                    {
                        throw new InvalidOperationException(
                            "Commencez une réponse dans le volet de lecture, puis réessayez.");
                    }

                    currentItem = null;
                    wordEditor = explorer.ActiveInlineResponseWordEditor;
                    if (wordEditor == null)
                    {
                        throw new InvalidOperationException(
                            "L'éditeur de la réponse intégrée n'est pas disponible.");
                    }

                    OutlookComposeContext inlineContext =
                        OutlookComposeContext.Capture(draft, wordEditor);
                    ReleaseComObject(wordEditor);
                    wordEditor = null;
                    Globals.ThisAddIn.StartComposeReplyRequest(
                        explorer,
                        draft,
                        inlineContext);
                    explorer = null;
                    draft = null;
                }

                LocalLogger.Info("Proposition Codex lancée en arrière-plan.");
            }
            catch (Exception exception)
            {
                HandleRibbonOperationError(
                    owner ?? GetOutlookWindowOwner(inspector),
                    "proposition de réponse",
                    exception);
            }
            finally
            {
                ReleaseComObject(wordEditor);
                ReleaseComObject(currentItem);
                ReleaseComObject(draft);
                ReleaseComObject(inspector);
                ReleaseComObject(explorer);
            }
        }

        public void ReplyAndPropose(Office.IRibbonControl control)
        {
            ReplyToSelectedMessageAndPropose(control, false);
        }

        public void ReplyAllAndPropose(Office.IRibbonControl control)
        {
            ReplyToSelectedMessageAndPropose(control, true);
        }

        private static void ReplyToSelectedMessageAndPropose(
            Office.IRibbonControl control,
            bool replyAll)
        {
            Outlook.Explorer explorer = null;
            Outlook.Inspector sourceInspector = null;
            Outlook.Selection selection = null;
            Outlook.MailItem source = null;
            Outlook.MailItem draft = null;
            Outlook.Inspector draftInspector = null;
            object selectedItem = null;
            object wordEditor = null;
            IWin32Window owner = null;
            try
            {
                object ribbonContext = control == null ? null : control.Context;
                explorer = ribbonContext as Outlook.Explorer;
                if (explorer != null)
                {
                    owner = GetOutlookWindowOwner(explorer);
                    selection = explorer.Selection;
                    if (selection == null || selection.Count != 1)
                    {
                        throw new InvalidOperationException(
                            "Sélectionnez exactement un courrier Outlook, puis réessayez.");
                    }

                    selectedItem = selection[1];
                }
                else
                {
                    sourceInspector = ribbonContext as Outlook.Inspector;
                    if (sourceInspector == null)
                    {
                        ReleaseComObject(ribbonContext);
                        throw new InvalidOperationException(
                            "Aucun courrier Outlook n'est disponible.");
                    }

                    owner = GetOutlookWindowOwner(sourceInspector);
                    selectedItem = sourceInspector.CurrentItem;
                }

                source = selectedItem as Outlook.MailItem;
                if (source == null)
                {
                    throw new InvalidOperationException(
                        "L'élément sélectionné n'est pas un courrier Outlook.");
                }

                selectedItem = null;
                draft = replyAll ? source.ReplyAll() : source.Reply();
                if (draft == null)
                {
                    throw new InvalidOperationException(
                        "Outlook n'a pas pu créer le brouillon de réponse.");
                }

                draft.Display(false);
                draftInspector = draft.GetInspector;
                if (draftInspector == null)
                {
                    throw new InvalidOperationException(
                        "Outlook n'a pas rendu la fenêtre de rédaction disponible.");
                }

                if (!draftInspector.IsWordMail())
                {
                    throw new InvalidOperationException(
                        "L'éditeur Word d'Outlook est nécessaire pour analyser la zone de réponse.");
                }

                wordEditor = draftInspector.WordEditor;
                if (wordEditor == null)
                {
                    throw new InvalidOperationException(
                        "Outlook n'a pas rendu l'éditeur du brouillon disponible.");
                }

                OutlookComposeContext context = OutlookComposeContext.Capture(
                    draft,
                    wordEditor);
                ReleaseComObject(wordEditor);
                wordEditor = null;
                Globals.ThisAddIn.StartComposeReplyRequest(
                    draftInspector,
                    draft,
                    context);
                draftInspector = null;
                draft = null;

                LocalLogger.Info(
                    replyAll
                        ? "Réponse à tous créée et proposition Codex lancée en arrière-plan."
                        : "Réponse créée et proposition Codex lancée en arrière-plan.");
            }
            catch (Exception exception)
            {
                HandleRibbonOperationError(
                    owner ?? GetOutlookWindowOwner(explorer ?? (object)sourceInspector),
                    replyAll
                        ? "réponse à tous avec proposition"
                        : "réponse avec proposition",
                    exception);
            }
            finally
            {
                ReleaseComObject(wordEditor);
                ReleaseComObject(draftInspector);
                ReleaseComObject(draft);
                ReleaseComObject(source);
                ReleaseComObject(selectedItem);
                ReleaseComObject(selection);
                ReleaseComObject(sourceInspector);
                ReleaseComObject(explorer);
            }
        }

        public void ToggleCodexPane(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.ToggleCodexPane(control == null ? null : control.Context);
            }
            catch (Exception exception)
            {
                LocalLogger.Error(
                    "Impossible d’afficher le volet Codex (" +
                    exception.GetType().Name + ").");
                ShowMessage(
                    GetOutlookWindowOwner(),
                    exception is InvalidOperationException
                        ? exception.Message
                        : "Le volet Codex n'a pas pu être ouvert. Consultez le journal Outcom.",
                    MessageBoxIcon.Error);
            }
        }

        public void OpenCodexConnection(Office.IRibbonControl control)
        {
            try
            {
                IWin32Window owner = GetOutlookWindowOwner();
                using (var form = new CodexConnectionForm(Globals.ThisAddIn.GetCodexService()))
                {
                    if (owner == null)
                    {
                        form.ShowDialog();
                    }
                    else
                    {
                        form.ShowDialog(owner);
                    }
                }

                Globals.ThisAddIn.RefreshCodexPanesConnectionStatus();
            }
            catch (Exception exception)
            {
                LocalLogger.Error(
                    "Impossible d’ouvrir la gestion Codex (" + exception.GetType().Name + ").");

                MessageBox.Show(
                    "Impossible d’ouvrir la gestion de la connexion ChatGPT via Codex. " +
                    "Consultez le journal Outcom.",
                    "Outcom",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public void OpenCodexContext(Office.IRibbonControl control)
        {
            IWin32Window owner = GetOutlookWindowOwner();
            if (!TryBeginRibbonOperation(owner))
            {
                return;
            }

            try
            {
                CodexService service = Globals.ThisAddIn.GetCodexService();
                using (var form = new OutcomGlobalContextForm(
                    service.GetGlobalContext(),
                    service))
                {
                    DialogResult result = owner == null
                        ? form.ShowDialog()
                        : form.ShowDialog(owner);
                    if (result == DialogResult.OK)
                    {
                        service.SaveGlobalContext(form.ResultContext);
                    }
                }
            }
            catch (Exception exception)
            {
                HandleRibbonOperationError(owner, "configuration du contexte Codex", exception);
            }
            finally
            {
                EndRibbonOperation();
            }
        }

        public async void TestCodexConnection(Office.IRibbonControl control)
        {
            IWin32Window owner = GetOutlookWindowOwner();
            if (!TryBeginRibbonOperation(owner))
            {
                return;
            }

            try
            {
                CodexConnectionStatus status = await Globals.ThisAddIn
                    .GetCodexService()
                    .TestConnectionAsync(CancellationToken.None);

                bool connected = status != null &&
                                 status.Account != null &&
                                 status.Account.IsConnected;

                ShowMessage(
                    owner,
                    connected
                        ? "La connexion ChatGPT via Codex fonctionne."
                        : "Le profil Codex d’Outcom n’est pas connecté à ChatGPT.",
                    connected ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            }
            catch (Exception exception)
            {
                HandleRibbonOperationError(owner, "test de connexion", exception);
            }
            finally
            {
                EndRibbonOperation();
            }
        }

        public async void DisconnectCodex(Office.IRibbonControl control)
        {
            IWin32Window owner = GetOutlookWindowOwner();
            if (!TryBeginRibbonOperation(owner))
            {
                return;
            }

            try
            {
                DialogResult answer = ShowQuestion(
                    owner,
                    "Déconnecter le profil Codex isolé d’Outcom de ChatGPT ?");

                if (answer != DialogResult.Yes)
                {
                    return;
                }

                await Globals.ThisAddIn
                    .GetCodexService()
                    .LogoutAsync(CancellationToken.None);

                Globals.ThisAddIn.RefreshCodexPanesConnectionStatus();

                ShowMessage(
                    owner,
                    "Le profil Codex d’Outcom est déconnecté de ChatGPT.",
                    MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                HandleRibbonOperationError(owner, "déconnexion", exception);
            }
            finally
            {
                EndRibbonOperation();
            }
        }

        private bool TryBeginRibbonOperation(IWin32Window owner)
        {
            if (Interlocked.CompareExchange(ref operationInProgress, 1, 0) == 0)
            {
                return true;
            }

            ShowMessage(
                owner,
                "Une opération Codex est déjà en cours.",
                MessageBoxIcon.Information);
            return false;
        }

        private void EndRibbonOperation()
        {
            Interlocked.Exchange(ref operationInProgress, 0);
        }

        private static void HandleRibbonOperationError(
            IWin32Window owner,
            string operationName,
            Exception exception)
        {
            // Le message du protocole n’est volontairement pas journalisé.
            LocalLogger.Error(
                "Échec de l’opération Codex « " + operationName + " » (" +
                exception.GetType().Name + ").");

            string message = exception is CodexAppServerException || exception is InvalidOperationException
                ? exception.Message
                : "L’opération Codex a échoué. Consultez le journal Outcom.";
            ShowMessage(owner, message, MessageBoxIcon.Error);
        }

        private static void ShowMessage(IWin32Window owner, string message, MessageBoxIcon icon)
        {
            if (owner == null)
            {
                MessageBox.Show(message, "Outcom — Codex", MessageBoxButtons.OK, icon);
            }
            else
            {
                MessageBox.Show(owner, message, "Outcom — Codex", MessageBoxButtons.OK, icon);
            }
        }

        private static DialogResult ShowQuestion(IWin32Window owner, string message)
        {
            if (owner == null)
            {
                return MessageBox.Show(
                    message,
                    "Outcom — Codex",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);
            }

            return MessageBox.Show(
                owner,
                message,
                "Outcom — Codex",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
        }

        private static IWin32Window GetOutlookWindowOwner()
        {
            return GetOutlookWindowOwner(null);
        }

        private static IWin32Window GetOutlookWindowOwner(object preferredWindow)
        {
            IntPtr handle = IntPtr.Zero;
            object explorer = null;
            object inspector = null;

            try
            {
                handle = GetOleWindowHandle(preferredWindow);

                if (handle == IntPtr.Zero && preferredWindow == null)
                {
                    explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                    handle = GetOleWindowHandle(explorer);
                }

                if (handle == IntPtr.Zero)
                {
                    inspector = Globals.ThisAddIn.Application.ActiveInspector();
                    handle = GetOleWindowHandle(inspector);
                }

                if (handle == IntPtr.Zero && explorer == null)
                {
                    explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                    handle = GetOleWindowHandle(explorer);
                }
            }
            catch (COMException)
            {
                // Outlook peut changer de fenêtre pendant le clic ; le premier plan reste un repli sûr.
            }
            finally
            {
                ReleaseComObject(inspector);
                ReleaseComObject(explorer);
            }

            if (handle == IntPtr.Zero)
            {
                handle = GetForegroundWindow();
            }

            return handle == IntPtr.Zero ? null : new WindowOwner(handle);
        }

        private static Outlook.Inspector GetComposeInspector(Office.IRibbonControl control)
        {
            object context = control == null ? null : control.Context;
            var inspector = context as Outlook.Inspector;
            if (inspector != null)
            {
                return inspector;
            }

            ReleaseComObject(context);
            inspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (inspector == null)
            {
                throw new InvalidOperationException(
                    "Aucune fenêtre de rédaction Outlook n'est active.");
            }

            return inspector;
        }

        private static IntPtr GetOleWindowHandle(object outlookWindow)
        {
            var oleWindow = outlookWindow as IOleWindow;
            if (oleWindow == null)
            {
                return IntPtr.Zero;
            }

            IntPtr handle;
            return oleWindow.GetWindow(out handle) >= 0 ? handle : IntPtr.Zero;
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

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [ComImport]
        [Guid("00000114-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleWindow
        {
            [PreserveSig]
            int GetWindow(out IntPtr windowHandle);

            [PreserveSig]
            int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool enterMode);
        }

        private sealed class WindowOwner : IWin32Window
        {
            internal WindowOwner(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; private set; }
        }
    }
}
