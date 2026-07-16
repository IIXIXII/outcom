using System;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace Outcom.AddIn
{
    public partial class ThisAddIn
    {
        private readonly object codexServiceSyncRoot = new object();
        private readonly object codexTaskPaneSyncRoot = new object();
        private readonly object codexComposeRequestSyncRoot = new object();
        private readonly CodexOperationTracker codexOperationTracker =
            new CodexOperationTracker();
        private CodexService codexService;
        private CodexTaskPaneManager codexTaskPaneManager;
        private CodexComposeRequestManager codexComposeRequestManager;
        private int shutdownCompleted;

        internal CodexService GetCodexService()
        {
            lock (codexServiceSyncRoot)
            {
                if (Volatile.Read(ref shutdownCompleted) != 0)
                {
                    throw new ObjectDisposedException(
                        nameof(ThisAddIn),
                        "Outcom est en cours d’arrêt.");
                }

                if (codexService == null)
                {
                    codexService = new CodexService();
                }

                return codexService;
            }
        }

        internal void ToggleCodexPane(object outlookWindow)
        {
            CodexTaskPaneManager manager;
            lock (codexTaskPaneSyncRoot)
            {
                if (Volatile.Read(ref shutdownCompleted) != 0)
                {
                    throw new ObjectDisposedException(
                        nameof(ThisAddIn),
                        "Outcom est en cours d’arrêt.");
                }

                if (codexTaskPaneManager == null)
                {
                    codexTaskPaneManager = new CodexTaskPaneManager(this);
                }

                manager = codexTaskPaneManager;
            }

            manager.Toggle(outlookWindow);
        }

        internal void RefreshCodexPanesConnectionStatus()
        {
            CodexTaskPaneManager manager;
            lock (codexTaskPaneSyncRoot)
            {
                manager = codexTaskPaneManager;
            }

            if (manager != null)
            {
                manager.RefreshConnectionStatus();
            }
        }

        internal CodexOperationTracker GetCodexOperationTracker()
        {
            return codexOperationTracker;
        }

        internal void StartComposeReplyRequest(
            Outlook.Inspector inspector,
            Outlook.MailItem draft,
            OutlookComposeContext context)
        {
            CodexComposeRequestManager manager = GetComposeRequestManager();
            manager.Start(inspector, draft, context);
        }

        internal void StartComposeReplyRequest(
            Outlook.Explorer explorer,
            Outlook.MailItem draft,
            OutlookComposeContext context)
        {
            CodexComposeRequestManager manager = GetComposeRequestManager();
            manager.Start(explorer, draft, context);
        }

        private CodexComposeRequestManager GetComposeRequestManager()
        {
            lock (codexComposeRequestSyncRoot)
            {
                if (Volatile.Read(ref shutdownCompleted) != 0)
                {
                    throw new ObjectDisposedException(
                        nameof(ThisAddIn),
                        "Outcom est en cours d’arrêt.");
                }

                if (codexComposeRequestManager == null)
                {
                    codexComposeRequestManager = new CodexComposeRequestManager(
                        this,
                        codexOperationTracker);
                }

                return codexComposeRequestManager;
            }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            LocalLogger.Info("Démarrage du complément.");
            ((Outlook.ApplicationEvents_11_Event)this.Application).Quit += Application_Quit;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            ShutdownAddIn();
        }

        private void Application_Quit()
        {
            // Outlook ne déclenche pas toujours ThisAddIn_Shutdown lors de son arrêt.
            ShutdownAddIn();
        }

        private void ShutdownAddIn()
        {
            if (Interlocked.Exchange(ref shutdownCompleted, 1) != 0)
            {
                return;
            }

            CodexComposeRequestManager composeRequestManagerToDispose;
            lock (codexComposeRequestSyncRoot)
            {
                composeRequestManagerToDispose = codexComposeRequestManager;
                codexComposeRequestManager = null;
            }

            if (composeRequestManagerToDispose != null)
            {
                try
                {
                    composeRequestManagerToDispose.Dispose();
                }
                catch (Exception exception)
                {
                    LocalLogger.Error(
                        "Impossible d’arrêter les propositions Codex (" +
                        exception.GetType().Name + ").");
                }
            }

            CodexTaskPaneManager paneManagerToDispose;
            lock (codexTaskPaneSyncRoot)
            {
                paneManagerToDispose = codexTaskPaneManager;
                codexTaskPaneManager = null;
            }

            if (paneManagerToDispose != null)
            {
                try
                {
                    paneManagerToDispose.Dispose();
                }
                catch (Exception exception)
                {
                    LocalLogger.Error(
                        "Impossible d’arrêter les volets Codex (" +
                        exception.GetType().Name + ").");
                }
            }

            CodexService serviceToDispose;
            lock (codexServiceSyncRoot)
            {
                serviceToDispose = codexService;
                codexService = null;
            }

            if (serviceToDispose != null)
            {
                try
                {
                    serviceToDispose.Dispose();
                }
                catch (Exception exception)
                {
                    LocalLogger.Error(
                        "Impossible d’arrêter le service Codex (" +
                        exception.GetType().Name + ").");
                }
            }

            LocalLogger.Info("Arrêt du complément.");
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new OutcomRibbon();
        }

        #region Code généré par VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
