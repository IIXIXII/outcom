using System;
using System.Drawing;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Outcom.AddIn
{
    internal sealed class CodexConnectionForm : Form
    {
        private readonly CodexService service;
        private readonly CancellationTokenSource lifetimeCancellation = new CancellationTokenSource();
        private readonly Icon windowIcon;

        private readonly Label statusValue;
        private readonly TextBox accountValue;
        private readonly TextBox planValue;
        private readonly TextBox runtimePathValue;
        private readonly TextBox runtimeVersionValue;
        private readonly TextBox modelValue;
        private readonly TextBox limitsValue;
        private readonly TextBox profilePathValue;
        private readonly Button signInButton;
        private readonly Button testButton;
        private readonly Button logoutButton;
        private readonly Button closeButton;

        private CodexConnectionStatus currentStatus;
        private bool operationInProgress;
        private int lifetimeDisposed;

        internal CodexConnectionForm(CodexService service)
        {
            this.service = service ?? throw new ArgumentNullException(nameof(service));
            windowIcon = OutcomBranding.CreateWindowIcon();

            Text = "Connexion ChatGPT — Outcom";
            Icon = windowIcon;
            ShowIcon = true;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(1680, 930);
            MinimumSize = new Size(520, 380);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                ColumnCount = 1,
                RowCount = 4
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            var title = new Label
            {
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Text = "Connexion à ChatGPT avec Codex"
            };
            root.Controls.Add(title, 0, 0);

            var introduction = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(800, 0),
                Margin = new Padding(0, 8, 0, 14),
                Text = "Outcom utilise la session ChatGPT de l’utilisateur par l’intermédiaire de Codex app-server. " +
                       "Aucune clé API n’est demandée et aucun contenu Outlook n’est envoyé depuis cet écran."
            };
            root.Controls.Add(introduction, 0, 1);
            root.SizeChanged += (sender, args) =>
            {
                int availableWidth = Math.Max(
                    200,
                    root.ClientSize.Width - root.Padding.Horizontal);
                if (introduction.MaximumSize.Width != availableWidth)
                {
                    introduction.MaximumSize = new Size(availableWidth, 0);
                }
            };

            var details = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 8,
                AutoSize = false,
                AutoScroll = true
            };
            details.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            details.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            for (int row = 0; row < 7; row++)
            {
                details.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            details.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.Controls.Add(details, 0, 2);

            statusValue = new Label
            {
                AutoSize = true,
                Margin = new Padding(3, 7, 3, 8),
                Text = "Chargement…"
            };
            AddDetailRow(details, 0, "État", statusValue);

            accountValue = CreateReadOnlyValue();
            AddDetailRow(details, 1, "Compte", accountValue);

            planValue = CreateReadOnlyValue();
            AddDetailRow(details, 2, "Formule", planValue);

            runtimePathValue = CreateReadOnlyValue();
            AddDetailRow(details, 3, "Exécutable Codex", runtimePathValue);

            runtimeVersionValue = CreateReadOnlyValue();
            AddDetailRow(details, 4, "Version Codex", runtimeVersionValue);

            modelValue = CreateReadOnlyValue();
            AddDetailRow(details, 5, "Modèle par défaut", modelValue);

            profilePathValue = CreateReadOnlyValue();
            AddDetailRow(details, 6, "Profil isolé", profilePathValue);

            limitsValue = CreateReadOnlyValue();
            limitsValue.Multiline = true;
            limitsValue.ScrollBars = ScrollBars.Vertical;
            limitsValue.Dock = DockStyle.Fill;
            limitsValue.MinimumSize = new Size(0, 100);
            AddDetailRow(details, 7, "Limites Codex", limitsValue);

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = true,
                Margin = new Padding(0, 14, 0, 0)
            };
            root.Controls.Add(buttons, 0, 3);

            closeButton = new Button
            {
                AutoSize = true,
                DialogResult = DialogResult.Cancel,
                Text = "Fermer",
                TabIndex = 3
            };
            closeButton.Click += CloseButton_Click;
            buttons.Controls.Add(closeButton);

            logoutButton = new Button
            {
                AutoSize = true,
                Enabled = false,
                Text = "Se déconnecter",
                TabIndex = 2
            };
            logoutButton.Click += LogoutButton_Click;
            buttons.Controls.Add(logoutButton);

            testButton = new Button
            {
                AutoSize = true,
                Text = "Actualiser / tester",
                TabIndex = 1
            };
            testButton.Click += TestButton_Click;
            buttons.Controls.Add(testButton);

            signInButton = new Button
            {
                AutoSize = true,
                Text = "Se connecter avec ChatGPT",
                TabIndex = 0
            };
            signInButton.Click += SignInButton_Click;
            buttons.Controls.Add(signInButton);

            AcceptButton = signInButton;
            CancelButton = closeButton;

            Shown += CodexConnectionForm_Shown;
            FormClosing += CodexConnectionForm_FormClosing;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && Interlocked.Exchange(ref lifetimeDisposed, 1) == 0)
            {
                lifetimeCancellation.Cancel();
                lifetimeCancellation.Dispose();
                windowIcon.Dispose();
            }

            base.Dispose(disposing);
        }

        private static TextBox CreateReadOnlyValue()
        {
            return new TextBox
            {
                ReadOnly = true,
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Top,
                Margin = new Padding(3, 3, 3, 6)
            };
        }

        private static void AddDetailRow(
            TableLayoutPanel panel,
            int row,
            string labelText,
            Control valueControl)
        {
            var label = new Label
            {
                AutoSize = true,
                Margin = new Padding(3, 7, 3, 3),
                Text = labelText + " :"
            };

            panel.Controls.Add(label, 0, row);
            panel.Controls.Add(valueControl, 1, row);
        }

        private async void CodexConnectionForm_Shown(object sender, EventArgs e)
        {
            CodexConnectionStatus status = await RunStatusOperationAsync(
                "initialisation",
                cancellationToken => service.GetStatusAsync(false, cancellationToken),
                false);

            if (status != null &&
                status.Account != null &&
                status.Account.IsConnected &&
                IsUsable &&
                !lifetimeCancellation.IsCancellationRequested)
            {
                // Le statut détaillé fournit le modèle et les limites du compte déjà connecté.
                await RunStatusOperationAsync(
                    "actualisation",
                    service.TestConnectionAsync,
                    false);
            }
        }

        private async void SignInButton_Click(object sender, EventArgs e)
        {
            await RunStatusOperationAsync(
                "connexion",
                service.SignInWithChatGptAsync,
                true);
        }

        private async void TestButton_Click(object sender, EventArgs e)
        {
            CodexConnectionStatus status = await RunStatusOperationAsync(
                "test de connexion",
                service.TestConnectionAsync,
                true);

            if (status != null && IsUsable && status.Account != null && status.Account.IsConnected)
            {
                MessageBox.Show(
                    this,
                    "La connexion ChatGPT via Codex fonctionne.",
                    "Outcom",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private async void LogoutButton_Click(object sender, EventArgs e)
        {
            if (operationInProgress)
            {
                return;
            }

            DialogResult answer = MessageBox.Show(
                this,
                "Déconnecter le profil Codex isolé d’Outcom de ChatGPT ?",
                "Outcom",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);

            if (answer != DialogResult.Yes)
            {
                return;
            }

            await RunLogoutAsync();
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            lifetimeCancellation.Cancel();
            Close();
        }

        private void CodexConnectionForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            lifetimeCancellation.Cancel();
        }

        private async Task<CodexConnectionStatus> RunStatusOperationAsync(
            string operationName,
            Func<CancellationToken, Task<CodexConnectionStatus>> operation,
            bool showErrorDialog)
        {
            if (!TryBeginOperation())
            {
                return null;
            }

            try
            {
                CodexConnectionStatus status = await operation(lifetimeCancellation.Token);
                if (!IsUsable || lifetimeCancellation.IsCancellationRequested)
                {
                    return null;
                }

                ApplyStatus(status);
                return status;
            }
            catch (OperationCanceledException) when (lifetimeCancellation.IsCancellationRequested)
            {
                return null;
            }
            catch (Exception exception)
            {
                HandleOperationError(operationName, exception, showErrorDialog);
                return null;
            }
            finally
            {
                EndOperation();
            }
        }

        private async Task RunLogoutAsync()
        {
            if (!TryBeginOperation())
            {
                return;
            }

            try
            {
                await service.LogoutAsync(lifetimeCancellation.Token);
                CodexConnectionStatus status = await service.GetStatusAsync(false, lifetimeCancellation.Token);
                if (IsUsable && !lifetimeCancellation.IsCancellationRequested)
                {
                    ApplyStatus(status);
                }
            }
            catch (OperationCanceledException) when (lifetimeCancellation.IsCancellationRequested)
            {
            }
            catch (Exception exception)
            {
                HandleOperationError("déconnexion", exception, true);
            }
            finally
            {
                EndOperation();
            }
        }

        private bool TryBeginOperation()
        {
            if (operationInProgress || lifetimeCancellation.IsCancellationRequested || !IsUsable)
            {
                return false;
            }

            operationInProgress = true;
            statusValue.Text = "Opération en cours…";
            statusValue.ForeColor = SystemColors.ControlText;
            UpdateButtons();
            UseWaitCursor = true;
            return true;
        }

        private void EndOperation()
        {
            if (!IsUsable)
            {
                return;
            }

            operationInProgress = false;
            UseWaitCursor = false;
            UpdateButtons();
        }

        private void ApplyStatus(CodexConnectionStatus status)
        {
            currentStatus = status;

            CodexAccountInfo account = status == null ? null : status.Account;
            bool connected = account != null && account.IsConnected;

            statusValue.Text = connected
                ? "Connecté à ChatGPT"
                : account != null && account.RequiresOpenAiAuth
                    ? "Connexion ChatGPT requise"
                    : "Non connecté";
            statusValue.ForeColor = connected ? Color.DarkGreen : Color.DarkOrange;

            accountValue.Text = connected
                ? FirstNonEmpty(account.Email, account.AccountType, "Compte ChatGPT")
                : "—";
            planValue.Text = connected ? FirstNonEmpty(account.PlanType, "Non indiquée") : "—";
            runtimePathValue.Text = FirstNonEmpty(status == null ? null : status.ExecutablePath, "Introuvable");
            runtimeVersionValue.Text = FirstNonEmpty(status == null ? null : status.CliVersion, "Non disponible");
            modelValue.Text = FormatModel(status == null ? null : status.DefaultModel);
            profilePathValue.Text = FormatProfilePath(status);
            limitsValue.Text = FormatRateLimit(status == null ? null : status.RateLimit);

            UpdateButtons();
        }

        private void HandleOperationError(string operationName, Exception exception, bool showDialog)
        {
            // Ne jamais écrire le message d’exception : il peut provenir du protocole app-server.
            LocalLogger.Error(
                "Échec de l’opération Codex « " + operationName + " » (" +
                exception.GetType().Name + ").");

            if (!IsUsable)
            {
                return;
            }

            statusValue.Text = "Échec — " + GetUserFacingMessage(exception);
            statusValue.ForeColor = Color.DarkRed;

            if (showDialog)
            {
                MessageBox.Show(
                    this,
                    GetUserFacingMessage(exception),
                    "Outcom — Codex",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void UpdateButtons()
        {
            if (!IsUsable)
            {
                return;
            }

            bool connected = currentStatus != null &&
                             currentStatus.Account != null &&
                             currentStatus.Account.IsConnected;

            signInButton.Enabled = !operationInProgress && !connected;
            testButton.Enabled = !operationInProgress;
            logoutButton.Enabled = !operationInProgress && connected;
            closeButton.Enabled = true;
            AcceptButton = connected ? testButton : signInButton;
        }

        private static string FormatModel(CodexModelInfo model)
        {
            if (model == null)
            {
                return "Non disponible";
            }

            if (!string.IsNullOrWhiteSpace(model.DisplayName) &&
                !string.Equals(model.DisplayName, model.Id, StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(model.Id))
            {
                return model.DisplayName + " (" + model.Id + ")";
            }

            return FirstNonEmpty(model.DisplayName, model.Id, "Non disponible");
        }

        private static string FormatProfilePath(CodexConnectionStatus status)
        {
            if (status == null)
            {
                return "Non disponible";
            }

            if (!string.IsNullOrWhiteSpace(status.CodexHomePath) &&
                !string.IsNullOrWhiteSpace(status.WorkspacePath))
            {
                return status.CodexHomePath + "  |  espace : " + status.WorkspacePath;
            }

            return FirstNonEmpty(status.CodexHomePath, status.WorkspacePath, "Non disponible");
        }

        private static string FormatRateLimit(CodexRateLimitInfo rateLimit)
        {
            if (rateLimit == null)
            {
                return "Non disponibles";
            }

            var text = new StringBuilder();
            string name = FirstNonEmpty(rateLimit.LimitName, rateLimit.LimitId, "Limite du compte");
            text.Append(name);

            AppendRateLimitWindow(text, "principale", rateLimit.Primary);
            AppendRateLimitWindow(text, "secondaire", rateLimit.Secondary);

            if (!string.IsNullOrWhiteSpace(rateLimit.RateLimitReachedType))
            {
                text.AppendLine();
                text.Append("Limite atteinte : ");
                text.Append(rateLimit.RateLimitReachedType);
            }

            return text.ToString();
        }

        private static void AppendRateLimitWindow(
            StringBuilder text,
            string label,
            CodexRateLimitWindow window)
        {
            if (window == null)
            {
                return;
            }

            text.AppendLine();
            text.Append(char.ToUpperInvariant(label[0]));
            text.Append(label.Substring(1));
            text.Append(" : ");
            text.Append(window.UsedPercent);
            text.Append(" % utilisés");

            if (window.ResetsAt.HasValue)
            {
                text.Append(", réinitialisation ");
                text.Append(window.ResetsAt.Value.ToLocalTime().ToString("g"));
            }
            else if (window.WindowDurationMinutes.HasValue)
            {
                text.Append(" (fenêtre de ");
                text.Append(window.WindowDurationMinutes.Value);
                text.Append(" min)");
            }
        }

        private static string GetUserFacingMessage(Exception exception)
        {
            if (exception is CodexAppServerException && !string.IsNullOrWhiteSpace(exception.Message))
            {
                return exception.Message;
            }

            if (exception is InvalidOperationException && !string.IsNullOrWhiteSpace(exception.Message))
            {
                return exception.Message;
            }

            return "L’opération Codex a échoué. Consultez le journal Outcom.";
        }

        private static string FirstNonEmpty(params string[] values)
        {
            foreach (string value in values)
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }

            return string.Empty;
        }

        private bool IsUsable
        {
            get { return !IsDisposed && !Disposing && IsHandleCreated; }
        }
    }
}
