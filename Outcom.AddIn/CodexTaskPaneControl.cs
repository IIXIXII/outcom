using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outcom.AddIn
{
    /// <summary>
    /// Interface de conversations natives hébergée dans un CustomTaskPane Outlook.
    /// Le contrôle n'expose jamais le modèle objet Outlook à Codex.
    /// </summary>
    [System.ComponentModel.DesignerCategory("Code")]
    internal sealed class CodexTaskPaneControl : UserControl
    {
        private const int MaximumExtractedContextCharacters = 160000;
        private readonly ThisAddIn _addIn;
        private readonly Outlook.Explorer _explorer;
        private readonly CodexService _service;
        private readonly CodexOperationTracker _operationTracker;
        private readonly Image _brandImage;
        private readonly CancellationTokenSource _lifetimeCancellation =
            new CancellationTokenSource();
        private readonly Label _connectionStatusLabel;
        private readonly Label _globalContextStatusLabel;
        private readonly FlowLayoutPanel _contextDocumentsPanel;
        private readonly Label _contextDropHintLabel;
        private readonly Label _operationStatusLabel;
        private readonly Label _activitySummaryLabel;
        private readonly Label _conversationSummaryLabel;
        private readonly ListBox _conversationList;
        private readonly ListBox _activityList;
        private readonly Button _cancelTrackedOperationButton;
        private readonly Button _clearCompletedOperationsButton;
        private readonly Button _newConversationButton;
        private readonly Button _closeConversationButton;
        private readonly Button _copyButton;
        private readonly Button _draftButton;
        private readonly Button _cancelButton;
        private readonly Button _sendButton;
        private readonly RichTextBox _transcript;
        private readonly TextBox _composer;

        private readonly List<ConversationWorkspace> _workspaces =
            new List<ConversationWorkspace>();
        private ConversationWorkspace _activeWorkspace;
        private bool _isConnected;
        private bool _loadStarted;
        private int _activeTrackedOperationCount;
        private int _conversationSequence;
        private int _disposeState;

        internal CodexTaskPaneControl(ThisAddIn addIn, Outlook.Explorer explorer)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
            _explorer = explorer ?? throw new ArgumentNullException(nameof(explorer));
            _service = _addIn.GetCodexService();
            _operationTracker = _addIn.GetCodexOperationTracker();
            _brandImage = OutcomBranding.CreateHeaderImage();

            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = SystemColors.Control;
            ForeColor = SystemColors.ControlText;
            MinimumSize = new Size(300, 300);
            Padding = new Padding(10);

            var root = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 3,
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                Padding = Padding.Empty,
                AutoScroll = true
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var header = new TableLayoutPanel
            {
                ColumnCount = 3,
                RowCount = 1,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 8)
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            var brandPicture = new PictureBox
            {
                Image = _brandImage,
                Size = new Size(24, 24),
                SizeMode = PictureBoxSizeMode.Zoom,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 2, 7, 0),
                TabStop = false,
                AccessibleName = "Icône générale Outcom"
            };

            var titleLabel = new Label
            {
                Text = "Codex",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 6, 10, 0),
                AccessibleName = "Codex"
            };
            _connectionStatusLabel = new Label
            {
                Text = "Non vérifié",
                AutoEllipsis = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 3, 6, 0),
                AccessibleName = "État de la connexion Codex"
            };
            header.Controls.Add(brandPicture, 0, 0);
            header.Controls.Add(titleLabel, 1, 0);
            header.Controls.Add(_connectionStatusLabel, 2, 0);

            var globalContextPanel = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 1,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 8)
            };
            globalContextPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _globalContextStatusLabel = new Label
            {
                Text = "Contexte Codex : non défini",
                AutoEllipsis = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 3, 8, 0),
                AccessibleName = "État du contexte Codex"
            };
            globalContextPanel.Controls.Add(_globalContextStatusLabel, 0, 0);

            var activityPanel = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 4,
                Dock = DockStyle.Fill,
                AutoSize = false,
                Padding = new Padding(8),
                Margin = Padding.Empty,
                BackColor = SystemColors.ControlLight
            };
            activityPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            activityPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            activityPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            activityPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            activityPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            _activitySummaryLabel = new Label
            {
                Text = "Activité Codex — aucune demande en cours",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Anchor = AnchorStyles.Left,
                AccessibleName = "Résumé de l'activité Codex"
            };

            _conversationSummaryLabel = new Label
            {
                Text = "Conversations — 1 conversation",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 0),
                AccessibleName = "Liste des conversations Codex"
            };

            _newConversationButton = CreateButton(
                "&Nouvelle conversation",
                NewConversationButton_Click);
            _newConversationButton.AccessibleDescription =
                "Créer une conversation indépendante sans effacer les conversations existantes.";
            _closeConversationButton = CreateButton(
                "&Fermer",
                CloseConversationButton_Click);
            _closeConversationButton.AccessibleDescription =
                "Fermer la conversation sélectionnée lorsqu'aucune réponse n'est en cours.";
            var conversationButtons = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Margin = new Padding(0, 5, 0, 5)
            };
            conversationButtons.Controls.Add(_newConversationButton);
            conversationButtons.Controls.Add(_closeConversationButton);
            _conversationList = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                HorizontalScrollbar = true,
                AccessibleName = "Conversations Codex"
            };
            _conversationList.SelectedIndexChanged +=
                ConversationList_SelectedIndexChanged;

            var conversationListPanel = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 3,
                Dock = DockStyle.Fill,
                AutoSize = false,
                Padding = new Padding(8),
                Margin = Padding.Empty,
                BackColor = SystemColors.ControlLight
            };
            conversationListPanel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100F));
            conversationListPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            conversationListPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            conversationListPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            conversationListPanel.Controls.Add(_conversationSummaryLabel, 0, 0);
            conversationListPanel.Controls.Add(conversationButtons, 0, 1);
            conversationListPanel.Controls.Add(_conversationList, 0, 2);

            var operationListTitle = new Label
            {
                Text = "Activités en cours et récentes",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Margin = new Padding(0, 6, 0, 0),
                AccessibleName = "Traitements Codex"
            };
            _cancelTrackedOperationButton = CreateButton(
                "Annuler la &demande",
                CancelTrackedOperationButton_Click);
            _cancelTrackedOperationButton.Enabled = false;
            _cancelTrackedOperationButton.AccessibleDescription =
                "Annuler la demande Codex sélectionnée si elle est encore en cours.";
            _clearCompletedOperationsButton = CreateButton(
                "Effacer les tâches &terminées",
                ClearCompletedOperationsButton_Click);
            _clearCompletedOperationsButton.Enabled = false;
            _clearCompletedOperationsButton.AccessibleDescription =
                "Supprimer de la liste toutes les demandes Codex terminées.";
            var activityButtons = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = true,
                Margin = new Padding(0, 5, 0, 5)
            };
            activityButtons.Controls.Add(_cancelTrackedOperationButton);
            activityButtons.Controls.Add(_clearCompletedOperationsButton);
            _activityList = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                HorizontalScrollbar = true,
                AccessibleName = "Demandes Codex en cours et récentes"
            };
            _activityList.SelectedIndexChanged += ActivityList_SelectedIndexChanged;
            activityPanel.Controls.Add(_activitySummaryLabel, 0, 0);
            activityPanel.Controls.Add(operationListTitle, 0, 1);
            activityPanel.Controls.Add(activityButtons, 0, 2);
            activityPanel.Controls.Add(_activityList, 0, 3);

            var contextPanel = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 2,
                Dock = DockStyle.Fill,
                AutoSize = false,
                Padding = new Padding(8, 4, 8, 4),
                Margin = new Padding(0, 0, 0, 4),
                BackColor = SystemColors.ControlLight
            };
            contextPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            contextPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            contextPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var contextTitle = new Label
            {
                Text = "Documents joints à la conversation",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 3)
            };
            _contextDocumentsPanel = new FlowLayoutPanel
            {
                AllowDrop = true,
                AutoScroll = true,
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(4),
                Margin = Padding.Empty,
                AccessibleName = "Documents de contexte de la conversation",
                AccessibleDescription =
                    "Déposez ici des fichiers Windows, des courriers ou des pièces jointes Outlook."
            };
            _contextDocumentsPanel.DragEnter += ContextDocumentsPanel_DragEnter;
            _contextDocumentsPanel.DragOver += ContextDocumentsPanel_DragEnter;
            _contextDocumentsPanel.DragLeave += ContextDocumentsPanel_DragLeave;
            _contextDocumentsPanel.DragDrop += ContextDocumentsPanel_DragDrop;
            _contextDropHintLabel = new Label
            {
                Text = "Déposez ici des fichiers ou des éléments Outlook",
                AutoSize = false,
                Size = new Size(420, 40),
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = SystemColors.GrayText,
                BackColor = SystemColors.Window,
                Margin = new Padding(2),
                AccessibleName = "Zone de dépôt de documents"
            };
            AttachDocumentDropHandlers(_contextDropHintLabel);
            _contextDocumentsPanel.Controls.Add(_contextDropHintLabel);
            contextPanel.Controls.Add(contextTitle, 0, 0);
            contextPanel.Controls.Add(_contextDocumentsPanel, 0, 1);

            _transcript = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                DetectUrls = true,
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.WindowText,
                BorderStyle = BorderStyle.FixedSingle,
                HideSelection = false,
                AccessibleName = "Conversation avec Codex",
                AccessibleDescription =
                    "Historique sélectionnable de la conversation. Le markdown est affiché en texte brut.",
                Margin = Padding.Empty,
                MinimumSize = new Size(0, 60)
            };

            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true,
                Margin = new Padding(0, 0, 0, 8)
            };
            _copyButton = CreateButton("&Copier", CopyButton_Click);
            _copyButton.Enabled = false;
            _draftButton = CreateButton("Créer un &brouillon", DraftButton_Click);
            _draftButton.Enabled = false;
            _draftButton.AccessibleDescription =
                "Créer et ouvrir un brouillon Outlook en texte brut, sans destinataire et sans l'envoyer.";
            actions.Controls.Add(_copyButton);
            actions.Controls.Add(_draftButton);

            var composerPanel = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 2,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Margin = Padding.Empty
            };
            composerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            composerPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 80F));
            composerPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _composer = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                AcceptsReturn = true,
                ScrollBars = ScrollBars.Vertical,
                Enabled = false,
                AccessibleName = "Demande à Codex",
                AccessibleDescription = "Saisissez une demande. Ctrl+Entrée pour l'envoyer."
            };
            _composer.KeyDown += Composer_KeyDown;
            _composer.TextChanged += Composer_TextChanged;

            var sendRow = new TableLayoutPanel
            {
                ColumnCount = 3,
                RowCount = 1,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Margin = new Padding(0, 5, 0, 0)
            };
            sendRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            sendRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            sendRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _operationStatusLabel = new Label
            {
                Text = "Utilisez Configurer Codex dans le ruban pour vous connecter.",
                AutoEllipsis = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AccessibleName = "État de la conversation"
            };
            _cancelButton = CreateButton("&Annuler", CancelButton_Click);
            _cancelButton.Enabled = false;
            _sendButton = CreateButton("&Envoyer", SendButton_Click);
            _sendButton.Enabled = false;
            sendRow.Controls.Add(_operationStatusLabel, 0, 0);
            sendRow.Controls.Add(_cancelButton, 1, 0);
            sendRow.Controls.Add(_sendButton, 2, 0);
            composerPanel.Controls.Add(_composer, 0, 0);
            composerPanel.Controls.Add(sendRow, 0, 1);

            var conversationEditorPanel = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 3,
                Dock = DockStyle.Fill,
                Padding = Padding.Empty,
                Margin = Padding.Empty
            };
            conversationEditorPanel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100F));
            conversationEditorPanel.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100F));
            conversationEditorPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            conversationEditorPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            conversationEditorPanel.Controls.Add(_transcript, 0, 0);
            conversationEditorPanel.Controls.Add(actions, 0, 1);
            conversationEditorPanel.Controls.Add(composerPanel, 0, 2);

            var currentConversationPanel = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 2,
                Dock = DockStyle.Fill,
                Padding = Padding.Empty,
                Margin = Padding.Empty
            };
            currentConversationPanel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100F));
            currentConversationPanel.RowStyles.Add(
                new RowStyle(SizeType.Percent, 27.5F));
            currentConversationPanel.RowStyles.Add(
                new RowStyle(SizeType.Percent, 72.5F));
            currentConversationPanel.Controls.Add(contextPanel, 0, 0);
            currentConversationPanel.Controls.Add(conversationEditorPanel, 0, 1);

            var conversationsPanel = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 2,
                Dock = DockStyle.Fill,
                Padding = Padding.Empty,
                Margin = Padding.Empty
            };
            conversationsPanel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100F));
            // 20 % et 40 % de la zone utile correspondent à 1/3 et 2/3
            // de la grande partie Conversations, qui occupe elle-même 60 %.
            conversationsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 33.333F));
            conversationsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 66.667F));
            conversationsPanel.Controls.Add(conversationListPanel, 0, 0);
            conversationsPanel.Controls.Add(currentConversationPanel, 0, 1);

            var activityConversationsSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                BorderStyle = BorderStyle.FixedSingle,
                SplitterWidth = 7,
                IsSplitterFixed = false,
                FixedPanel = FixedPanel.None,
                Margin = Padding.Empty,
                MinimumSize = new Size(0, 520),
                Size = new Size(300, 700),
                Panel1MinSize = 180,
                Panel2MinSize = 320,
                SplitterDistance = 280,
                AccessibleName =
                    "Répartition entre les activités Codex et les conversations"
            };
            activityConversationsSplit.Panel1.Controls.Add(activityPanel);
            activityConversationsSplit.Panel2.Controls.Add(conversationsPanel);

            root.Controls.Add(header, 0, 0);
            root.Controls.Add(globalContextPanel, 0, 1);
            root.Controls.Add(activityConversationsSplit, 0, 2);
            Controls.Add(root);

            _service.GlobalContextChanged += CodexService_GlobalContextChanged;
            _operationTracker.Changed += CodexOperationTracker_Changed;
            RefreshGlobalContextDisplay(_service.GetGlobalContext());
            RefreshActivityDisplay(_operationTracker.GetSnapshot());
            CreateConversationWorkspace(true);
            Load += CodexTaskPaneControl_Load;
        }

        internal void CancelActiveTurn()
        {
            foreach (ConversationWorkspace workspace in
                new List<ConversationWorkspace>(_workspaces))
            {
                CancelWorkspaceTurn(workspace);
            }
        }

        private ConversationWorkspace CreateConversationWorkspace(bool select)
        {
            var workspace = new ConversationWorkspace(
                ++_conversationSequence,
                "Conversation " + _conversationSequence);
            _workspaces.Add(workspace);
            _conversationList.Items.Add(workspace);
            if (select || _activeWorkspace == null)
            {
                _conversationList.SelectedItem = workspace;
            }

            RefreshConversationSummary();
            return workspace;
        }

        private void ConversationList_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selected = _conversationList.SelectedItem as ConversationWorkspace;
            if (selected == null || ReferenceEquals(selected, _activeWorkspace))
            {
                return;
            }

            _activeWorkspace = selected;
            _transcript.Text = selected.Transcript.ToString();
            _transcript.SelectionStart = _transcript.TextLength;
            _transcript.ScrollToCaret();
            _composer.Text = selected.ComposerText ?? string.Empty;
            _composer.SelectionStart = _composer.TextLength;
            RefreshContextDocumentsDisplay();
            RefreshActiveWorkspaceControls();
        }

        private void Composer_TextChanged(object sender, EventArgs e)
        {
            if (_activeWorkspace != null)
            {
                _activeWorkspace.ComposerText = _composer.Text;
            }
        }

        private void CloseConversationButton_Click(object sender, EventArgs e)
        {
            ConversationWorkspace workspace = _activeWorkspace;
            if (workspace == null || workspace.IsBusy)
            {
                return;
            }

            if ((HasConversationContent(workspace) ||
                workspace.ContextDocuments.Count > 0) && MessageBox.Show(
                this,
                "Fermer cette conversation et oublier son contexte Codex éphémère ?",
                "Outcom — Codex",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2) != DialogResult.Yes)
            {
                return;
            }

            InvalidateWorkspaceConversation(workspace);
            DisposeContextDocuments(workspace);
            int index = _workspaces.IndexOf(workspace);
            _workspaces.Remove(workspace);
            _conversationList.Items.Remove(workspace);
            _activeWorkspace = null;
            if (_workspaces.Count == 0)
            {
                CreateConversationWorkspace(true);
            }
            else
            {
                _conversationList.SelectedItem =
                    _workspaces[Math.Min(index, _workspaces.Count - 1)];
            }

            RefreshConversationSummary();
        }

        private static void CancelWorkspaceTurn(ConversationWorkspace workspace)
        {
            CancellationTokenSource cancellation = workspace == null
                ? null
                : workspace.TurnCancellation;
            if (cancellation != null && !cancellation.IsCancellationRequested)
            {
                cancellation.Cancel();
            }
        }

        private void RefreshConversationSummary()
        {
            _conversationList.Refresh();
            _closeConversationButton.Enabled = _activeWorkspace != null &&
                !_activeWorkspace.IsBusy;
            string conversationSummary = _workspaces.Count +
                (_workspaces.Count == 1 ? " conversation" : " conversations");
            string activitySummary = _activeTrackedOperationCount == 0
                ? "aucun traitement en cours"
                : _activeTrackedOperationCount +
                    (_activeTrackedOperationCount == 1
                        ? " traitement en cours"
                        : " traitements en cours");
            _conversationSummaryLabel.Text = "Conversations — " + conversationSummary;
            _activitySummaryLabel.Text = "Activités Codex — " + activitySummary;
        }

        internal async void RefreshConnectionStatus()
        {
            if (IsDisposed || Disposing)
            {
                return;
            }

            await RefreshConnectionStatusAsync();
        }

        private void CodexOperationTracker_Changed(
            object sender,
            CodexOperationChangedEventArgs e)
        {
            if (IsDisposed || Disposing)
            {
                return;
            }

            if (InvokeRequired)
            {
                try
                {
                    BeginInvoke(new Action(
                        () => RefreshActivityDisplay(_operationTracker.GetSnapshot())));
                }
                catch (InvalidOperationException)
                {
                }

                return;
            }

            RefreshActivityDisplay(_operationTracker.GetSnapshot());
        }

        private void RefreshActivityDisplay(IReadOnlyList<CodexOperationInfo> operations)
        {
            Guid selectedId = Guid.Empty;
            var selected = _activityList.SelectedItem as CodexOperationInfo;
            if (selected != null)
            {
                selectedId = selected.Id;
            }

            int activeCount = 0;
            int completedCount = 0;
            int selectedIndex = -1;
            _activityList.BeginUpdate();
            try
            {
                _activityList.Items.Clear();
                if (operations != null)
                {
                    foreach (CodexOperationInfo operation in operations)
                    {
                        if (operation.IsActive)
                        {
                            activeCount++;
                        }
                        else
                        {
                            completedCount++;
                        }

                        int index = _activityList.Items.Add(operation);
                        if (operation.Id == selectedId)
                        {
                            selectedIndex = index;
                        }
                    }
                }

                if (selectedIndex < 0 && _activityList.Items.Count > 0)
                {
                    selectedIndex = 0;
                }

                _activityList.SelectedIndex = selectedIndex;
            }
            finally
            {
                _activityList.EndUpdate();
            }

            _activeTrackedOperationCount = activeCount;
            RefreshConversationSummary();
            _clearCompletedOperationsButton.Enabled = completedCount > 0;
            UpdateTrackedOperationCancelButton();
        }

        private void ActivityList_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateTrackedOperationCancelButton();
        }

        private void UpdateTrackedOperationCancelButton()
        {
            var operation = _activityList.SelectedItem as CodexOperationInfo;
            _cancelTrackedOperationButton.Enabled = operation != null &&
                operation.State == CodexOperationState.Running &&
                operation.CanCancel;
        }

        private void CancelTrackedOperationButton_Click(object sender, EventArgs e)
        {
            var operation = _activityList.SelectedItem as CodexOperationInfo;
            if (operation != null)
            {
                _operationTracker.RequestCancellation(operation.Id);
            }
        }

        private void ClearCompletedOperationsButton_Click(object sender, EventArgs e)
        {
            _operationTracker.ClearCompleted();
        }

        private static Button CreateButton(string text, EventHandler clickHandler)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                UseVisualStyleBackColor = true,
                Margin = new Padding(0, 0, 5, 0)
            };
            button.Click += clickHandler;
            return button;
        }

        private async void CodexTaskPaneControl_Load(object sender, EventArgs e)
        {
            if (_loadStarted)
            {
                return;
            }

            _loadStarted = true;
            await RefreshConnectionStatusAsync();
            if (_isConnected && !IsDisposed)
            {
                _composer.Focus();
            }
        }

        private async Task RefreshConnectionStatusAsync()
        {
            SetConnectionState(false, "Vérification…", "Vérification de la connexion…");
            try
            {
                CodexConnectionStatus status = await _addIn
                    .GetCodexService()
                    .GetStatusAsync(false, _lifetimeCancellation.Token);
                bool connected = status != null &&
                    status.Account != null &&
                    status.Account.IsConnected &&
                    string.Equals(status.Account.AccountType, "chatgpt", StringComparison.Ordinal);

                SetConnectionState(
                    connected,
                    connected ? "Connecté" : "Non connecté",
                    connected
                        ? "Prêt — Ctrl+Entrée pour envoyer."
                        : "Utilisez Configurer Codex dans le ruban pour vous connecter.");
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception exception)
            {
                LocalLogger.Error(
                    "Impossible de vérifier Codex depuis le volet (" +
                    exception.GetType().Name + ").");
                SetConnectionState(false, "Indisponible", GetUserFacingError(exception));
            }
        }

        private void CodexService_GlobalContextChanged(
            object sender,
            OutcomGlobalContextChangedEventArgs e)
        {
            if (IsDisposed || Disposing)
            {
                return;
            }

            if (InvokeRequired)
            {
                try
                {
                    BeginInvoke(new Action(
                        () => ApplyGlobalContextChange(e == null ? null : e.Context)));
                }
                catch (InvalidOperationException)
                {
                }

                return;
            }

            ApplyGlobalContextChange(e == null ? null : e.Context);
        }

        private void ApplyGlobalContextChange(OutcomGlobalContext context)
        {
            if (IsDisposed || Disposing)
            {
                return;
            }

            RefreshGlobalContextDisplay(context);
            bool resetAnyWorkspace = false;
            foreach (ConversationWorkspace workspace in
                new List<ConversationWorkspace>(_workspaces))
            {
                if (!HasConversationContent(workspace) && !workspace.IsBusy)
                {
                    continue;
                }

                resetAnyWorkspace = true;
                CancelWorkspaceTurn(workspace);
                ResetConversation(workspace, true);
                AppendSystemNote(
                    workspace,
                    "Le contexte Codex a été modifié. La prochaine demande démarrera un " +
                    "nouveau fil avec ces réglages.");
                SetOperationStatus(
                    workspace,
                    "Contexte Codex mis à jour — prochaine demande dans un nouveau fil.");
            }

            if (resetAnyWorkspace)
            {
                RefreshActiveWorkspaceDisplay();
            }

            if (!resetAnyWorkspace)
            {
                SetOperationStatus(
                    _activeWorkspace,
                    "Contexte Codex mis à jour — il s'appliquera à la prochaine demande.");
            }
        }

        private void RefreshGlobalContextDisplay(OutcomGlobalContext context)
        {
            OutcomGlobalContext current = context ?? new OutcomGlobalContext();
            int sectionCount = current.SectionCount;
            string model = string.IsNullOrWhiteSpace(current.ModelId)
                ? "modèle par défaut"
                : current.ModelId;
            string effort = string.IsNullOrWhiteSpace(current.ReasoningEffort)
                ? "raisonnement par défaut"
                : "raisonnement " + new CodexReasoningEffortInfo
                {
                    Value = current.ReasoningEffort
                };
            string directives = sectionCount == 0
                ? "aucune directive"
                : sectionCount + (sectionCount == 1 ? " directive" : " directives");
            _globalContextStatusLabel.Text =
                "Contexte Codex : " + model + " — " + effort + " — " + directives;
            _globalContextStatusLabel.ForeColor = current.IsEmpty
                ? SystemColors.GrayText
                : Color.DarkGreen;
        }

        private void ContextDocumentsPanel_DragEnter(object sender, DragEventArgs e)
        {
            bool supported = e.Data != null &&
                (e.Data.GetDataPresent(DataFormats.FileDrop) ||
                    OutlookVirtualFileDrop.IsSupported(e.Data));
            if (!supported || _activeWorkspace == null || _activeWorkspace.IsBusy)
            {
                e.Effect = DragDropEffects.None;
                return;
            }

            e.Effect = DragDropEffects.Copy;
            _contextDocumentsPanel.BackColor = SystemColors.Info;
        }

        private void ContextDocumentsPanel_DragLeave(object sender, EventArgs e)
        {
            _contextDocumentsPanel.BackColor = SystemColors.Window;
        }

        private void ContextDocumentsPanel_DragDrop(object sender, DragEventArgs e)
        {
            _contextDocumentsPanel.BackColor = SystemColors.Window;
            ConversationWorkspace workspace = _activeWorkspace;
            if (workspace == null || workspace.IsBusy || e.Data == null)
            {
                return;
            }

            if (workspace.ContextDocumentsSent &&
                HasConversationContent(workspace) && !ConfirmConversationReset(
                "Les documents actuels font déjà partie du contexte Codex. Modifier la " +
                "liste redémarre cette conversation. Continuer ?"))
            {
                return;
            }

            try
            {
                if (workspace.ContextDocumentsSent)
                {
                    ResetConversation(workspace, false);
                }

                var paths = new List<string>();
                bool temporary = false;
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var dropped = e.Data.GetData(DataFormats.FileDrop) as string[];
                    if (dropped != null)
                    {
                        paths.AddRange(dropped.Where(File.Exists));
                    }
                }
                else if (OutlookVirtualFileDrop.IsSupported(e.Data))
                {
                    try
                    {
                        paths.AddRange(OutlookVirtualFileDrop.Extract(e.Data));
                    }
                    catch (Exception extractionException)
                    {
                        LocalLogger.Info(
                            "Le flux virtuel Outlook n'a pas pu être matérialisé ; " +
                            "utilisation de la sélection Outlook (" +
                            extractionException.GetType().Name + ").");
                    }

                    if (paths.Count == 0)
                    {
                        paths.AddRange(
                            OutlookMailContext.SaveSelectionAsNativeMessages(_explorer));
                    }

                    temporary = true;
                }

                AddContextDocuments(workspace, paths, temporary);
            }
            catch (Exception exception)
            {
                LocalLogger.Error(
                    "Impossible d'ajouter les documents déposés (" +
                    exception.GetType().Name + ").");
                ShowError(GetUserFacingError(exception));
            }
        }

        private void AddContextDocuments(
            ConversationWorkspace workspace,
            IEnumerable<string> paths,
            bool temporary)
        {
            const int maximumDocumentCount = 10;
            var existingPaths = new HashSet<string>(
                workspace.ContextDocuments
                    .Where(item => !string.IsNullOrWhiteSpace(item.FilePath))
                    .Select(item => item.FilePath),
                StringComparer.OrdinalIgnoreCase);
            var candidates = (paths ?? Enumerable.Empty<string>())
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(path => !existingPaths.Contains(Path.GetFullPath(path)))
                .ToList();
            if (candidates.Count == 0)
            {
                throw new InvalidOperationException(
                    "Aucun nouveau fichier exploitable n'a été déposé.");
            }

            if (workspace.ContextDocuments.Count + candidates.Count > maximumDocumentCount)
            {
                if (temporary)
                {
                    DeleteTemporaryFiles(candidates);
                }

                throw new InvalidOperationException(
                    "Une conversation peut contenir au maximum " +
                    maximumDocumentCount + " documents.");
            }

            int added = 0;
            int separatedAttachmentCount = 0;
            int skippedAttachmentCount = 0;
            bool attachmentExtractionFailed = false;
            int remainingAttachmentSlots = maximumDocumentCount -
                workspace.ContextDocuments.Count - candidates.Count;
            var pending = new List<ConversationContextDocument>();
            try
            {
                foreach (string path in candidates)
                {
                    ConversationContextDocument document;
                    if (temporary && string.Equals(
                        Path.GetExtension(path),
                        ".msg",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        OutlookMailContext capturedMail = null;
                        try { capturedMail = OutlookMailContext.Capture(_explorer); }
                        catch (InvalidOperationException) { }
                        document = ConversationContextDocument.FromFile(
                            path,
                            true,
                            capturedMail);
                    }
                    else
                    {
                        document = ConversationContextDocument.FromFile(path, temporary);
                    }

                    pending.Add(document);

                    string extension = Path.GetExtension(path) ?? string.Empty;
                    if (extension.Equals(".msg", StringComparison.OrdinalIgnoreCase) ||
                        extension.Equals(".eml", StringComparison.OrdinalIgnoreCase))
                    {
                        EmailAttachmentExtractionResult attachments =
                            DocumentContextExtractor.ExtractEmailAttachments(
                                path,
                                remainingAttachmentSlots,
                                CancellationToken.None);
                        attachmentExtractionFailed |= attachments.HadError;
                        skippedAttachmentCount += attachments.SkippedCount;
                        foreach (string attachmentPath in attachments.Paths)
                        {
                            pending.Add(ConversationContextDocument.FromFile(
                                attachmentPath,
                                true,
                                null,
                                document.DisplayName));
                            separatedAttachmentCount++;
                            remainingAttachmentSlots--;
                        }
                    }
                }

                workspace.ContextDocuments.AddRange(pending);
                added = pending.Count;
            }
            catch
            {
                foreach (ConversationContextDocument document in pending)
                {
                    document.Dispose();
                }

                if (temporary)
                {
                    DeleteTemporaryFiles(candidates);
                }

                throw;
            }

            workspace.ContextDocumentsSent = false;
            RefreshConversationSummary();
            RefreshContextDocumentsDisplay();
            SetOperationStatus(
                workspace,
                BuildDocumentImportStatus(
                    added,
                    separatedAttachmentCount,
                    skippedAttachmentCount,
                    attachmentExtractionFailed));
            _composer.Focus();
        }

        private static string BuildDocumentImportStatus(
            int addedCount,
            int separatedAttachmentCount,
            int skippedAttachmentCount,
            bool attachmentExtractionFailed)
        {
            string status = addedCount == 1
                ? "Document ajouté au contexte de la conversation."
                : addedCount + " documents ajoutés au contexte de la conversation.";
            if (separatedAttachmentCount > 0)
            {
                status += " " + separatedAttachmentCount +
                    (separatedAttachmentCount == 1
                        ? " pièce jointe séparée."
                        : " pièces jointes séparées.");
            }

            if (skippedAttachmentCount > 0)
            {
                status += " " + skippedAttachmentCount +
                    (skippedAttachmentCount == 1
                        ? " pièce jointe ignorée (limite ou taille)."
                        : " pièces jointes ignorées (limite ou taille).");
            }

            if (attachmentExtractionFailed)
            {
                status += " Certaines pièces jointes n'ont pas pu être séparées.";
            }

            return status;
        }

        private static void DeleteTemporaryFiles(IEnumerable<string> paths)
        {
            var directories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in paths ?? Enumerable.Empty<string>())
            {
                string directory = null;
                try { directory = Path.GetDirectoryName(path); } catch (Exception) { }
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    directories.Add(directory);
                }

                try { File.Delete(path); } catch (Exception) { }
            }

            foreach (string directory in directories)
            {
                try
                {
                    if (Directory.Exists(directory) &&
                        !Directory.EnumerateFileSystemEntries(directory).Any())
                    {
                        Directory.Delete(directory, false);
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        private void ContextDocument_DoubleClick(object sender, EventArgs e)
        {
            var control = sender as Control;
            var document = control == null
                ? null
                : control.Tag as ConversationContextDocument;
            if (document == null)
            {
                return;
            }

            try
            {
                document.Open(_addIn.Application);
                SetOperationStatus(_activeWorkspace, "Document ouvert : " + document.DisplayName);
            }
            catch (Exception exception)
            {
                LocalLogger.Error(
                    "Impossible d'ouvrir le document de contexte (" +
                    exception.GetType().Name + ").");
                ShowError(GetUserFacingError(exception));
            }
        }

        private void RemoveContextDocument_Click(object sender, EventArgs e)
        {
            var button = sender as Button;
            var document = button == null
                ? null
                : button.Tag as ConversationContextDocument;
            ConversationWorkspace workspace = _activeWorkspace;
            if (document == null || workspace == null || workspace.IsBusy ||
                !workspace.ContextDocuments.Contains(document))
            {
                return;
            }

            if (workspace.ContextDocumentsSent &&
                HasConversationContent(workspace) && !ConfirmConversationReset(
                "Ce document fait déjà partie du contexte Codex. Le retirer redémarre " +
                "cette conversation. Continuer ?"))
            {
                return;
            }

            if (workspace.ContextDocumentsSent)
            {
                ResetConversation(workspace, false);
            }

            workspace.ContextDocuments.Remove(document);
            document.Dispose();
            workspace.ContextDocumentsSent = false;
            RefreshConversationSummary();
            RefreshContextDocumentsDisplay();
            SetOperationStatus(workspace, "Document retiré du contexte.");
        }

        private void NewConversationButton_Click(object sender, EventArgs e)
        {
            ConversationWorkspace workspace = CreateConversationWorkspace(true);
            SetOperationStatus(workspace, "Nouvelle conversation indépendante.");
            _composer.Focus();
        }

        private async void SendButton_Click(object sender, EventArgs e)
        {
            await SendCurrentMessageAsync();
        }

        private async Task SendCurrentMessageAsync()
        {
            ConversationWorkspace workspace = _activeWorkspace;
            if (workspace == null ||
                workspace.IsBusy ||
                !_isConnected ||
                IsDisposed)
            {
                return;
            }

            string userText = (workspace.ComposerText ?? string.Empty).Trim();
            if (userText.Length == 0)
            {
                SetOperationStatus(
                    workspace,
                    "Saisissez une demande avant de l'envoyer.");
                _composer.Focus();
                return;
            }

            if (workspace.IsUntitled)
            {
                workspace.Title = BuildConversationTitle(userText);
                workspace.IsUntitled = false;
                RefreshConversationSummary();
            }

            CancellationTokenSource turnCancellation =
                CancellationTokenSource.CreateLinkedTokenSource(
                    _lifetimeCancellation.Token);
            CodexOperationHandle trackedOperation = _operationTracker.Begin(
                "Conversation — " + workspace.Title,
                () => turnCancellation.Cancel());
            workspace.TurnCancellation = turnCancellation;
            SetBusyState(workspace, true);
            AppendTranscript(workspace, "Vous", userText);
            if (workspace.ContextDocuments.Count > 0 && !workspace.ContextDocumentsSent)
            {
                AppendSystemNote(
                    workspace,
                    workspace.ContextDocuments.Count == 1
                        ? "Document joint : " + workspace.ContextDocuments[0].DisplayName
                        : workspace.ContextDocuments.Count + " documents joints au contexte.");
            }

            BeginAssistantTranscript(workspace);
            var streamedText = new StringBuilder();
            using (var progress = new BufferedUiProgress(
                this,
                delta =>
                {
                    streamedText.Append(delta);
                    AppendAssistantDelta(workspace, delta);
                }))
            {
                try
                {
                    CodexService service = _addIn.GetCodexService();
                    if (workspace.Conversation == null)
                    {
                        SetOperationStatus(
                            workspace,
                            "Initialisation de la conversation…");
                        workspace.Conversation = await service.StartConversationAsync(
                            turnCancellation.Token);
                    }

                    IReadOnlyList<CodexContextFile> contextFiles = null;
                    if (workspace.ContextDocuments.Count > 0 &&
                        !workspace.ContextDocumentsSent)
                    {
                        List<ConversationContextDocument> documents =
                            workspace.ContextDocuments.ToList();
                        SetOperationStatus(workspace, "Extraction locale des documents…");
                        contextFiles = await Task.Run(
                            () => PrepareContextFiles(
                                documents,
                                turnCancellation.Token),
                            turnCancellation.Token);
                        if (!IsDisposed && !Disposing)
                        {
                            RefreshContextDocumentsDisplay();
                        }

                        // Le serveur peut recevoir le prompt même si le transport échoue ensuite.
                        // Ne jamais le renvoyer silencieusement dans cette conversation.
                        workspace.ContextDocumentsSent = true;
                        SetOperationStatus(workspace, "Préparation des documents pour Codex…");
                    }

                    SetOperationStatus(workspace, "Codex répond…");
                    CodexTurnResult result = await service.SendConversationMessageAsync(
                        workspace.Conversation,
                        userText,
                        contextFiles,
                        progress,
                        turnCancellation.Token);
                    progress.Flush();
                    trackedOperation.Complete("réponse reçue");
                    if (IsDisposed || Disposing)
                    {
                        return;
                    }

                    string finalText = result == null ? null : result.Text;
                    if (streamedText.Length == 0 && !string.IsNullOrEmpty(finalText))
                    {
                        AppendAssistantDelta(workspace, finalText);
                    }

                    workspace.LastResponse = string.IsNullOrWhiteSpace(finalText)
                        ? streamedText.ToString()
                        : finalText;
                    workspace.DraftCreatedForLastResponse = false;
                    EndAssistantTranscript(workspace, null);
                    workspace.ComposerText = string.Empty;
                    if (ReferenceEquals(workspace, _activeWorkspace))
                    {
                        _composer.Clear();
                    }

                    SetOperationStatus(workspace, "Réponse terminée.");
                }
                catch (OperationCanceledException)
                {
                    trackedOperation.Cancel("demande annulée");
                    progress.Flush();
                    workspace.LastResponse = null;
                    EndAssistantTranscript(workspace, "Réponse annulée.");
                    InvalidateConversationAfterFailedTurn(workspace);
                    SetOperationStatus(
                        workspace,
                        "Génération annulée. Le prochain envoi démarrera un nouveau fil.");
                }
                catch (Exception exception)
                {
                    trackedOperation.Fail(GetUserFacingError(exception));
                    progress.Flush();
                    workspace.LastResponse = null;
                    EndAssistantTranscript(
                        workspace,
                        "Erreur : " + GetUserFacingError(exception));
                    InvalidateConversationAfterFailedTurn(workspace);
                    LocalLogger.Error(
                        "Échec d’un tour Codex dans le volet (" +
                        exception.GetType().Name + ").");
                    SetOperationStatus(workspace, GetUserFacingError(exception));
                }
                finally
                {
                    if (ReferenceEquals(workspace.TurnCancellation, turnCancellation))
                    {
                        workspace.TurnCancellation = null;
                    }

                    turnCancellation.Dispose();
                    SetBusyState(workspace, false);
                }
            }
        }

        private static IReadOnlyList<CodexContextFile> PrepareContextFiles(
            IReadOnlyList<ConversationContextDocument> documents,
            CancellationToken cancellationToken)
        {
            int remainingCharacters = MaximumExtractedContextCharacters;
            var result = new List<CodexContextFile>(documents.Count);
            foreach (ConversationContextDocument document in documents)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int documentLimit = Math.Min(
                    remainingCharacters,
                    DocumentContextExtractor.MaximumCharactersPerDocument);
                CodexContextFile contextFile = document.ToCodexContextFile(
                    documentLimit,
                    cancellationToken);
                result.Add(contextFile);
                remainingCharacters = Math.Max(
                    0,
                    remainingCharacters -
                        (contextFile.ExtractedText == null
                            ? 0
                            : contextFile.ExtractedText.Length));
            }

            return result.AsReadOnly();
        }

        private void InvalidateConversationAfterFailedTurn(
            ConversationWorkspace workspace)
        {
            bool hadConversation = workspace.Conversation != null;
            InvalidateWorkspaceConversation(workspace);

            workspace.ContextDocumentsSent = false;
            if (hadConversation)
            {
                AppendSystemNote(
                    workspace,
                    "Le contexte Codex précédent n'est pas repris. Le prochain message " +
                    "démarrera un nouveau fil.");
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            ConversationWorkspace workspace = _activeWorkspace;
            CancelWorkspaceTurn(workspace);
            SetOperationStatus(workspace, "Annulation demandée…");
        }

        private void CopyButton_Click(object sender, EventArgs e)
        {
            ConversationWorkspace workspace = _activeWorkspace;
            if (workspace == null || string.IsNullOrWhiteSpace(workspace.LastResponse))
            {
                return;
            }

            try
            {
                Clipboard.SetText(workspace.LastResponse);
                SetOperationStatus(workspace, "Dernière réponse copiée.");
            }
            catch (Exception exception)
            {
                LocalLogger.Error(
                    "Impossible de copier la réponse Codex (" +
                    exception.GetType().Name + ").");
                ShowError("Le Presse-papiers n'est pas disponible.");
            }
        }

        private void DraftButton_Click(object sender, EventArgs e)
        {
            ConversationWorkspace workspace = _activeWorkspace;
            if (workspace == null ||
                FindFirstMailContext(workspace) == null ||
                string.IsNullOrWhiteSpace(workspace.LastResponse) ||
                workspace.IsBusy)
            {
                return;
            }

            try
            {
                OutlookMailContext mailContext = FindFirstMailContext(workspace);
                bool opened = mailContext.CreatePlainTextDraft(
                    _addIn.Application,
                    workspace.LastResponse);
                workspace.DraftCreatedForLastResponse = true;
                _draftButton.Enabled = false;
                MessageBox.Show(
                    this,
                    opened
                        ? "Le brouillon a été enregistré et ouvert. Aucun destinataire n'a " +
                            "été ajouté et aucun message n'a été envoyé."
                        : "Le brouillon a bien été enregistré, mais sa fenêtre n'a pas pu " +
                            "être ouverte. Retrouvez-le dans le dossier Brouillons ; ne créez " +
                            "pas un second brouillon pour cette réponse.",
                    "Outcom — Codex",
                    MessageBoxButtons.OK,
                    opened ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
                SetOperationStatus(
                    workspace,
                    opened
                        ? "Brouillon Outlook créé — aucun envoi effectué."
                        : "Brouillon enregistré dans le dossier Brouillons, mais non ouvert.");
            }
            catch (Exception exception)
            {
                LocalLogger.Error(
                    "Impossible de créer le brouillon Outlook (" +
                    exception.GetType().Name + ").");
                ShowError(GetUserFacingError(exception));
            }
        }

        private void Composer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
                if (_activeWorkspace != null &&
                    !_activeWorkspace.IsBusy &&
                    _isConnected)
                {
                    _ = SendCurrentMessageAsync();
                }
            }
        }

        private void SetConnectionState(bool connected, string label, string operationStatus)
        {
            if (IsDisposed)
            {
                return;
            }

            _isConnected = connected;
            _connectionStatusLabel.Text = label;
            RefreshActiveWorkspaceControls();
            SetOperationStatus(_activeWorkspace, operationStatus);
        }

        private void SetBusyState(ConversationWorkspace workspace, bool busy)
        {
            if (workspace == null || IsDisposed || Disposing)
            {
                return;
            }

            workspace.IsBusy = busy;
            RefreshConversationSummary();
            if (ReferenceEquals(workspace, _activeWorkspace))
            {
                RefreshContextDocumentsDisplay();
                RefreshActiveWorkspaceControls();
            }
        }

        private void RefreshActiveWorkspaceControls()
        {
            ConversationWorkspace workspace = _activeWorkspace;
            bool busy = workspace != null && workspace.IsBusy;
            _contextDocumentsPanel.AllowDrop = !busy && workspace != null;
            _newConversationButton.Enabled = true;
            _closeConversationButton.Enabled = workspace != null && !busy;
            _cancelButton.Enabled = busy;
            _sendButton.Enabled = !busy &&
                _isConnected &&
                workspace != null;
            _composer.Enabled = !busy && _isConnected && workspace != null;
            _copyButton.Enabled = !busy &&
                workspace != null &&
                !string.IsNullOrWhiteSpace(workspace.LastResponse);
            _draftButton.Enabled = !busy &&
                workspace != null &&
                FindFirstMailContext(workspace) != null &&
                !string.IsNullOrWhiteSpace(workspace.LastResponse) &&
                !workspace.DraftCreatedForLastResponse;
            _operationStatusLabel.Text = workspace == null
                ? "Aucune conversation sélectionnée."
                : workspace.Status;
        }

        private void RefreshActiveWorkspaceDisplay()
        {
            ConversationWorkspace workspace = _activeWorkspace;
            _transcript.Text = workspace == null
                ? string.Empty
                : workspace.Transcript.ToString();
            _transcript.SelectionStart = _transcript.TextLength;
            _transcript.ScrollToCaret();
            _composer.Text = workspace == null
                ? string.Empty
                : workspace.ComposerText ?? string.Empty;
            _composer.SelectionStart = _composer.TextLength;
            RefreshContextDocumentsDisplay();
            RefreshActiveWorkspaceControls();
        }

        private void RefreshContextDocumentsDisplay()
        {
            foreach (Control control in _contextDocumentsPanel.Controls.Cast<Control>().ToList())
            {
                _contextDocumentsPanel.Controls.Remove(control);
                if (!ReferenceEquals(control, _contextDropHintLabel))
                {
                    control.Dispose();
                }
            }

            ConversationWorkspace workspace = _activeWorkspace;
            if (workspace == null || workspace.ContextDocuments.Count == 0)
            {
                _contextDocumentsPanel.Controls.Add(_contextDropHintLabel);
            }
            else
            {
                foreach (ConversationContextDocument document in workspace.ContextDocuments)
                {
                    _contextDocumentsPanel.Controls.Add(CreateContextDocumentCard(
                        document,
                        !workspace.IsBusy));
                }
            }

            _draftButton.Enabled = workspace != null &&
                !workspace.IsBusy &&
                FindFirstMailContext(workspace) != null &&
                !string.IsNullOrWhiteSpace(workspace.LastResponse) &&
                !workspace.DraftCreatedForLastResponse;
        }

        private Control CreateContextDocumentCard(
            ConversationContextDocument document,
            bool removalEnabled)
        {
            var card = new Panel
            {
                Size = new Size(82, 82),
                BackColor = SystemColors.Control,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(2),
                Tag = document,
                Cursor = Cursors.Hand,
                AccessibleName = document.DisplayName,
                AccessibleDescription = document.Details +
                    ". Double-cliquez pour ouvrir le document."
            };
            var icon = new PictureBox
            {
                Image = document.CreateIcon(),
                Size = new Size(64, 64),
                Location = new Point(8, 8),
                SizeMode = PictureBoxSizeMode.Zoom,
                Tag = document,
                Cursor = Cursors.Hand,
                TabStop = false
            };
            var remove = new Button
            {
                Text = "×",
                FlatStyle = FlatStyle.Flat,
                Size = new Size(24, 24),
                Location = new Point(56, 1),
                Margin = Padding.Empty,
                TabStop = true,
                Enabled = removalEnabled,
                Tag = document,
                BackColor = Color.Firebrick,
                ForeColor = Color.White,
                Font = new Font(Font.FontFamily, 12F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                UseVisualStyleBackColor = false,
                AccessibleName = "Retirer " + document.DisplayName,
                AccessibleDescription =
                    "Retirer ce document du contexte de la conversation."
            };
            remove.FlatAppearance.BorderSize = 1;
            remove.FlatAppearance.BorderColor = Color.DarkRed;
            remove.FlatAppearance.MouseOverBackColor = Color.Red;
            remove.Click += RemoveContextDocument_Click;
            card.DoubleClick += ContextDocument_DoubleClick;
            icon.DoubleClick += ContextDocument_DoubleClick;
            AttachDocumentDropHandlers(card);
            AttachDocumentDropHandlers(icon);
            card.Controls.Add(icon);
            card.Controls.Add(remove);
            var toolTip = new ToolTip();
            toolTip.SetToolTip(card, document.DisplayName + Environment.NewLine + document.Details);
            toolTip.SetToolTip(icon, document.DisplayName + Environment.NewLine + document.Details);
            toolTip.SetToolTip(remove, "Retirer du contexte");
            card.Disposed += (sender, args) => toolTip.Dispose();
            return card;
        }

        private void AttachDocumentDropHandlers(Control control)
        {
            control.AllowDrop = true;
            control.DragEnter += ContextDocumentsPanel_DragEnter;
            control.DragOver += ContextDocumentsPanel_DragEnter;
            control.DragLeave += ContextDocumentsPanel_DragLeave;
            control.DragDrop += ContextDocumentsPanel_DragDrop;
        }

        private static OutlookMailContext FindFirstMailContext(
            ConversationWorkspace workspace)
        {
            if (workspace == null)
            {
                return null;
            }

            ConversationContextDocument document = workspace.ContextDocuments.FirstOrDefault(
                item => item.MailContext != null);
            return document == null ? null : document.MailContext;
        }

        private void ResetConversation(
            ConversationWorkspace workspace,
            bool clearTranscript)
        {
            if (workspace == null)
            {
                return;
            }

            InvalidateWorkspaceConversation(workspace);
            workspace.ContextDocumentsSent = false;
            workspace.LastResponse = null;
            workspace.DraftCreatedForLastResponse = false;
            if (clearTranscript || workspace.Transcript.Length > 0)
            {
                workspace.Transcript.Clear();
            }

            if (ReferenceEquals(workspace, _activeWorkspace))
            {
                RefreshActiveWorkspaceDisplay();
            }
        }

        private static void InvalidateWorkspaceConversation(
            ConversationWorkspace workspace)
        {
            if (workspace != null && workspace.Conversation != null)
            {
                workspace.Conversation.Invalidate();
                workspace.Conversation = null;
            }
        }

        private static bool HasConversationContent(ConversationWorkspace workspace)
        {
            return workspace != null &&
                (workspace.Conversation != null || workspace.Transcript.Length > 0);
        }

        private bool ConfirmConversationReset(string message)
        {
            return MessageBox.Show(
                this,
                message,
                "Outcom — Codex",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2) == DialogResult.Yes;
        }

        private void AppendTranscript(
            ConversationWorkspace workspace,
            string role,
            string text)
        {
            AppendPlainText(workspace, role + Environment.NewLine);
            AppendPlainText(workspace, text + Environment.NewLine + Environment.NewLine);
        }

        private void AppendSystemNote(ConversationWorkspace workspace, string text)
        {
            AppendPlainText(
                workspace,
                "[" + text + "]" + Environment.NewLine + Environment.NewLine);
        }

        private void BeginAssistantTranscript(ConversationWorkspace workspace)
        {
            AppendPlainText(workspace, "Codex" + Environment.NewLine);
        }

        private void AppendAssistantDelta(
            ConversationWorkspace workspace,
            string delta)
        {
            if (!string.IsNullOrEmpty(delta))
            {
                AppendPlainText(workspace, delta);
            }
        }

        private void EndAssistantTranscript(
            ConversationWorkspace workspace,
            string note)
        {
            if (!string.IsNullOrWhiteSpace(note))
            {
                AppendPlainText(
                    workspace,
                    Environment.NewLine + "[" + note + "]");
            }

            AppendPlainText(workspace, Environment.NewLine + Environment.NewLine);
        }

        private void AppendPlainText(ConversationWorkspace workspace, string text)
        {
            if (workspace == null ||
                IsDisposed ||
                Disposing ||
                string.IsNullOrEmpty(text))
            {
                return;
            }

            workspace.Transcript.Append(text);
            if (!ReferenceEquals(workspace, _activeWorkspace))
            {
                return;
            }

            _transcript.SelectionStart = _transcript.TextLength;
            _transcript.SelectionLength = 0;
            _transcript.SelectedText = text;
            _transcript.SelectionStart = _transcript.TextLength;
            _transcript.ScrollToCaret();
        }

        private void SetOperationStatus(
            ConversationWorkspace workspace,
            string text)
        {
            if (workspace == null || IsDisposed || Disposing)
            {
                return;
            }

            workspace.Status = text;
            if (ReferenceEquals(workspace, _activeWorkspace))
            {
                _operationStatusLabel.Text = text;
            }
        }

        private void ShowError(string message)
        {
            SetOperationStatus(_activeWorkspace, message);
            MessageBox.Show(
                this,
                message,
                "Outcom — Codex",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        private static string GetUserFacingError(Exception exception)
        {
            if (exception is CodexAppServerException ||
                exception is InvalidOperationException ||
                exception is ArgumentException)
            {
                return exception.Message;
            }

            return "L'opération a échoué. Consultez le journal Outcom.";
        }

        private static string BuildConversationTitle(string text)
        {
            string title = string.IsNullOrWhiteSpace(text)
                ? "Nouvelle conversation"
                : text.Trim().Replace('\r', ' ').Replace('\n', ' ');
            while (title.Contains("  "))
            {
                title = title.Replace("  ", " ");
            }

            return title.Length <= 48 ? title : title.Substring(0, 45) + "…";
        }

        private static void DisposeContextDocuments(ConversationWorkspace workspace)
        {
            if (workspace == null)
            {
                return;
            }

            foreach (ConversationContextDocument document in workspace.ContextDocuments)
            {
                document.Dispose();
            }

            workspace.ContextDocuments.Clear();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && Interlocked.Exchange(ref _disposeState, 1) == 0)
            {
                Load -= CodexTaskPaneControl_Load;
                _service.GlobalContextChanged -= CodexService_GlobalContextChanged;
                _operationTracker.Changed -= CodexOperationTracker_Changed;
                CancelActiveTurn();
                foreach (ConversationWorkspace workspace in _workspaces)
                {
                    InvalidateWorkspaceConversation(workspace);
                    DisposeContextDocuments(workspace);
                }

                _lifetimeCancellation.Cancel();
                _brandImage.Dispose();
            }

            base.Dispose(disposing);
        }

        private sealed class ConversationWorkspace
        {
            internal ConversationWorkspace(int sequence, string title)
            {
                Sequence = sequence;
                Title = title;
                Transcript = new StringBuilder();
                ContextDocuments = new List<ConversationContextDocument>();
                Status = "Prêt — Ctrl+Entrée pour envoyer.";
                IsUntitled = true;
            }

            internal int Sequence { get; private set; }

            internal string Title { get; set; }

            internal bool IsUntitled { get; set; }

            internal StringBuilder Transcript { get; private set; }

            internal string ComposerText { get; set; }

            internal string Status { get; set; }

            internal CodexConversationSession Conversation { get; set; }

            internal List<ConversationContextDocument> ContextDocuments { get; private set; }

            internal bool ContextDocumentsSent { get; set; }

            internal string LastResponse { get; set; }

            internal bool DraftCreatedForLastResponse { get; set; }

            internal CancellationTokenSource TurnCancellation { get; set; }

            internal bool IsBusy { get; set; }

            public override string ToString()
            {
                string prefix = IsBusy ? "En cours — " : string.Empty;
                string context = ContextDocuments.Count == 0
                    ? string.Empty
                    : " — " + ContextDocuments.Count +
                        (ContextDocuments.Count == 1 ? " document" : " documents");
                return prefix + Title + context;
            }
        }

        /// <summary>
        /// Tampon thread-safe : les notifications JSONL ne touchent jamais WinForms
        /// directement et plusieurs petits deltas sont regroupés avant affichage.
        /// </summary>
        private sealed class BufferedUiProgress : IProgress<string>, IDisposable
        {
            private readonly Control _owner;
            private readonly Action<string> _append;
            private readonly object _syncRoot = new object();
            private readonly StringBuilder _buffer = new StringBuilder();
            private int _scheduled;
            private int _disposed;

            internal BufferedUiProgress(Control owner, Action<string> append)
            {
                _owner = owner;
                _append = append;
            }

            public void Report(string value)
            {
                if (string.IsNullOrEmpty(value) || Volatile.Read(ref _disposed) != 0)
                {
                    return;
                }

                lock (_syncRoot)
                {
                    _buffer.Append(value);
                }

                ScheduleDrain();
            }

            internal void Flush()
            {
                if (_owner.IsDisposed || _owner.Disposing)
                {
                    return;
                }

                if (_owner.InvokeRequired)
                {
                    _owner.Invoke(new Action(Drain));
                }
                else
                {
                    Drain();
                }
            }

            private void ScheduleDrain()
            {
                if (Interlocked.CompareExchange(ref _scheduled, 1, 0) != 0)
                {
                    return;
                }

                try
                {
                    _owner.BeginInvoke(new Action(Drain));
                }
                catch (InvalidOperationException)
                {
                    Interlocked.Exchange(ref _scheduled, 0);
                }
            }

            private void Drain()
            {
                string text;
                lock (_syncRoot)
                {
                    text = _buffer.ToString();
                    _buffer.Clear();
                }

                Interlocked.Exchange(ref _scheduled, 0);
                if (!string.IsNullOrEmpty(text) && Volatile.Read(ref _disposed) == 0)
                {
                    _append(text);
                }

                lock (_syncRoot)
                {
                    if (_buffer.Length > 0)
                    {
                        ScheduleDrain();
                    }
                }
            }

            public void Dispose()
            {
                Flush();
                Interlocked.Exchange(ref _disposed, 1);
            }
        }
    }
}
