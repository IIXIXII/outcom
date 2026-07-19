using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Outcom.AddIn
{
    [System.ComponentModel.DesignerCategory("Code")]
    internal sealed class OutcomGlobalContextForm : Form
    {
        private readonly OutcomGlobalContext _initialContext;
        private readonly CodexService _service;
        private readonly Icon _windowIcon;
        private readonly CancellationTokenSource _lifetimeCancellation =
            new CancellationTokenSource();
        private readonly ComboBox _modelComboBox;
        private readonly ComboBox _reasoningEffortComboBox;
        private readonly Label _modelStatusLabel;
        private readonly TextBox _workContextTextBox;
        private readonly TextBox _vocabularyTextBox;
        private readonly TextBox _instructionsTextBox;
        private readonly CheckBox _insertProposalAtBeginningOnConflictCheckBox;
        private readonly Label _characterCountLabel;
        private readonly Button _saveButton;

        private OutcomGlobalContext _resultContext;
        private bool _modelsLoaded;
        private bool _updatingModelChoices;

        internal OutcomGlobalContextForm(OutcomGlobalContext context, CodexService service)
        {
            _initialContext = OutcomGlobalContext.ValidateAndNormalize(context);
            _resultContext = _initialContext.Clone();
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _windowIcon = OutcomBranding.CreateWindowIcon();

            Text = "Contexte Codex";
            Icon = _windowIcon;
            ShowIcon = true;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            SizeGripStyle = SizeGripStyle.Show;
            MaximizeBox = true;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(1640, 1035);
            MinimumSize = new Size(620, 500);

            var root = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 5,
                Dock = DockStyle.Fill,
                Padding = new Padding(16)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            var introduction = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(780, 0),
                Margin = new Padding(0, 0, 0, 12),
                Text =
                    "Ce contexte, le modèle et la profondeur de raisonnement s'appliquent à " +
                    "toutes les nouvelles conversations Outcom. Les réglages sont chiffrés " +
                    "pour votre compte Windows lorsqu'ils sont stockés localement."
            };
            root.Controls.Add(introduction, 0, 0);
            root.SizeChanged += (sender, args) =>
            {
                int availableWidth = Math.Max(240, root.ClientSize.Width - root.Padding.Horizontal);
                if (introduction.MaximumSize.Width != availableWidth)
                {
                    introduction.MaximumSize = new Size(availableWidth, 0);
                }
            };

            var settings = new TableLayoutPanel
            {
                ColumnCount = 2,
                RowCount = 3,
                Dock = DockStyle.Top,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 12)
            };
            settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            settings.Controls.Add(CreateSettingLabel("Modèle :"), 0, 0);
            _modelComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled = false,
                AccessibleName = "Modèle Codex général"
            };
            _modelComboBox.SelectedIndexChanged += ModelComboBox_SelectedIndexChanged;
            settings.Controls.Add(_modelComboBox, 1, 0);
            settings.Controls.Add(CreateSettingLabel("Profondeur de raisonnement :"), 0, 1);
            _reasoningEffortComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled = false,
                AccessibleName = "Profondeur de raisonnement générale"
            };
            settings.Controls.Add(_reasoningEffortComboBox, 1, 1);
            _modelStatusLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 5, 0, 0),
                Text = "Chargement des modèles Codex…",
                ForeColor = SystemColors.GrayText,
                AccessibleName = "État du catalogue des modèles"
            };
            settings.Controls.Add(_modelStatusLabel, 1, 2);
            root.Controls.Add(settings, 0, 1);

            var tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                AccessibleName = "Sections du contexte Codex"
            };
            tabs.TabPages.Add(CreateEditorTab(
                "Contexte de travail",
                "Décrivez votre activité, votre rôle, les dossiers récurrents et les éléments " +
                    "que Codex doit connaître dans toutes les conversations.",
                _initialContext.WorkContext,
                "Contexte de travail général",
                out _workContextTextBox));
            tabs.TabPages.Add(CreateEditorTab(
                "Vocabulaire",
                "Indiquez les termes à privilégier, les formulations à éviter, les sigles et " +
                    "leur signification, ainsi que les règles de nommage.",
                _initialContext.VocabularyGuidelines,
                "Directives générales de vocabulaire",
                out _vocabularyTextBox));
            tabs.TabPages.Add(CreateEditorTab(
                "Instructions transversales",
                "Définissez la langue, le ton, le niveau de détail, la structure attendue et " +
                    "les validations à effectuer avant de proposer une réponse.",
                _initialContext.CrossConversationInstructions,
                "Instructions transversales générales",
                out _instructionsTextBox));
            tabs.TabPages.Add(CreateBehaviorTab(
                _initialContext.InsertProposalAtBeginningOnConflict,
                out _insertProposalAtBeginningOnConflictCheckBox));
            root.Controls.Add(tabs, 0, 2);

            _characterCountLabel = new Label
            {
                AutoSize = true,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0, 8, 0, 0),
                AccessibleName = "Nombre de caractères du contexte Codex"
            };
            root.Controls.Add(_characterCountLabel, 0, 3);

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = true,
                Margin = new Padding(0, 12, 0, 0)
            };
            root.Controls.Add(buttons, 0, 4);

            var cancelButton = new Button
            {
                AutoSize = true,
                DialogResult = DialogResult.Cancel,
                Text = "Annuler"
            };
            buttons.Controls.Add(cancelButton);

            _saveButton = new Button { AutoSize = true, Text = "Enregistrer", Enabled = false };
            _saveButton.Click += SaveButton_Click;
            buttons.Controls.Add(_saveButton);

            var clearButton = new Button { AutoSize = true, Text = "Effacer les directives" };
            clearButton.Click += ClearButton_Click;
            buttons.Controls.Add(clearButton);

            AcceptButton = _saveButton;
            CancelButton = cancelButton;
            _workContextTextBox.TextChanged += EditorTextChanged;
            _vocabularyTextBox.TextChanged += EditorTextChanged;
            _instructionsTextBox.TextChanged += EditorTextChanged;
            Shown += Form_Shown;
            UpdateCharacterCount();
        }

        internal OutcomGlobalContext ResultContext
        {
            get { return _resultContext.Clone(); }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _lifetimeCancellation.Cancel();
                _lifetimeCancellation.Dispose();
                _windowIcon.Dispose();
            }

            base.Dispose(disposing);
        }

        private async void Form_Shown(object sender, EventArgs e)
        {
            try
            {
                IReadOnlyList<CodexModelInfo> models = await _service.ListModelsAsync(
                    _lifetimeCancellation.Token);
                if (IsDisposed)
                {
                    return;
                }

                ApplyModels(models);
                _workContextTextBox.Focus();
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception exception)
            {
                LocalLogger.Error(
                    "Impossible de charger les modèles du contexte Codex (" +
                    exception.GetType().Name + ").");
                _modelStatusLabel.Text = "Modèles indisponibles — configurez la connexion Codex.";
                _modelStatusLabel.ForeColor = Color.DarkRed;
                MessageBox.Show(
                    this,
                    "Les modèles Codex n'ont pas pu être chargés. Vérifiez la connexion avec " +
                        "Configurer Codex dans le ruban, puis rouvrez cette fenêtre.",
                    "Outcom — Contexte Codex",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ApplyModels(IReadOnlyList<CodexModelInfo> models)
        {
            _updatingModelChoices = true;
            try
            {
                _modelComboBox.Items.Clear();
                foreach (CodexModelInfo model in models)
                {
                    _modelComboBox.Items.Add(model);
                }

                int selectedIndex = FindModelIndex(_initialContext.ModelId);
                if (selectedIndex < 0)
                {
                    selectedIndex = FindDefaultModelIndex();
                }

                if (selectedIndex < 0 && _modelComboBox.Items.Count > 0)
                {
                    selectedIndex = 0;
                }

                _modelComboBox.SelectedIndex = selectedIndex;
                PopulateReasoningEfforts(
                    _modelComboBox.SelectedItem as CodexModelInfo,
                    _initialContext.ReasoningEffort);
                _modelsLoaded = selectedIndex >= 0;
                _modelComboBox.Enabled = _modelsLoaded;
                _reasoningEffortComboBox.Enabled = _modelsLoaded;
                _modelStatusLabel.Text = _modelsLoaded
                    ? "Ces choix s'appliqueront aux nouveaux fils et aux propositions de réponse."
                    : "Aucun modèle texte n'est disponible pour ce compte.";
                _modelStatusLabel.ForeColor = _modelsLoaded ? SystemColors.GrayText : Color.DarkRed;
            }
            finally
            {
                _updatingModelChoices = false;
            }

            UpdateCharacterCount();
        }

        private void ModelComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_updatingModelChoices)
            {
                return;
            }

            string currentEffort = (_reasoningEffortComboBox.SelectedItem as
                CodexReasoningEffortInfo)?.Value;
            PopulateReasoningEfforts(
                _modelComboBox.SelectedItem as CodexModelInfo,
                currentEffort);
        }

        private void PopulateReasoningEfforts(CodexModelInfo model, string selectedValue)
        {
            _reasoningEffortComboBox.Items.Clear();
            _reasoningEffortComboBox.Items.Add(new CodexReasoningEffortInfo
            {
                Value = string.Empty,
                Description = "Utiliser la profondeur par défaut annoncée par le modèle."
            });

            if (model?.SupportedReasoningEfforts != null)
            {
                foreach (CodexReasoningEffortInfo effort in model.SupportedReasoningEfforts)
                {
                    _reasoningEffortComboBox.Items.Add(effort);
                }
            }

            int selectedIndex = 0;
            if (!string.IsNullOrWhiteSpace(selectedValue))
            {
                for (int index = 1; index < _reasoningEffortComboBox.Items.Count; index++)
                {
                    var effort = (CodexReasoningEffortInfo)_reasoningEffortComboBox.Items[index];
                    if (string.Equals(effort.Value, selectedValue, StringComparison.OrdinalIgnoreCase))
                    {
                        selectedIndex = index;
                        break;
                    }
                }
            }

            _reasoningEffortComboBox.SelectedIndex = selectedIndex;
        }

        private int FindModelIndex(string modelId)
        {
            for (int index = 0; index < _modelComboBox.Items.Count; index++)
            {
                var model = (CodexModelInfo)_modelComboBox.Items[index];
                if (string.Equals(model.Id, modelId, StringComparison.Ordinal))
                {
                    return index;
                }
            }

            return -1;
        }

        private int FindDefaultModelIndex()
        {
            for (int index = 0; index < _modelComboBox.Items.Count; index++)
            {
                if (((CodexModelInfo)_modelComboBox.Items[index]).IsDefault)
                {
                    return index;
                }
            }

            return -1;
        }

        private static Label CreateSettingLabel(string text)
        {
            return new Label
            {
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 6, 12, 5),
                Text = text
            };
        }

        private static TabPage CreateEditorTab(
            string title,
            string description,
            string value,
            string accessibleName,
            out TextBox editor)
        {
            var page = new TabPage
            {
                Text = title,
                Padding = new Padding(10),
                UseVisualStyleBackColor = true
            };
            var layout = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 2,
                Dock = DockStyle.Fill,
                Margin = Padding.Empty
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            page.Controls.Add(layout);

            var descriptionLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                MaximumSize = new Size(700, 0),
                Margin = new Padding(0, 0, 0, 8),
                Text = description
            };
            layout.Controls.Add(descriptionLabel, 0, 0);
            page.SizeChanged += (sender, args) =>
            {
                int availableWidth = Math.Max(200, page.ClientSize.Width - page.Padding.Horizontal);
                if (descriptionLabel.MaximumSize.Width != availableWidth)
                {
                    descriptionLabel.MaximumSize = new Size(availableWidth, 0);
                }
            };

            editor = new TextBox
            {
                AcceptsReturn = true,
                AcceptsTab = true,
                Dock = DockStyle.Fill,
                MaxLength = OutcomGlobalContext.MaximumSectionLength,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Text = value ?? string.Empty,
                AccessibleName = accessibleName
            };
            layout.Controls.Add(editor, 0, 1);
            return page;
        }

        private static TabPage CreateBehaviorTab(
            bool insertProposalAtBeginningOnConflict,
            out CheckBox insertProposalAtBeginningOnConflictCheckBox)
        {
            var page = new TabPage("Comportement")
            {
                Padding = new Padding(12),
                UseVisualStyleBackColor = true
            };
            var layout = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 2,
                Dock = DockStyle.Top,
                AutoSize = true
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            page.Controls.Add(layout);

            insertProposalAtBeginningOnConflictCheckBox = new CheckBox
            {
                AutoSize = true,
                Checked = insertProposalAtBeginningOnConflict,
                Margin = new Padding(0, 0, 0, 10),
                Text =
                    "Si le remplacement précis est impossible, insérer la proposition " +
                    "au début du message",
                AccessibleName =
                    "Insérer la proposition au début du message en cas de conflit"
            };
            layout.Controls.Add(insertProposalAtBeginningOnConflictCheckBox, 0, 0);
            layout.Controls.Add(
                new Label
                {
                    AutoSize = true,
                    Dock = DockStyle.Fill,
                    ForeColor = SystemColors.GrayText,
                    Margin = Padding.Empty,
                    Text =
                        "Outcom conserve alors tout le contenu existant et ajoute la proposition " +
                        "tout en haut. Ce repli reste interdit si le brouillon a été envoyé, " +
                        "fermé ou remplacé par un autre message."
                },
                0,
                1);
            return page;
        }

        private void EditorTextChanged(object sender, EventArgs e)
        {
            UpdateCharacterCount();
        }

        private void UpdateCharacterCount()
        {
            int totalLength = _workContextTextBox.Text.Length +
                _vocabularyTextBox.Text.Length +
                _instructionsTextBox.Text.Length;
            _characterCountLabel.Text = totalLength + " / " +
                OutcomGlobalContext.MaximumTotalLength + " caractères";
            _characterCountLabel.ForeColor = totalLength > OutcomGlobalContext.MaximumTotalLength
                ? Color.DarkRed
                : SystemColors.ControlText;
            _saveButton.Enabled = _modelsLoaded &&
                totalLength <= OutcomGlobalContext.MaximumTotalLength;
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            if (_workContextTextBox.TextLength == 0 &&
                _vocabularyTextBox.TextLength == 0 &&
                _instructionsTextBox.TextLength == 0)
            {
                return;
            }

            DialogResult answer = MessageBox.Show(
                this,
                "Effacer les trois sections de directives ? Le modèle et la profondeur de " +
                    "raisonnement seront conservés.",
                "Outcom — Contexte Codex",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
            if (answer == DialogResult.Yes)
            {
                _workContextTextBox.Clear();
                _vocabularyTextBox.Clear();
                _instructionsTextBox.Clear();
            }
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            var selectedModel = _modelComboBox.SelectedItem as CodexModelInfo;
            var selectedEffort = _reasoningEffortComboBox.SelectedItem as
                CodexReasoningEffortInfo;
            if (selectedModel == null || selectedEffort == null)
            {
                return;
            }

            OutcomGlobalContext context;
            try
            {
                context = OutcomGlobalContext.ValidateAndNormalize(new OutcomGlobalContext
                {
                    WorkContext = _workContextTextBox.Text,
                    VocabularyGuidelines = _vocabularyTextBox.Text,
                    CrossConversationInstructions = _instructionsTextBox.Text,
                    ModelId = selectedModel.Id,
                    ReasoningEffort = selectedEffort.Value,
                    InsertProposalAtBeginningOnConflict =
                        _insertProposalAtBeginningOnConflictCheckBox.Checked
                });
            }
            catch (ArgumentException exception)
            {
                MessageBox.Show(
                    this,
                    exception.Message,
                    "Outcom — Contexte Codex",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (!_initialContext.ContentEquals(context))
            {
                DialogResult answer = MessageBox.Show(
                    this,
                    "Enregistrer ce contexte Codex ? Les conversations Outcom ouvertes seront " +
                        "réinitialisées pour appliquer les nouveaux réglages.",
                    "Outcom — Contexte Codex",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);
                if (answer != DialogResult.Yes)
                {
                    return;
                }
            }

            _resultContext = context;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
