using System;
using System.Drawing;
using System.Windows.Forms;

namespace Outcom.AddIn
{
    [System.ComponentModel.DesignerCategory("Code")]
    internal sealed class OutcomAboutForm : Form
    {
        private readonly Icon _windowIcon;
        private readonly Bitmap _headerImage;

        internal OutcomAboutForm()
        {
            OutcomBuildInfo buildInfo = OutcomBuildInfo.Read();
            _windowIcon = OutcomBranding.CreateWindowIcon();
            _headerImage = OutcomBranding.CreateHeaderImage();

            Text = "À propos d’Outcom";
            Icon = _windowIcon;
            ShowIcon = true;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            SizeGripStyle = SizeGripStyle.Show;
            MaximizeBox = true;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(700, 510);
            MinimumSize = new Size(570, 430);

            var root = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 4,
                Dock = DockStyle.Fill,
                Padding = new Padding(20)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            root.Controls.Add(CreateHeader(), 0, 0);

            var versionLabel = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 18, 0, 14),
                Font = new Font(Font, FontStyle.Bold),
                Text = "Version " + buildInfo.Version +
                    "  •  compilée le " + buildInfo.BuildDateDisplay,
                AccessibleName = "Version et date de compilation d’Outcom"
            };
            root.Controls.Add(versionLabel, 0, 1);

            var information = new TableLayoutPanel
            {
                ColumnCount = 2,
                RowCount = 7,
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(0, 2, 0, 0)
            };
            information.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 165F));
            information.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            for (int row = 0; row < information.RowCount; row++)
            {
                information.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            AddInformationRow(
                information,
                0,
                "Compatibilité",
                "Outlook classique pour Windows — .NET Framework 4.8");
            AddInformationRow(
                information,
                1,
                "Architecture active",
                buildInfo.ProcessArchitecture);
            AddInformationRow(
                information,
                2,
                "Connexion IA",
                "Codex app-server avec le compte ChatGPT de l’utilisateur, sans clé API");
            AddInformationRow(
                information,
                3,
                "Runtime validé",
                CodexRuntimeManifest.Current.CliVersion + " — Windows x64");
            AddInformationRow(
                information,
                4,
                "Configuration",
                buildInfo.AssemblyConfiguration);
            AddInformationRow(
                information,
                5,
                "Journal local",
                LocalLogger.LogFilePath);
            AddInformationRow(
                information,
                6,
                "Licence",
                "MIT — Copyright © 2026, contributeurs Outcom");
            root.Controls.Add(information, 0, 2);

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Margin = new Padding(0, 16, 0, 0)
            };
            var closeButton = new Button
            {
                AutoSize = true,
                DialogResult = DialogResult.OK,
                Text = "Fermer",
                AccessibleName = "Fermer la fenêtre À propos d’Outcom"
            };
            buttons.Controls.Add(closeButton);
            root.Controls.Add(buttons, 0, 3);

            AcceptButton = closeButton;
            CancelButton = closeButton;
        }

        private Control CreateHeader()
        {
            var header = new TableLayoutPanel
            {
                ColumnCount = 2,
                RowCount = 1,
                Dock = DockStyle.Top,
                AutoSize = true,
                Margin = Padding.Empty
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 82F));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            var icon = new PictureBox
            {
                Image = _headerImage,
                Size = new Size(64, 64),
                SizeMode = PictureBoxSizeMode.Zoom,
                Margin = new Padding(0, 0, 18, 0),
                AccessibleName = "Icône Outcom"
            };
            header.Controls.Add(icon, 0, 0);

            var titles = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 2,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Margin = Padding.Empty
            };
            titles.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            titles.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            titles.Controls.Add(
                new Label
                {
                    AutoSize = true,
                    Font = new Font(Font.FontFamily, 20F, FontStyle.Bold),
                    Margin = new Padding(0, 2, 0, 4),
                    Text = "Outcom"
                },
                0,
                0);
            titles.Controls.Add(
                new Label
                {
                    AutoSize = true,
                    Margin = Padding.Empty,
                    Text = "Le lien local entre Outlook et ChatGPT par Codex"
                },
                0,
                1);
            header.Controls.Add(titles, 1, 0);
            return header;
        }

        private static void AddInformationRow(
            TableLayoutPanel panel,
            int row,
            string name,
            string value)
        {
            panel.Controls.Add(
                new Label
                {
                    AutoSize = true,
                    Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                    Margin = new Padding(0, 6, 10, 10),
                    Text = name
                },
                0,
                row);

            panel.Controls.Add(
                new Label
                {
                    AutoSize = true,
                    MaximumSize = new Size(430, 0),
                    Margin = new Padding(0, 6, 0, 10),
                    Text = value,
                    AccessibleName = name + " : " + value
                },
                1,
                row);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _headerImage.Dispose();
                _windowIcon.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
