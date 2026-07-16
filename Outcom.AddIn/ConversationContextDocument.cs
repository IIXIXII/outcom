using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outcom.AddIn
{
    /// <summary>
    /// Référence un document natif explicitement déposé dans une conversation.
    /// Le fichier natif reste intact ; une représentation textuelle bornée peut être
    /// préparée localement au premier envoi à Codex.
    /// </summary>
    internal sealed class ConversationContextDocument : IDisposable
    {
        private const long MaximumFileSize = 50L * 1024L * 1024L;
        private const uint ShgfiIcon = 0x000000100;
        private const uint ShgfiLargeIcon = 0x000000000;
        private readonly object _extractionLock = new object();
        private bool _disposed;
        private string _sizeDescription;
        private string _originDescription;
        private DocumentExtractionResult _extraction;

        private ConversationContextDocument()
        {
        }

        internal string Id { get; private set; }

        internal string DisplayName { get; private set; }

        internal string Details { get; private set; }

        internal string FilePath { get; private set; }

        internal bool IsTemporary { get; private set; }

        internal OutlookMailContext MailContext { get; private set; }

        internal static ConversationContextDocument FromFile(
            string filePath,
            bool isTemporary,
            OutlookMailContext mailContext = null,
            string originDescription = null)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("Le chemin du document est vide.", nameof(filePath));
            }

            string fullPath = Path.GetFullPath(filePath);
            var file = new FileInfo(fullPath);
            if (!file.Exists)
            {
                throw new FileNotFoundException("Le document déposé est introuvable.", fullPath);
            }

            if (file.Length > MaximumFileSize)
            {
                throw new InvalidOperationException(
                    "Le document « " + file.Name + " » dépasse la limite de 50 Mo.");
            }

            return new ConversationContextDocument
            {
                Id = Guid.NewGuid().ToString("N"),
                DisplayName = file.Name,
                Details = BuildDetails(
                    FormatSize(file.Length),
                    originDescription,
                    "en attente d'extraction locale"),
                FilePath = fullPath,
                IsTemporary = isTemporary,
                MailContext = mailContext,
                _sizeDescription = FormatSize(file.Length),
                _originDescription = originDescription
            };
        }

        internal CodexContextFile ToCodexContextFile(
            int maximumCharacters,
            CancellationToken cancellationToken)
        {
            if (_disposed || string.IsNullOrWhiteSpace(FilePath) || !File.Exists(FilePath))
            {
                throw new InvalidOperationException(
                    "Le document « " + (DisplayName ?? string.Empty) + " » n'est plus disponible.");
            }

            DocumentExtractionResult extraction;
            lock (_extractionLock)
            {
                if (_extraction == null)
                {
                    _extraction = DocumentContextExtractor.Extract(
                        FilePath,
                        DisplayName,
                        DocumentContextExtractor.MaximumCharactersPerDocument,
                        cancellationToken);
                    int characterCount = string.IsNullOrEmpty(_extraction.TextContent)
                        ? 0
                        : _extraction.TextContent.Length;
                    Details = BuildDetails(
                        _sizeDescription,
                        _originDescription,
                        _extraction.Status) +
                        (characterCount > 0
                            ? " (" + characterCount.ToString("N0", CultureInfo.CurrentCulture) +
                                " caractères)"
                            : string.Empty);
                }

                extraction = _extraction;
            }

            string transmittedText = LimitForTransmission(
                extraction.TextContent,
                maximumCharacters);
            bool limitedForTransmission =
                !string.IsNullOrEmpty(extraction.TextContent) &&
                (transmittedText == null ||
                    transmittedText.Length < extraction.TextContent.Length);

            return new CodexContextFile
            {
                Id = Id,
                DisplayName = DisplayName,
                SourcePath = FilePath,
                ExtractedText = transmittedText,
                ExtractionStatus = extraction.Status +
                    (limitedForTransmission
                        ? " ; texte limité par le budget global de contexte"
                        : string.Empty),
                IsLocalImage = extraction.IsLocalImage
            };
        }

        private static string LimitForTransmission(string text, int maximumCharacters)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            int limit = Math.Max(0, maximumCharacters);
            if (text.Length <= limit)
            {
                return text;
            }

            const string Marker = "\r\n[Contenu limité par le budget global de contexte.]";
            if (limit <= Marker.Length)
            {
                return null;
            }

            return text.Substring(0, limit - Marker.Length).TrimEnd() + Marker;
        }

        internal Image CreateIcon()
        {
            IntPtr iconHandle = IntPtr.Zero;
            try
            {
                var info = new ShellFileInfo();
                IntPtr result = SHGetFileInfo(
                    FilePath,
                    0,
                    ref info,
                    (uint)Marshal.SizeOf(typeof(ShellFileInfo)),
                    ShgfiIcon | ShgfiLargeIcon);
                iconHandle = info.IconHandle;
                if (result != IntPtr.Zero && iconHandle != IntPtr.Zero)
                {
                    using (Icon shellIcon = (Icon)Icon.FromHandle(iconHandle).Clone())
                    {
                        return new Bitmap(shellIcon.ToBitmap(), new Size(64, 64));
                    }
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                if (iconHandle != IntPtr.Zero)
                {
                    DestroyIcon(iconHandle);
                }
            }

            return new Bitmap(SystemIcons.Application.ToBitmap(), new Size(64, 64));
        }

        internal void Open(Outlook.Application application)
        {
            if (string.IsNullOrWhiteSpace(FilePath) || !File.Exists(FilePath))
            {
                throw new InvalidOperationException("Le document n'est plus disponible.");
            }

            try
            {
                Process.Start(new ProcessStartInfo(FilePath) { UseShellExecute = true });
            }
            catch (Exception exception)
            {
                if (MailContext != null)
                {
                    MailContext.Open(application);
                    return;
                }

                throw new InvalidOperationException(
                    "Windows n'a pas pu ouvrir le document.",
                    exception);
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            if (IsTemporary && !string.IsNullOrWhiteSpace(FilePath))
            {
                try
                {
                    File.Delete(FilePath);
                    string directory = Path.GetDirectoryName(FilePath);
                    if (!string.IsNullOrWhiteSpace(directory) &&
                        Directory.Exists(directory) &&
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

        private static string FormatSize(long bytes)
        {
            if (bytes < 1024)
            {
                return bytes.ToString(CultureInfo.CurrentCulture) + " octets";
            }

            if (bytes < 1024L * 1024L)
            {
                return (bytes / 1024D).ToString("0.#", CultureInfo.CurrentCulture) + " Ko";
            }

            return (bytes / (1024D * 1024D)).ToString("0.#", CultureInfo.CurrentCulture) + " Mo";
        }

        private static string BuildDetails(
            string size,
            string originDescription,
            string status)
        {
            string origin = string.IsNullOrWhiteSpace(originDescription)
                ? string.Empty
                : " — pièce jointe de « " + originDescription.Trim() + " »";
            return size + origin + " — " + status;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SHGetFileInfo(
            string path,
            uint fileAttributes,
            ref ShellFileInfo fileInfo,
            uint fileInfoSize,
            uint flags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr iconHandle);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct ShellFileInfo
        {
            public IntPtr IconHandle;
            public int IconIndex;
            public uint Attributes;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string DisplayName;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string TypeName;
        }
    }
}
