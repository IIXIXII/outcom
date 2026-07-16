using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using MsgReader.Mime;
using MsgReader.Mime.Header;
using MsgReader.Outlook;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace Outcom.AddIn
{
    /// <summary>
    /// Extrait localement une représentation textuelle bornée d'un document.
    /// Le fichier natif reste toujours la source de référence et n'est jamais modifié.
    /// </summary>
    internal static class DocumentContextExtractor
    {
        internal const int MaximumCharactersPerDocument = 60000;
        private const int MaximumPdfPages = 300;
        private const long MaximumXmlCharacters = 10000000;
        private const long MaximumExtractedAttachmentSize = 25L * 1024L * 1024L;

        private static readonly HashSet<string> PlainTextExtensions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".txt", ".md", ".markdown", ".csv", ".tsv", ".json", ".xml",
                ".yaml", ".yml", ".log", ".ini", ".config", ".sql", ".cs",
                ".vb", ".fs", ".js", ".jsx", ".ts", ".tsx", ".css", ".scss",
                ".ps1", ".psm1", ".py", ".java", ".c", ".h", ".cpp", ".hpp"
            };

        private static readonly HashSet<string> ImageExtensions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".png", ".jpg", ".jpeg", ".gif", ".webp"
            };

        static DocumentContextExtractor()
        {
            // Certaines dépendances net462 demandent une ancienne version forte de
            // System.Memory. Un complément VSTO ne contrôle pas Outlook.exe.config :
            // on unifie donc uniquement cette assembly vers la copie déployée avec Outcom.
            AppDomain.CurrentDomain.AssemblyResolve += ResolveSystemMemory;
        }

        internal static DocumentExtractionResult Extract(
            string filePath,
            string displayName,
            int maximumCharacters,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("Le chemin du document est vide.", nameof(filePath));
            }

            int limit = Math.Max(0, Math.Min(
                maximumCharacters,
                MaximumCharactersPerDocument));
            string extension = Path.GetExtension(filePath) ?? string.Empty;
            if (ImageExtensions.Contains(extension))
            {
                return DocumentExtractionResult.Image(
                    "image transmise directement à Codex");
            }

            if (limit == 0)
            {
                return DocumentExtractionResult.MetadataOnly(
                    "limite globale de texte atteinte ; fichier natif conservé");
            }

            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                LimitedTextBuilder output = new LimitedTextBuilder(limit);
                output.AppendLine("# " + SafeSingleLine(displayName));
                output.AppendLine();
                bool contentDetected = true;

                switch (extension.ToLowerInvariant())
                {
                    case ".pdf":
                        contentDetected = ExtractPdf(filePath, output, cancellationToken);
                        break;
                    case ".msg":
                        ExtractMsg(filePath, output, cancellationToken);
                        break;
                    case ".eml":
                        ExtractEml(filePath, output, cancellationToken);
                        break;
                    case ".docx":
                        ExtractDocx(filePath, output, cancellationToken);
                        break;
                    case ".pptx":
                        ExtractPptx(filePath, output, cancellationToken);
                        break;
                    case ".xlsx":
                        ExtractXlsx(filePath, output, cancellationToken);
                        break;
                    case ".html":
                    case ".htm":
                        output.AppendLine(HtmlToText(ReadTextFile(
                            filePath,
                            Math.Min(limit * 3, MaximumCharactersPerDocument * 3),
                            cancellationToken)));
                        break;
                    default:
                        if (PlainTextExtensions.Contains(extension))
                        {
                            output.AppendLine(ReadTextFile(filePath, limit, cancellationToken));
                        }
                        else
                        {
                            return DocumentExtractionResult.MetadataOnly(
                                "format non extrait ; fichier natif conservé");
                        }

                        break;
                }

                string text = output.ToString();
                if (!contentDetected ||
                    string.IsNullOrWhiteSpace(RemoveHeading(text, displayName)))
                {
                    return DocumentExtractionResult.MetadataOnly(
                        extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase)
                            ? "aucun texte détecté dans le PDF ; fichier natif conservé"
                            : "aucun texte détecté ; fichier natif conservé");
                }

                if (output.IsTruncated)
                {
                    string marker = Environment.NewLine +
                        "[Extraction tronquée afin de respecter la limite de contexte.]";
                    int contentLength = Math.Max(0, limit - marker.Length);
                    if (text.Length > contentLength)
                    {
                        text = text.Substring(0, contentLength).TrimEnd();
                    }

                    text += marker;
                }

                return DocumentExtractionResult.Text(
                    text,
                    output.IsTruncated
                        ? "texte extrait partiellement"
                        : "texte extrait",
                    output.IsTruncated);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception)
            {
                LocalLogger.Error(
                    "Extraction du document de contexte impossible pour " +
                    SafeSingleLine(displayName) + " (" +
                    exception.GetType().Name + ").");
                return DocumentExtractionResult.MetadataOnly(
                    "extraction impossible ; fichier natif conservé");
            }
        }

        internal static EmailAttachmentExtractionResult ExtractEmailAttachments(
            string filePath,
            int maximumCount,
            CancellationToken cancellationToken)
        {
            string extension = Path.GetExtension(filePath) ?? string.Empty;
            if (!extension.Equals(".msg", StringComparison.OrdinalIgnoreCase) &&
                !extension.Equals(".eml", StringComparison.OrdinalIgnoreCase))
            {
                return EmailAttachmentExtractionResult.Empty();
            }

            string directory = Path.Combine(
                Path.GetTempPath(),
                "Outcom",
                "Context",
                "Attachments-" + Guid.NewGuid().ToString("N"));
            var paths = new List<string>();
            int skippedCount = 0;
            try
            {
                Directory.CreateDirectory(directory);
                if (extension.Equals(".msg", StringComparison.OrdinalIgnoreCase))
                {
                    using (var message = new MsgReader.Outlook.Storage.Message(
                        filePath,
                        FileAccess.Read))
                    {
                        foreach (object item in message.Attachments ??
                            Enumerable.Empty<object>())
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            var attachment = item as MsgReader.Outlook.Storage.Attachment;
                            if (attachment != null)
                            {
                                if (attachment.IsInline)
                                {
                                    continue;
                                }

                                if (paths.Count >= maximumCount)
                                {
                                    skippedCount++;
                                    continue;
                                }

                                byte[] data = attachment.Data;
                                if (data == null ||
                                    data.LongLength > MaximumExtractedAttachmentSize)
                                {
                                    skippedCount++;
                                    continue;
                                }

                                string target = CreateAttachmentPath(
                                    directory,
                                    attachment.FileName,
                                    paths.Count + 1,
                                    ".bin");
                                File.WriteAllBytes(target, data);
                                paths.Add(target);
                                continue;
                            }

                            var embeddedMessage = item as MsgReader.Outlook.Storage.Message;
                            if (embeddedMessage == null)
                            {
                                skippedCount++;
                                continue;
                            }

                            if (paths.Count >= maximumCount)
                            {
                                skippedCount++;
                                continue;
                            }

                            string messageTarget = CreateAttachmentPath(
                                directory,
                                embeddedMessage.FileName,
                                paths.Count + 1,
                                ".msg");
                            embeddedMessage.Save(messageTarget);
                            if (new FileInfo(messageTarget).Length >
                                MaximumExtractedAttachmentSize)
                            {
                                File.Delete(messageTarget);
                                skippedCount++;
                            }
                            else
                            {
                                paths.Add(messageTarget);
                            }
                        }
                    }
                }
                else
                {
                    Message message;
                    using (FileStream stream = File.OpenRead(filePath))
                    {
                        message = Message.Load(stream, false, true);
                    }

                    foreach (MessagePart attachment in message.Attachments ??
                        Enumerable.Empty<MessagePart>())
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (attachment.IsInline)
                        {
                            continue;
                        }

                        if (paths.Count >= maximumCount)
                        {
                            skippedCount++;
                            continue;
                        }

                        if (attachment.Body == null ||
                            attachment.Body.LongLength > MaximumExtractedAttachmentSize)
                        {
                            skippedCount++;
                            continue;
                        }

                        string target = CreateAttachmentPath(
                            directory,
                            attachment.FileName,
                            paths.Count + 1,
                            ".bin");
                        attachment.Save(new FileInfo(target));
                        paths.Add(target);
                    }
                }

                if (paths.Count == 0)
                {
                    DeleteAttachmentDirectory(directory, paths);
                }

                return EmailAttachmentExtractionResult.Success(
                    paths.AsReadOnly(),
                    skippedCount);
            }
            catch (OperationCanceledException)
            {
                DeleteAttachmentDirectory(directory, paths);
                throw;
            }
            catch (Exception exception)
            {
                DeleteAttachmentDirectory(directory, paths);
                LocalLogger.Error(
                    "Séparation des pièces jointes impossible pour " +
                    SafeSingleLine(Path.GetFileName(filePath)) + " (" +
                    exception.GetType().Name + ").");
                return EmailAttachmentExtractionResult.Failure();
            }
        }

        private static Assembly ResolveSystemMemory(object sender, ResolveEventArgs arguments)
        {
            AssemblyName requested;
            try
            {
                requested = new AssemblyName(arguments.Name);
            }
            catch (Exception)
            {
                return null;
            }

            if (!string.Equals(
                requested.Name,
                "System.Memory",
                StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            Assembly loaded = AppDomain.CurrentDomain.GetAssemblies()
                .FirstOrDefault(assembly => string.Equals(
                    assembly.GetName().Name,
                    requested.Name,
                    StringComparison.OrdinalIgnoreCase));
            if (loaded != null)
            {
                return loaded;
            }

            string directory = Path.GetDirectoryName(typeof(DocumentContextExtractor)
                .Assembly.Location);
            string path = string.IsNullOrWhiteSpace(directory)
                ? null
                : Path.Combine(directory, "System.Memory.dll");
            return !string.IsNullOrWhiteSpace(path) && File.Exists(path)
                ? Assembly.LoadFrom(path)
                : null;
        }

        private static string CreateAttachmentPath(
            string directory,
            string fileName,
            int index,
            string fallbackExtension)
        {
            string name = string.IsNullOrWhiteSpace(fileName)
                ? "Pièce jointe " + index.ToString(CultureInfo.InvariantCulture) +
                    fallbackExtension
                : Path.GetFileName(fileName);
            foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(invalidCharacter, '_');
            }

            name = name.Trim().TrimEnd('.');
            if (string.IsNullOrWhiteSpace(name))
            {
                name = "Pièce jointe " + index.ToString(CultureInfo.InvariantCulture) +
                    fallbackExtension;
            }

            if (string.IsNullOrWhiteSpace(Path.GetExtension(name)))
            {
                name += fallbackExtension;
            }

            string stem = Path.GetFileNameWithoutExtension(name);
            string extension = Path.GetExtension(name);
            string target = Path.Combine(directory, name);
            int duplicateIndex = 2;
            while (File.Exists(target))
            {
                target = Path.Combine(
                    directory,
                    stem + " (" + duplicateIndex.ToString(CultureInfo.InvariantCulture) +
                    ")" + extension);
                duplicateIndex++;
            }

            return target;
        }

        private static void DeleteAttachmentDirectory(
            string directory,
            IEnumerable<string> paths)
        {
            foreach (string path in paths ?? Enumerable.Empty<string>())
            {
                try { File.Delete(path); } catch (Exception) { }
            }

            try
            {
                if (Directory.Exists(directory))
                {
                    Directory.Delete(directory, true);
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool ExtractPdf(
            string filePath,
            LimitedTextBuilder output,
            CancellationToken cancellationToken)
        {
            bool contentDetected = false;
            using (PdfDocument document = PdfDocument.Open(filePath))
            {
                int count = Math.Min(document.NumberOfPages, MaximumPdfPages);
                for (int pageNumber = 1; pageNumber <= count && !output.IsFull; pageNumber++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    output.AppendLine("## Page " + pageNumber.ToString(CultureInfo.InvariantCulture));
                    output.AppendLine();
                    string pageText = ContentOrderTextExtractor.GetText(
                        document.GetPage(pageNumber),
                        true);
                    if (!string.IsNullOrWhiteSpace(pageText))
                    {
                        contentDetected = true;
                    }

                    output.AppendLine(pageText);
                    output.AppendLine();
                }

                if (document.NumberOfPages > MaximumPdfPages && !output.IsFull)
                {
                    output.AppendLine(
                        "[PDF limité aux " + MaximumPdfPages.ToString(CultureInfo.InvariantCulture) +
                        " premières pages.]" );
                }
            }

            return contentDetected;
        }

        private static void ExtractMsg(
            string filePath,
            LimitedTextBuilder output,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            using (var message = new MsgReader.Outlook.Storage.Message(
                filePath,
                FileAccess.Read))
            {
                output.AppendLine("## Métadonnées");
                AppendMetadata(output, "Objet", message.Subject);
                AppendMetadata(output, "De", message.GetEmailSender(false, false));
                AppendMetadata(output, "À", message.GetEmailRecipients(
                    RecipientType.To,
                    false,
                    false));
                AppendMetadata(output, "Cc", message.GetEmailRecipients(
                    RecipientType.Cc,
                    false,
                    false));
                if (message.SentOn.HasValue)
                {
                    AppendMetadata(
                        output,
                        "Date",
                        message.SentOn.Value.ToLocalTime().ToString("G", CultureInfo.CurrentCulture));
                }

                AppendOutlookAttachments(output, message.Attachments);
                output.AppendLine();
                output.AppendLine("## Corps du message");
                output.AppendLine();
                output.AppendLine(!string.IsNullOrWhiteSpace(message.BodyText)
                    ? message.BodyText
                    : HtmlToText(message.BodyHtml));
            }
        }

        private static void ExtractEml(
            string filePath,
            LimitedTextBuilder output,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Message message;
            using (FileStream stream = File.OpenRead(filePath))
            {
                message = Message.Load(stream, false, false);
            }
            MessageHeader headers = message.Headers;
            output.AppendLine("## Métadonnées");
            AppendMetadata(output, "Objet", headers == null ? null : headers.Subject);
            AppendMetadata(output, "De", headers == null ? null : FormatAddress(headers.From));
            AppendMetadata(output, "À", headers == null ? null : FormatAddresses(headers.To));
            AppendMetadata(output, "Cc", headers == null ? null : FormatAddresses(headers.Cc));
            if (headers != null && headers.DateSent != DateTimeOffset.MinValue)
            {
                AppendMetadata(
                    output,
                    "Date",
                    headers.DateSent.ToLocalTime().ToString("G", CultureInfo.CurrentCulture));
            }

            if (message.Attachments != null && message.Attachments.Count > 0)
            {
                AppendMetadata(
                    output,
                    "Pièces jointes",
                    string.Join(", ", message.Attachments
                        .Select(item => SafeSingleLine(item.FileName))
                        .Where(item => !string.IsNullOrWhiteSpace(item))));
            }

            output.AppendLine();
            output.AppendLine("## Corps du message");
            output.AppendLine();
            if (message.TextBody != null)
            {
                output.AppendLine(message.TextBody.GetBodyAsText());
            }
            else if (message.HtmlBody != null)
            {
                output.AppendLine(HtmlToText(message.HtmlBody.GetBodyAsText()));
            }
        }

        private static void ExtractDocx(
            string filePath,
            LimitedTextBuilder output,
            CancellationToken cancellationToken)
        {
            using (ZipArchive archive = ZipFile.OpenRead(filePath))
            {
                var parts = archive.Entries
                    .Where(entry =>
                        entry.FullName.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase) ||
                        Regex.IsMatch(
                            entry.FullName,
                            @"^word/(header|footer)\d+\.xml$",
                            RegexOptions.IgnoreCase))
                    .OrderBy(entry => entry.FullName)
                    .ToList();

                foreach (ZipArchiveEntry part in parts)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (output.IsFull)
                    {
                        break;
                    }

                    output.AppendLine(part.FullName.Equals(
                        "word/document.xml",
                        StringComparison.OrdinalIgnoreCase)
                        ? "## Document"
                        : "## " + Path.GetFileNameWithoutExtension(part.Name));
                    XDocument xml = LoadXml(part);
                    foreach (XElement paragraph in xml.Descendants()
                        .Where(element => element.Name.LocalName == "p"))
                    {
                        string text = string.Concat(paragraph.Descendants()
                            .Where(element => element.Name.LocalName == "t")
                            .Select(element => element.Value));
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            output.AppendLine(text.Trim());
                        }

                        if (output.IsFull)
                        {
                            break;
                        }
                    }

                    output.AppendLine();
                }
            }
        }

        private static void ExtractPptx(
            string filePath,
            LimitedTextBuilder output,
            CancellationToken cancellationToken)
        {
            using (ZipArchive archive = ZipFile.OpenRead(filePath))
            {
                var slides = archive.Entries
                    .Where(entry => Regex.IsMatch(
                        entry.FullName,
                        @"^ppt/slides/slide\d+\.xml$",
                        RegexOptions.IgnoreCase))
                    .OrderBy(entry => GetTrailingNumber(entry.Name))
                    .ToList();

                int slideNumber = 0;
                foreach (ZipArchiveEntry slide in slides)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (output.IsFull)
                    {
                        break;
                    }

                    slideNumber++;
                    output.AppendLine("## Diapositive " +
                        slideNumber.ToString(CultureInfo.InvariantCulture));
                    XDocument xml = LoadXml(slide);
                    foreach (XElement paragraph in xml.Descendants()
                        .Where(element => element.Name.LocalName == "p"))
                    {
                        string text = string.Concat(paragraph.Descendants()
                            .Where(element => element.Name.LocalName == "t")
                            .Select(element => element.Value));
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            output.AppendLine(text.Trim());
                        }
                    }

                    output.AppendLine();
                }
            }
        }

        private static void ExtractXlsx(
            string filePath,
            LimitedTextBuilder output,
            CancellationToken cancellationToken)
        {
            using (ZipArchive archive = ZipFile.OpenRead(filePath))
            {
                List<string> sharedStrings = ReadSharedStrings(archive);
                var sheets = archive.Entries
                    .Where(entry => Regex.IsMatch(
                        entry.FullName,
                        @"^xl/worksheets/sheet\d+\.xml$",
                        RegexOptions.IgnoreCase))
                    .OrderBy(entry => GetTrailingNumber(entry.Name))
                    .ToList();

                int sheetNumber = 0;
                foreach (ZipArchiveEntry sheet in sheets)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (output.IsFull)
                    {
                        break;
                    }

                    sheetNumber++;
                    output.AppendLine("## Feuille " +
                        sheetNumber.ToString(CultureInfo.InvariantCulture));
                    XDocument xml = LoadXml(sheet);
                    foreach (XElement row in xml.Descendants()
                        .Where(element => element.Name.LocalName == "row"))
                    {
                        var cells = new List<string>();
                        foreach (XElement cell in row.Elements()
                            .Where(element => element.Name.LocalName == "c"))
                        {
                            string value = ReadSpreadsheetCell(cell, sharedStrings);
                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                string reference = (string)cell.Attribute("r");
                                cells.Add((string.IsNullOrWhiteSpace(reference)
                                    ? string.Empty
                                    : reference + " = ") + value);
                            }
                        }

                        if (cells.Count > 0)
                        {
                            output.AppendLine(string.Join(" | ", cells));
                        }

                        if (output.IsFull)
                        {
                            break;
                        }
                    }

                    output.AppendLine();
                }
            }
        }

        private static List<string> ReadSharedStrings(ZipArchive archive)
        {
            ZipArchiveEntry entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
            {
                return new List<string>();
            }

            XDocument xml = LoadXml(entry);
            return xml.Descendants()
                .Where(element => element.Name.LocalName == "si")
                .Select(item => string.Concat(item.Descendants()
                    .Where(element => element.Name.LocalName == "t")
                    .Select(element => element.Value)))
                .ToList();
        }

        private static string ReadSpreadsheetCell(
            XElement cell,
            IReadOnlyList<string> sharedStrings)
        {
            string cellType = (string)cell.Attribute("t");
            if (string.Equals(cellType, "inlineStr", StringComparison.OrdinalIgnoreCase))
            {
                return string.Concat(cell.Descendants()
                    .Where(element => element.Name.LocalName == "t")
                    .Select(element => element.Value));
            }

            XElement valueElement = cell.Elements()
                .FirstOrDefault(element => element.Name.LocalName == "v");
            string value = valueElement == null ? string.Empty : valueElement.Value;
            int sharedStringIndex;
            if (string.Equals(cellType, "s", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture,
                    out sharedStringIndex) &&
                sharedStringIndex >= 0 &&
                sharedStringIndex < sharedStrings.Count)
            {
                return sharedStrings[sharedStringIndex];
            }

            return value;
        }

        private static XDocument LoadXml(ZipArchiveEntry entry)
        {
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaximumXmlCharacters
            };
            using (Stream stream = entry.Open())
            using (XmlReader reader = XmlReader.Create(stream, settings))
            {
                return XDocument.Load(reader, LoadOptions.None);
            }
        }

        private static void AppendOutlookAttachments(
            LimitedTextBuilder output,
            IEnumerable<object> attachments)
        {
            if (attachments == null)
            {
                return;
            }

            var names = new List<string>();
            foreach (object attachment in attachments)
            {
                var file = attachment as MsgReader.Outlook.Storage.Attachment;
                var embeddedMessage = attachment as MsgReader.Outlook.Storage.Message;
                string name = file != null
                    ? file.FileName
                    : embeddedMessage == null ? null : embeddedMessage.FileName;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    names.Add(SafeSingleLine(name));
                }
            }

            if (names.Count > 0)
            {
                AppendMetadata(output, "Pièces jointes", string.Join(", ", names));
            }
        }

        private static string FormatAddress(RfcMailAddress address)
        {
            if (address == null)
            {
                return string.Empty;
            }

            return string.IsNullOrWhiteSpace(address.DisplayName)
                ? address.Address
                : address.DisplayName + " <" + address.Address + ">";
        }

        private static string FormatAddresses(IEnumerable<RfcMailAddress> addresses)
        {
            return addresses == null
                ? string.Empty
                : string.Join(", ", addresses.Select(FormatAddress));
        }

        private static void AppendMetadata(
            LimitedTextBuilder output,
            string name,
            string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                output.AppendLine("- " + name + " : " + SafeSingleLine(value));
            }
        }

        private static string ReadTextFile(
            string filePath,
            int maximumCharacters,
            CancellationToken cancellationToken)
        {
            var result = new StringBuilder(Math.Min(maximumCharacters, 8192));
            var buffer = new char[4096];
            using (var reader = new StreamReader(
                filePath,
                Encoding.UTF8,
                true,
                4096))
            {
                while (result.Length < maximumCharacters)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int requested = Math.Min(buffer.Length, maximumCharacters - result.Length);
                    int read = reader.Read(buffer, 0, requested);
                    if (read <= 0)
                    {
                        break;
                    }

                    result.Append(buffer, 0, read);
                }
            }

            return result.ToString();
        }

        private static string HtmlToText(string html)
        {
            if (string.IsNullOrWhiteSpace(html))
            {
                return string.Empty;
            }

            string value = Regex.Replace(
                html,
                @"<(script|style)[^>]*>.*?</\1>",
                string.Empty,
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            value = Regex.Replace(value, @"<(br|/p|/div|/li|/tr)\b[^>]*>", "\n",
                RegexOptions.IgnoreCase);
            value = Regex.Replace(value, @"<[^>]+>", " ");
            value = WebUtility.HtmlDecode(value);
            value = Regex.Replace(value, @"[ \t]+", " ");
            value = Regex.Replace(value, @"(\r?\n)[ \t]+", "$1");
            value = Regex.Replace(value, @"(\r?\n){3,}", Environment.NewLine + Environment.NewLine);
            return value.Trim();
        }

        private static int GetTrailingNumber(string fileName)
        {
            Match match = Regex.Match(fileName ?? string.Empty, @"(\d+)");
            int value;
            return match.Success && int.TryParse(match.Groups[1].Value, out value)
                ? value
                : int.MaxValue;
        }

        private static string SafeSingleLine(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? string.Empty
                : Regex.Replace(value.Trim(), @"\s+", " ");
        }

        private static string RemoveHeading(string text, string displayName)
        {
            string heading = "# " + SafeSingleLine(displayName);
            return (text ?? string.Empty).Replace(heading, string.Empty).Trim();
        }

        private sealed class LimitedTextBuilder
        {
            private readonly int _limit;
            private readonly StringBuilder _value;

            internal LimitedTextBuilder(int limit)
            {
                _limit = Math.Max(0, limit);
                _value = new StringBuilder(Math.Min(_limit, 8192));
            }

            internal bool IsFull
            {
                get { return _value.Length >= _limit; }
            }

            internal bool IsTruncated { get; private set; }

            internal void AppendLine()
            {
                Append(Environment.NewLine);
            }

            internal void AppendLine(string value)
            {
                Append((value ?? string.Empty) + Environment.NewLine);
            }

            public override string ToString()
            {
                return _value.ToString().TrimEnd();
            }

            private void Append(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return;
                }

                int remaining = _limit - _value.Length;
                if (remaining <= 0)
                {
                    IsTruncated = true;
                    return;
                }

                if (value.Length > remaining)
                {
                    _value.Append(value, 0, remaining);
                    IsTruncated = true;
                    return;
                }

                _value.Append(value);
            }
        }
    }

    internal sealed class DocumentExtractionResult
    {
        private DocumentExtractionResult()
        {
        }

        internal string TextContent { get; private set; }

        internal string Status { get; private set; }

        internal bool IsLocalImage { get; private set; }

        internal bool IsTruncated { get; private set; }

        internal static DocumentExtractionResult Text(
            string text,
            string status,
            bool isTruncated)
        {
            return new DocumentExtractionResult
            {
                TextContent = text,
                Status = status,
                IsTruncated = isTruncated
            };
        }

        internal static DocumentExtractionResult Image(string status)
        {
            return new DocumentExtractionResult
            {
                Status = status,
                IsLocalImage = true
            };
        }

        internal static DocumentExtractionResult MetadataOnly(string status)
        {
            return new DocumentExtractionResult
            {
                Status = status
            };
        }
    }

    internal sealed class EmailAttachmentExtractionResult
    {
        private EmailAttachmentExtractionResult()
        {
        }

        internal IReadOnlyList<string> Paths { get; private set; }

        internal int SkippedCount { get; private set; }

        internal bool HadError { get; private set; }

        internal static EmailAttachmentExtractionResult Empty()
        {
            return Success(new List<string>().AsReadOnly(), 0);
        }

        internal static EmailAttachmentExtractionResult Success(
            IReadOnlyList<string> paths,
            int skippedCount)
        {
            return new EmailAttachmentExtractionResult
            {
                Paths = paths,
                SkippedCount = skippedCount
            };
        }

        internal static EmailAttachmentExtractionResult Failure()
        {
            return new EmailAttachmentExtractionResult
            {
                Paths = new List<string>().AsReadOnly(),
                HadError = true
            };
        }
    }
}
