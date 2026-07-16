using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outcom.AddIn
{
    /// <summary>
    /// Métadonnées minimales permettant de rouvrir le courrier source ou de nommer
    /// un brouillon. Le contenu du fichier déposé est préparé séparément par
    /// DocumentContextExtractor.
    /// </summary>
    internal sealed class OutlookMailContext
    {
        private OutlookMailContext()
        {
        }

        internal string Subject { get; private set; }

        internal string SenderName { get; private set; }

        internal string ReceivedAt { get; private set; }

        internal string EntryId { get; private set; }

        internal string StoreId { get; private set; }

        internal string DisplayName
        {
            get
            {
                string subject = NormalizeSingleLine(Subject);
                if (string.IsNullOrWhiteSpace(subject))
                {
                    return "(Sans objet)";
                }

                return subject.Length <= 80 ? subject : subject.Substring(0, 77) + "…";
            }
        }

        internal string Details
        {
            get
            {
                string sender = string.IsNullOrWhiteSpace(SenderName)
                    ? "Expéditeur inconnu"
                    : NormalizeSingleLine(SenderName);
                return sender + " — " + ReceivedAt;
            }
        }

        internal static OutlookMailContext Capture(Outlook.Explorer explorer)
        {
            if (explorer == null)
            {
                throw new InvalidOperationException("Aucune fenêtre Outlook n'est disponible.");
            }

            Outlook.Selection selection = null;
            object selectedItem = null;
            Outlook.MAPIFolder parentFolder = null;
            try
            {
                selection = explorer.Selection;
                if (selection == null || selection.Count != 1)
                {
                    throw new InvalidOperationException(
                        "Sélectionnez exactement un courrier Outlook, puis réessayez.");
                }

                selectedItem = selection[1];
                var mail = selectedItem as Outlook.MailItem;
                if (mail == null)
                {
                    throw new InvalidOperationException(
                        "L'élément sélectionné n'est pas un courrier Outlook.");
                }

                parentFolder = mail.Parent as Outlook.MAPIFolder;

                return new OutlookMailContext
                {
                    Subject = NormalizeSingleLine(mail.Subject),
                    SenderName = NormalizeSingleLine(mail.SenderName),
                    ReceivedAt = mail.ReceivedTime.ToString(
                        "yyyy-MM-dd HH:mm",
                        CultureInfo.InvariantCulture),
                    EntryId = mail.EntryID,
                    StoreId = parentFolder == null ? null : parentFolder.StoreID
                };
            }
            catch (COMException exception)
            {
                throw new InvalidOperationException(
                    "Outlook n'a pas pu lire le courrier sélectionné.",
                    exception);
            }
            finally
            {
                ReleaseComObject(parentFolder);
                ReleaseComObject(selectedItem);
                ReleaseComObject(selection);
            }
        }

        internal static IList<string> SaveSelectionAsNativeMessages(
            Outlook.Explorer explorer)
        {
            if (explorer == null)
            {
                throw new InvalidOperationException("Aucune fenêtre Outlook n'est disponible.");
            }

            Outlook.Selection selection = null;
            var paths = new List<string>();
            string directory = Path.Combine(
                Path.GetTempPath(),
                "Outcom",
                "Context",
                Guid.NewGuid().ToString("N"));
            try
            {
                selection = explorer.Selection;
                if (selection == null || selection.Count == 0)
                {
                    throw new InvalidOperationException(
                        "Aucun courrier Outlook n'est sélectionné.");
                }

                if (selection.Count > 10)
                {
                    throw new InvalidOperationException(
                        "Déposez au maximum 10 courriers Outlook à la fois.");
                }

                Directory.CreateDirectory(directory);
                for (int index = 1; index <= selection.Count; index++)
                {
                    object selectedItem = null;
                    try
                    {
                        selectedItem = selection[index];
                        var mail = selectedItem as Outlook.MailItem;
                        if (mail == null)
                        {
                            throw new InvalidOperationException(
                                "La sélection Outlook contient un élément qui n'est pas un courrier.");
                        }

                        string subject = NormalizeSingleLine(mail.Subject);
                        if (string.IsNullOrWhiteSpace(subject))
                        {
                            subject = "Sans objet";
                        }

                        string path = Path.Combine(
                            directory,
                            index.ToString("00", CultureInfo.InvariantCulture) + "-" +
                                SanitizeFileName(subject) + ".msg");
                        paths.Add(path);
                        mail.SaveAs(path, Outlook.OlSaveAsType.olMSGUnicode);
                    }
                    finally
                    {
                        ReleaseComObject(selectedItem);
                    }
                }

                return paths.AsReadOnly();
            }
            catch (COMException exception)
            {
                DeleteSavedMessages(paths, directory);
                throw new InvalidOperationException(
                    "Outlook n'a pas pu enregistrer le courrier déposé.",
                    exception);
            }
            catch
            {
                DeleteSavedMessages(paths, directory);
                throw;
            }
            finally
            {
                ReleaseComObject(selection);
            }
        }

        internal void Open(Outlook.Application application)
        {
            if (application == null)
            {
                throw new ArgumentNullException(nameof(application));
            }

            if (string.IsNullOrWhiteSpace(EntryId))
            {
                throw new InvalidOperationException("Le courrier Outlook n'est plus accessible.");
            }

            object item = null;
            try
            {
                item = application.Session.GetItemFromID(EntryId, StoreId);
                var mail = item as Outlook.MailItem;
                if (mail == null)
                {
                    throw new InvalidOperationException("L'élément Outlook n'est pas un courrier.");
                }

                mail.Display(false);
            }
            catch (COMException exception)
            {
                throw new InvalidOperationException(
                    "Outlook n'a pas pu ouvrir le courrier.",
                    exception);
            }
            finally
            {
                ReleaseComObject(item);
            }
        }

        /// <returns>
        /// true si le brouillon a été enregistré et sa fenêtre ouverte ; false s'il
        /// est bien enregistré dans Brouillons mais que l'ouverture a échoué.
        /// </returns>
        internal bool CreatePlainTextDraft(
            Outlook.Application application,
            string generatedText)
        {
            if (application == null)
            {
                throw new ArgumentNullException(nameof(application));
            }

            if (string.IsNullOrWhiteSpace(generatedText))
            {
                throw new InvalidOperationException("Aucune réponse Codex n'est disponible.");
            }

            Outlook.MailItem draft = null;
            bool saved = false;
            try
            {
                draft = (Outlook.MailItem)application.CreateItem(
                    Outlook.OlItemType.olMailItem);
                draft.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
                draft.Subject = BuildDraftSubject(Subject);
                draft.Body = generatedText.Trim();
                draft.Save();
                saved = true;
                draft.Display(false);
                return true;
            }
            catch (COMException exception)
            {
                if (saved)
                {
                    // Ne pas présenter cette situation comme un échec de création :
                    // un nouvel essai produirait un doublon dans le dossier Brouillons.
                    return false;
                }

                if (!saved && draft != null)
                {
                    try
                    {
                        draft.Close(Outlook.OlInspectorClose.olDiscard);
                    }
                    catch (COMException)
                    {
                    }
                }

                throw new InvalidOperationException(
                    "Outlook n'a pas pu créer le brouillon.",
                    exception);
            }
            finally
            {
                ReleaseComObject(draft);
            }
        }

        private static string BuildDraftSubject(string sourceSubject)
        {
            string subject = NormalizeSingleLine(sourceSubject);
            if (string.IsNullOrWhiteSpace(subject))
            {
                subject = "Sans objet";
            }

            const string prefix = "Brouillon Codex — ";
            int maximumSourceLength = 255 - prefix.Length;
            if (subject.Length > maximumSourceLength)
            {
                subject = subject.Substring(0, maximumSourceLength);
            }

            return prefix + subject;
        }

        private static string NormalizeSingleLine(string value)
        {
            return (value ?? string.Empty)
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\0', ' ')
                .Trim();
        }

        private static string SanitizeFileName(string value)
        {
            string name = value ?? string.Empty;
            foreach (char invalid in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(invalid, '_');
            }

            name = name.Trim().TrimEnd('.');
            if (string.IsNullOrWhiteSpace(name))
            {
                name = "Courrier Outlook";
            }

            return name.Length <= 120 ? name : name.Substring(0, 120);
        }

        private static void DeleteSavedMessages(IEnumerable<string> paths, string directory)
        {
            foreach (string path in paths)
            {
                try { File.Delete(path); } catch (Exception) { }
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
                {
                    Directory.Delete(directory, false);
                }
            }
            catch (Exception)
            {
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
