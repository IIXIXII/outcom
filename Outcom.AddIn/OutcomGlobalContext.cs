using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace Outcom.AddIn
{
    /// <summary>
    /// Directives durables définies explicitement par l'utilisateur. Elles sont
    /// distinctes du contenu Outlook, qui reste une donnée non fiable jointe au cas par cas.
    /// </summary>
    internal sealed class OutcomGlobalContext
    {
        internal const int MaximumSectionLength = 12000;
        internal const int MaximumTotalLength = 30000;
        internal const int MaximumModelIdLength = 200;
        internal const int MaximumReasoningEffortLength = 64;

        public string WorkContext { get; set; }

        public string VocabularyGuidelines { get; set; }

        public string CrossConversationInstructions { get; set; }

        public string ModelId { get; set; }

        public string ReasoningEffort { get; set; }

        internal bool IsEmpty
        {
            get
            {
                return string.IsNullOrWhiteSpace(WorkContext) &&
                    string.IsNullOrWhiteSpace(VocabularyGuidelines) &&
                    string.IsNullOrWhiteSpace(CrossConversationInstructions) &&
                    string.IsNullOrWhiteSpace(ModelId) &&
                    string.IsNullOrWhiteSpace(ReasoningEffort);
            }
        }

        internal int SectionCount
        {
            get
            {
                int count = 0;
                if (!string.IsNullOrWhiteSpace(WorkContext))
                {
                    count++;
                }

                if (!string.IsNullOrWhiteSpace(VocabularyGuidelines))
                {
                    count++;
                }

                if (!string.IsNullOrWhiteSpace(CrossConversationInstructions))
                {
                    count++;
                }

                return count;
            }
        }

        internal OutcomGlobalContext Clone()
        {
            return new OutcomGlobalContext
            {
                WorkContext = WorkContext,
                VocabularyGuidelines = VocabularyGuidelines,
                CrossConversationInstructions = CrossConversationInstructions,
                ModelId = ModelId,
                ReasoningEffort = ReasoningEffort
            };
        }

        internal bool ContentEquals(OutcomGlobalContext other)
        {
            return other != null &&
                string.Equals(WorkContext, other.WorkContext, StringComparison.Ordinal) &&
                string.Equals(
                    VocabularyGuidelines,
                    other.VocabularyGuidelines,
                    StringComparison.Ordinal) &&
                string.Equals(
                    CrossConversationInstructions,
                    other.CrossConversationInstructions,
                    StringComparison.Ordinal) &&
                string.Equals(ModelId, other.ModelId, StringComparison.Ordinal) &&
                string.Equals(ReasoningEffort, other.ReasoningEffort, StringComparison.Ordinal);
        }

        internal static OutcomGlobalContext ValidateAndNormalize(OutcomGlobalContext value)
        {
            OutcomGlobalContext context = value ?? new OutcomGlobalContext();
            string workContext = Normalize(context.WorkContext);
            string vocabulary = Normalize(context.VocabularyGuidelines);
            string instructions = Normalize(context.CrossConversationInstructions);
            string modelId = Normalize(context.ModelId);
            string reasoningEffort = Normalize(context.ReasoningEffort);

            ValidateSection("Contexte de travail", workContext);
            ValidateSection("Vocabulaire", vocabulary);
            ValidateSection("Instructions transversales", instructions);
            ValidateIdentifier("modèle", modelId, MaximumModelIdLength);
            ValidateIdentifier(
                "profondeur de raisonnement",
                reasoningEffort,
                MaximumReasoningEffortLength);

            long totalLength = (long)workContext.Length + vocabulary.Length + instructions.Length;
            if (totalLength > MaximumTotalLength)
            {
                throw new ArgumentException(
                    "Le contexte Codex dépasse la limite totale de " +
                    MaximumTotalLength + " caractères.");
            }

            return new OutcomGlobalContext
            {
                WorkContext = workContext,
                VocabularyGuidelines = vocabulary,
                CrossConversationInstructions = instructions,
                ModelId = modelId,
                ReasoningEffort = reasoningEffort
            };
        }

        private static string Normalize(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static void ValidateSection(string name, string value)
        {
            if (value.Length > MaximumSectionLength)
            {
                throw new ArgumentException(
                    "La section « " + name + " » dépasse la limite de " +
                    MaximumSectionLength + " caractères.");
            }
        }

        private static void ValidateIdentifier(string name, string value, int maximumLength)
        {
            if (value.Length > maximumLength)
            {
                throw new ArgumentException(
                    "Le " + name + " dépasse la limite de " + maximumLength + " caractères.");
            }

            foreach (char character in value)
            {
                if (char.IsControl(character))
                {
                    throw new ArgumentException("Le " + name + " contient un caractère invalide.");
                }
            }
        }
    }

    internal sealed class OutcomGlobalContextChangedEventArgs : EventArgs
    {
        internal OutcomGlobalContextChangedEventArgs(OutcomGlobalContext context)
        {
            Context = context == null ? new OutcomGlobalContext() : context.Clone();
        }

        internal OutcomGlobalContext Context { get; private set; }
    }

    /// <summary>
    /// Stockage local chiffré par DPAPI pour le compte Windows courant. Le fichier ne
    /// contient jamais de texte exploitable sans la session Windows de l'utilisateur.
    /// </summary>
    internal sealed class OutcomGlobalContextStore
    {
        private const int CurrentFormatVersion = 1;
        private static readonly Encoding Utf8WithoutBom = new UTF8Encoding(false);
        private static readonly byte[] OptionalEntropy = Utf8WithoutBom.GetBytes(
            "Outcom.GlobalContext.v1");

        private readonly string _filePath;

        internal OutcomGlobalContextStore()
            : this(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Outcom",
                "Context",
                "global-context.dat"))
        {
        }

        internal OutcomGlobalContextStore(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("Le chemin du contexte Codex est vide.", nameof(filePath));
            }

            _filePath = Path.GetFullPath(filePath);
        }

        internal string FilePath
        {
            get { return _filePath; }
        }

        internal OutcomGlobalContext Load()
        {
            if (!File.Exists(_filePath))
            {
                return new OutcomGlobalContext();
            }

            byte[] protectedBytes = null;
            byte[] clearBytes = null;
            try
            {
                protectedBytes = File.ReadAllBytes(_filePath);
                if (protectedBytes.Length == 0)
                {
                    throw new InvalidDataException("Le fichier de contexte est vide.");
                }

                clearBytes = ProtectedData.Unprotect(
                    protectedBytes,
                    OptionalEntropy,
                    DataProtectionScope.CurrentUser);
                string json = Utf8WithoutBom.GetString(clearBytes);
                var serializer = new JavaScriptSerializer
                {
                    MaxJsonLength = OutcomGlobalContext.MaximumTotalLength * 8 + 4096
                };
                OutcomGlobalContextDocument document =
                    serializer.Deserialize<OutcomGlobalContextDocument>(json);
                if (document == null || document.Version != CurrentFormatVersion)
                {
                    throw new InvalidDataException(
                        "La version du fichier de contexte n'est pas reconnue.");
                }

                return OutcomGlobalContext.ValidateAndNormalize(document.Context);
            }
            catch (Exception exception) when (
                exception is IOException ||
                exception is UnauthorizedAccessException ||
                exception is CryptographicException ||
                exception is InvalidDataException ||
                exception is InvalidOperationException ||
                exception is ArgumentException)
            {
                LocalLogger.Error(
                    "Impossible de lire le contexte Codex (" +
                    exception.GetType().Name + ").");
                return new OutcomGlobalContext();
            }
            finally
            {
                ClearBytes(clearBytes);
                ClearBytes(protectedBytes);
            }
        }

        internal void Save(OutcomGlobalContext value)
        {
            OutcomGlobalContext context = OutcomGlobalContext.ValidateAndNormalize(value);
            try
            {
                if (context.IsEmpty)
                {
                    if (File.Exists(_filePath))
                    {
                        File.Delete(_filePath);
                    }

                    return;
                }

                var serializer = new JavaScriptSerializer
                {
                    MaxJsonLength = OutcomGlobalContext.MaximumTotalLength * 8 + 4096
                };
                string json = serializer.Serialize(new OutcomGlobalContextDocument
                {
                    Version = CurrentFormatVersion,
                    Context = context
                });
                byte[] clearBytes = Utf8WithoutBom.GetBytes(json);
                byte[] protectedBytes = null;
                string temporaryPath = _filePath + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try
                {
                    protectedBytes = ProtectedData.Protect(
                        clearBytes,
                        OptionalEntropy,
                        DataProtectionScope.CurrentUser);
                    string directoryPath = Path.GetDirectoryName(_filePath);
                    if (string.IsNullOrWhiteSpace(directoryPath))
                    {
                        throw new InvalidOperationException(
                            "Le dossier du contexte Codex est invalide.");
                    }

                    Directory.CreateDirectory(directoryPath);
                    File.WriteAllBytes(temporaryPath, protectedBytes);
                    if (File.Exists(_filePath))
                    {
                        File.Replace(temporaryPath, _filePath, null, true);
                    }
                    else
                    {
                        File.Move(temporaryPath, _filePath);
                    }
                }
                finally
                {
                    ClearBytes(clearBytes);
                    ClearBytes(protectedBytes);
                    if (File.Exists(temporaryPath))
                    {
                        File.Delete(temporaryPath);
                    }
                }
            }
            catch (Exception exception) when (
                exception is IOException ||
                exception is UnauthorizedAccessException ||
                exception is CryptographicException ||
                exception is InvalidOperationException)
            {
                throw new InvalidOperationException(
                    "Le contexte Codex n'a pas pu être enregistré localement.",
                    exception);
            }
        }

        private static void ClearBytes(byte[] value)
        {
            if (value != null)
            {
                Array.Clear(value, 0, value.Length);
            }
        }

        private sealed class OutcomGlobalContextDocument
        {
            public OutcomGlobalContextDocument()
            {
            }

            public int Version { get; set; }

            public OutcomGlobalContext Context { get; set; }
        }
    }
}
