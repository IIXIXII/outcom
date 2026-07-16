using System;
using System.Collections.Generic;

namespace Outcom.AddIn
{
    internal sealed class CodexAccountInfo
    {
        internal bool IsConnected { get; set; }

        internal bool RequiresOpenAiAuth { get; set; }

        internal string AccountType { get; set; }

        internal string Email { get; set; }

        internal string PlanType { get; set; }

        internal static CodexAccountInfo Disconnected(bool requiresOpenAiAuth)
        {
            return new CodexAccountInfo
            {
                IsConnected = false,
                RequiresOpenAiAuth = requiresOpenAiAuth,
                AccountType = null,
                Email = null,
                PlanType = null
            };
        }
    }

    internal sealed class CodexRateLimitInfo
    {
        internal string LimitId { get; set; }

        internal string LimitName { get; set; }

        internal string RateLimitReachedType { get; set; }

        internal CodexRateLimitWindow Primary { get; set; }

        internal CodexRateLimitWindow Secondary { get; set; }
    }

    internal sealed class CodexRateLimitWindow
    {
        internal int UsedPercent { get; set; }

        internal long? WindowDurationMinutes { get; set; }

        internal DateTimeOffset? ResetsAt { get; set; }
    }

    internal sealed class CodexModelInfo
    {
        internal string Id { get; set; }

        internal string DisplayName { get; set; }

        internal bool IsDefault { get; set; }

        internal bool IsHidden { get; set; }

        internal bool SupportsTextInput { get; set; }

        internal string DefaultReasoningEffort { get; set; }

        internal IReadOnlyList<CodexReasoningEffortInfo> SupportedReasoningEfforts { get; set; }

        public override string ToString()
        {
            string name = string.IsNullOrWhiteSpace(DisplayName) ? Id : DisplayName;
            return IsDefault ? name + " (par défaut)" : name;
        }
    }

    internal sealed class CodexReasoningEffortInfo
    {
        internal string Value { get; set; }

        internal string Description { get; set; }

        public override string ToString()
        {
            switch ((Value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "none": return "Aucune";
                case "minimal": return "Minimale";
                case "low": return "Faible";
                case "medium": return "Moyenne";
                case "high": return "Élevée";
                case "xhigh": return "Très élevée";
                default: return string.IsNullOrWhiteSpace(Value) ? "Par défaut du modèle" : Value;
            }
        }
    }

    internal sealed class CodexConnectionStatus
    {
        internal CodexAccountInfo Account { get; set; }

        internal CodexRateLimitInfo RateLimit { get; set; }

        internal CodexModelInfo DefaultModel { get; set; }

        internal string ExecutablePath { get; set; }

        internal string CliVersion { get; set; }

        internal string CodexHomePath { get; set; }

        internal string WorkspacePath { get; set; }
    }

    internal sealed class CodexLoginChallenge
    {
        internal string LoginId { get; set; }

        internal string AuthenticationUrl { get; set; }
    }

    internal sealed class CodexLoginResult
    {
        internal bool Success { get; set; }

        internal string Error { get; set; }
    }

    internal sealed class CodexTurnResult
    {
        internal string Text { get; set; }

        internal string TurnId { get; set; }
    }

    internal sealed class CodexContextFile
    {
        internal string Id { get; set; }

        internal string DisplayName { get; set; }

        internal string SourcePath { get; set; }

        internal string ExtractedText { get; set; }

        internal string ExtractionStatus { get; set; }

        internal bool IsLocalImage { get; set; }
    }

    /// <summary>
    /// Conversation éphémère conservée uniquement en mémoire par un volet Codex.
    /// Elle est volontairement liée au processus app-server qui l'a créée : si ce
    /// processus redémarre, l'utilisateur doit démarrer une nouvelle conversation.
    /// </summary>
    internal sealed class CodexConversationSession
    {
        private int _invalidated;

        internal CodexConversationSession(
            CodexAppServerClient owner,
            string threadId,
            string modelId,
            string reasoningEffort,
            long clientEpoch)
        {
            Owner = owner;
            ThreadId = threadId;
            ModelId = modelId;
            ReasoningEffort = reasoningEffort;
            ClientEpoch = clientEpoch;
            DocumentDirectoryName = Guid.NewGuid().ToString("N");
        }

        internal CodexAppServerClient Owner { get; private set; }

        internal string ThreadId { get; private set; }

        internal string ModelId { get; private set; }

        internal string ReasoningEffort { get; private set; }

        internal long ClientEpoch { get; private set; }

        internal string DocumentDirectoryName { get; private set; }

        internal bool IsValid
        {
            get { return System.Threading.Volatile.Read(ref _invalidated) == 0; }
        }

        internal void Invalidate()
        {
            if (System.Threading.Interlocked.Exchange(ref _invalidated, 1) == 0)
            {
                try
                {
                    Owner.RemoveConversationDocuments(this);
                }
                catch (Exception)
                {
                }
            }
        }
    }

    internal sealed class CodexExecutableInfo
    {
        internal string Path { get; set; }

        internal string Version { get; set; }
    }

    internal sealed class CodexAppServerException : Exception
    {
        internal CodexAppServerException(string message)
            : base(message)
        {
        }

        internal CodexAppServerException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        internal CodexAppServerException(long? code, string message)
            : base(message)
        {
            Code = code;
        }

        internal long? Code { get; private set; }
    }
}
