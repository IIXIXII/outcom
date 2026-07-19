using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace Outcom.AddIn
{
    internal sealed class CodexRuntimeManifest
    {
        private const string ResourceName = "Outcom.AddIn.CodexRuntime.json";
        private static readonly Lazy<CodexRuntimeManifest> CurrentManifest =
            new Lazy<CodexRuntimeManifest>(ReadEmbeddedManifest, true);

        public int SchemaVersion { get; set; }

        public string Version { get; set; }

        public string CliVersion { get; set; }

        public string Platform { get; set; }

        public string ExecutableSha256 { get; set; }

        internal static CodexRuntimeManifest Current
        {
            get { return CurrentManifest.Value; }
        }

        internal static bool IsPinnedInstallationPath(string executablePath)
        {
            if (string.IsNullOrWhiteSpace(executablePath))
            {
                return false;
            }

            string localApplicationData = Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData);
            string userRuntimePath = Path.Combine(
                localApplicationData,
                "Outcom",
                "CodexRuntime",
                "codex.exe");

            string assemblyDirectory = Path.GetDirectoryName(
                typeof(CodexRuntimeManifest).Assembly.Location);
            string packagedRuntimePath = string.IsNullOrWhiteSpace(assemblyDirectory)
                ? string.Empty
                : Path.Combine(assemblyDirectory, "CodexRuntime", "codex.exe");

            return PathsEqual(executablePath, userRuntimePath) ||
                (!string.IsNullOrWhiteSpace(packagedRuntimePath) &&
                 PathsEqual(executablePath, packagedRuntimePath));
        }

        internal void ValidateExecutable(string executablePath, string reportedVersion)
        {
            if (!string.Equals(reportedVersion, CliVersion, StringComparison.Ordinal))
            {
                throw new CodexAppServerException(
                    "Le runtime Codex distribué n'a pas la version attendue. " +
                    "Version requise : " + CliVersion + ".");
            }

            string actualHash;
            try
            {
                actualHash = ComputeSha256(executablePath);
            }
            catch (Exception exception) when (
                exception is IOException ||
                exception is UnauthorizedAccessException ||
                exception is CryptographicException)
            {
                throw new CodexAppServerException(
                    "L'intégrité du runtime Codex distribué n'a pas pu être vérifiée.",
                    exception);
            }

            if (!string.Equals(
                    actualHash,
                    ExecutableSha256,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new CodexAppServerException(
                    "Le runtime Codex distribué ne correspond pas à l'empreinte validée. " +
                    "Réinstallez le runtime épinglé avant d'utiliser Outcom.");
            }
        }

        private static CodexRuntimeManifest ReadEmbeddedManifest()
        {
            Assembly assembly = typeof(CodexRuntimeManifest).Assembly;
            using (Stream stream = assembly.GetManifestResourceStream(ResourceName))
            {
                if (stream == null)
                {
                    throw new CodexAppServerException(
                        "Le manifeste du runtime Codex distribué est absent.");
                }

                using (var reader = new StreamReader(stream, Encoding.UTF8, true))
                {
                    var serializer = new JavaScriptSerializer();
                    CodexRuntimeManifest manifest =
                        serializer.Deserialize<CodexRuntimeManifest>(reader.ReadToEnd());
                    if (manifest == null ||
                        manifest.SchemaVersion != 1 ||
                        string.IsNullOrWhiteSpace(manifest.Version) ||
                        string.IsNullOrWhiteSpace(manifest.CliVersion) ||
                        string.IsNullOrWhiteSpace(manifest.ExecutableSha256))
                    {
                        throw new CodexAppServerException(
                            "Le manifeste du runtime Codex distribué est invalide.");
                    }

                    return manifest;
                }
            }
        }

        private static string ComputeSha256(string filePath)
        {
            using (var stream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                1024 * 1024,
                FileOptions.SequentialScan))
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(stream);
                var result = new StringBuilder(hash.Length * 2);
                foreach (byte value in hash)
                {
                    result.Append(value.ToString("x2"));
                }

                return result.ToString();
            }
        }

        private static bool PathsEqual(string left, string right)
        {
            try
            {
                return string.Equals(
                    Path.GetFullPath(left).TrimEnd(
                        Path.DirectorySeparatorChar,
                        Path.AltDirectorySeparatorChar),
                    Path.GetFullPath(right).TrimEnd(
                        Path.DirectorySeparatorChar,
                        Path.AltDirectorySeparatorChar),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
