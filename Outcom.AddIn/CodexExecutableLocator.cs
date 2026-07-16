using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Outcom.AddIn
{
    internal static class CodexExecutableLocator
    {
        private const string ExplicitPathEnvironmentVariable = "OUTCOM_CODEX_PATH";

        internal static CodexExecutableInfo Locate()
        {
            string executablePath = FindExecutablePath();
            if (executablePath == null)
            {
                throw new CodexAppServerException(
                    "Codex est introuvable. Installez l'extension Codex officielle pour VS Code " +
                    "ou définissez OUTCOM_CODEX_PATH vers une version validée de codex.exe.");
            }

            return new CodexExecutableInfo
            {
                Path = executablePath,
                Version = ReadVersion(executablePath)
            };
        }

        private static string FindExecutablePath()
        {
            string explicitPath = Environment.GetEnvironmentVariable(
                ExplicitPathEnvironmentVariable,
                EnvironmentVariableTarget.Process);
            string candidate = NormalizeCandidate(explicitPath);
            if (candidate != null)
            {
                return candidate;
            }

            candidate = NormalizeCandidate(Environment.GetEnvironmentVariable(
                ExplicitPathEnvironmentVariable,
                EnvironmentVariableTarget.User));
            if (candidate != null)
            {
                return candidate;
            }

            string localApplicationData = Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData);
            candidate = NormalizeCandidate(Path.Combine(
                localApplicationData,
                "Outcom",
                "CodexRuntime",
                "codex.exe"));
            if (candidate != null)
            {
                return candidate;
            }

            candidate = FindOnPath();
            if (candidate != null)
            {
                return candidate;
            }

            return FindInEditorExtensions();
        }

        private static string FindOnPath()
        {
            string pathValue = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            foreach (string rawDirectory in pathValue.Split(Path.PathSeparator))
            {
                string directory = rawDirectory.Trim().Trim('"');
                if (directory.Length == 0)
                {
                    continue;
                }

                string candidate = NormalizeCandidate(Path.Combine(directory, "codex.exe"));
                if (candidate != null)
                {
                    return candidate;
                }
            }

            return null;
        }

        private static string FindInEditorExtensions()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var extensionRoots = new[]
            {
                Path.Combine(userProfile, ".vscode", "extensions"),
                Path.Combine(userProfile, ".vscode-insiders", "extensions")
            };

            var candidates = new List<FileInfo>();
            string processorArchitecture = Environment.GetEnvironmentVariable(
                "PROCESSOR_ARCHITECTURE") ?? string.Empty;
            string[] runtimeFolders = processorArchitecture.IndexOf(
                "ARM64",
                StringComparison.OrdinalIgnoreCase) >= 0
                    ? new[] { "windows-arm64", "windows-x86_64" }
                    : new[] { "windows-x86_64", "windows-arm64" };
            foreach (string root in extensionRoots)
            {
                try
                {
                    if (!Directory.Exists(root))
                    {
                        continue;
                    }

                    foreach (string extensionDirectory in Directory.EnumerateDirectories(
                        root,
                        "openai.chatgpt-*",
                        SearchOption.TopDirectoryOnly))
                    {
                        foreach (string runtimeFolder in runtimeFolders)
                        {
                            string executablePath = Path.Combine(
                                extensionDirectory,
                                "bin",
                                runtimeFolder,
                                "codex.exe");
                            if (File.Exists(executablePath))
                            {
                                candidates.Add(new FileInfo(executablePath));
                            }
                        }
                    }
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }
            }

            FileInfo newestCandidate = candidates
                .OrderByDescending(file => file.LastWriteTimeUtc)
                .FirstOrDefault();
            return newestCandidate == null ? null : newestCandidate.FullName;
        }

        private static string NormalizeCandidate(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            string candidate = value.Trim().Trim('"');
            if (Directory.Exists(candidate))
            {
                candidate = Path.Combine(candidate, "codex.exe");
            }

            if (!File.Exists(candidate))
            {
                return null;
            }

            return Path.GetFullPath(candidate);
        }

        private static string ReadVersion(string executablePath)
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = executablePath,
                    Arguments = "--version",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (Process process = Process.Start(startInfo))
                {
                    if (process == null)
                    {
                        return "version inconnue";
                    }

                    Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
                    Task<string> errorDrainTask = process.StandardError.ReadToEndAsync();
                    if (!process.WaitForExit(3000))
                    {
                        try
                        {
                            process.Kill();
                            process.WaitForExit(1000);
                        }
                        catch (InvalidOperationException)
                        {
                        }

                        return "version inconnue";
                    }

                    if (!Task.WaitAll(
                        new Task[] { outputTask, errorDrainTask },
                        1000))
                    {
                        return "version inconnue";
                    }

                    string output = outputTask.Result.Trim();
                    return output.Length == 0 ? "version inconnue" : output;
                }
            }
            catch (Exception)
            {
                return "version inconnue";
            }
        }
    }
}
