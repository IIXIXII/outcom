using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace Outcom.AddIn
{
    internal static class LocalLogger
    {
        private static readonly object SyncRoot = new object();
        private static readonly Encoding Utf8WithoutBom = new UTF8Encoding(false);
        private static readonly string LogDirectoryPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Outcom",
            "Logs");

        internal static readonly string LogFilePath = Path.Combine(LogDirectoryPath, "outcom.log");

        internal static void Info(string message)
        {
            Write("INFO", message);
        }

        internal static void Error(string message)
        {
            Write("ERROR", message);
        }

        private static void Write(string level, string message)
        {
            try
            {
                string entry = string.Format(
                    "{0:O} [{1}] {2}{3}",
                    DateTimeOffset.Now,
                    level,
                    message,
                    Environment.NewLine);

                lock (SyncRoot)
                {
                    Directory.CreateDirectory(LogDirectoryPath);
                    File.AppendAllText(LogFilePath, entry, Utf8WithoutBom);
                }
            }
            catch (Exception exception)
            {
                // Une défaillance de journalisation ne doit jamais empêcher le chargement du complément.
                Debug.WriteLine("Impossible d'écrire le journal Outcom : " + exception.Message);
            }
        }
    }
}
