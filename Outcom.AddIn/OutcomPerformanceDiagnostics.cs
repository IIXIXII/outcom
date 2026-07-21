#if DEBUG
using System;
using System.Threading;
using Microsoft.Win32;

namespace Outcom.AddIn
{
    internal static class OutcomPerformanceDiagnostics
    {
        private const string DiagnosticsRegistryPath = @"Software\Outcom\Diagnostics";
        private const string StartupDelayValueName = "StartupDelayMilliseconds";
        private const int MaximumStartupDelayMilliseconds = 15000;

        internal static void ApplyStartupDelayIfConfigured()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(DiagnosticsRegistryPath))
                {
                    object configuredValue = key == null
                        ? null
                        : key.GetValue(StartupDelayValueName, null);

                    int delayMilliseconds;
                    if (configuredValue == null ||
                        !int.TryParse(Convert.ToString(configuredValue), out delayMilliseconds) ||
                        delayMilliseconds <= 0)
                    {
                        return;
                    }

                    delayMilliseconds = Math.Min(
                        delayMilliseconds,
                        MaximumStartupDelayMilliseconds);

                    LocalLogger.Info(
                        "Test de resilience Outlook : delai de demarrage Debug de " +
                        delayMilliseconds +
                        " ms.");
                    Thread.Sleep(delayMilliseconds);
                }
            }
            catch (Exception exception)
            {
                LocalLogger.Error(
                    "Impossible d'appliquer le delai du test de resilience Outlook (" +
                    exception.GetType().Name +
                    ").");
            }
        }
    }
}
#endif
