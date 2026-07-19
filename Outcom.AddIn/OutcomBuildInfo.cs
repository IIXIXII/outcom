using System;
using System.Globalization;
using System.IO;
using System.Reflection;

namespace Outcom.AddIn
{
    internal sealed class OutcomBuildInfo
    {
        private const string BuildDateMetadataKey = "BuildDateUtc";
        private const string BuildConfigurationMetadataKey = "BuildConfiguration";

        private OutcomBuildInfo(
            string version,
            DateTime buildDateUtc,
            string assemblyConfiguration)
        {
            Version = version;
            BuildDateUtc = DateTime.SpecifyKind(buildDateUtc, DateTimeKind.Utc);
            AssemblyConfiguration = assemblyConfiguration;
        }

        internal string Version { get; private set; }

        internal DateTime BuildDateUtc { get; private set; }

        internal string AssemblyConfiguration { get; private set; }

        internal string BuildDateDisplay
        {
            get
            {
                return BuildDateUtc
                    .ToLocalTime()
                    .ToString("dd/MM/yyyy HH:mm", CultureInfo.CurrentCulture) +
                    " (heure locale)";
            }
        }

        internal string ProcessArchitecture
        {
            get { return Environment.Is64BitProcess ? "64 bits" : "32 bits"; }
        }

        internal static OutcomBuildInfo Read()
        {
            Assembly assembly = typeof(OutcomBuildInfo).Assembly;
            string version = ReadInformationalVersion(assembly);
            DateTime buildDateUtc = ReadBuildDateUtc(assembly);
            string buildConfiguration = ReadMetadataValue(
                assembly,
                BuildConfigurationMetadataKey);
            var configuration = (AssemblyConfigurationAttribute)Attribute.GetCustomAttribute(
                assembly,
                typeof(AssemblyConfigurationAttribute));

            return new OutcomBuildInfo(
                version,
                buildDateUtc,
                string.IsNullOrWhiteSpace(buildConfiguration) &&
                    (configuration == null ||
                     string.IsNullOrWhiteSpace(configuration.Configuration))
                    ? "Standard"
                    : !string.IsNullOrWhiteSpace(buildConfiguration)
                        ? buildConfiguration
                        : configuration.Configuration);
        }

        private static string ReadInformationalVersion(Assembly assembly)
        {
            var attribute = (AssemblyInformationalVersionAttribute)Attribute.GetCustomAttribute(
                assembly,
                typeof(AssemblyInformationalVersionAttribute));
            if (attribute != null &&
                !string.IsNullOrWhiteSpace(attribute.InformationalVersion))
            {
                return attribute.InformationalVersion.Trim();
            }

            Version assemblyVersion = assembly.GetName().Version;
            return assemblyVersion == null ? "inconnue" : assemblyVersion.ToString();
        }

        private static DateTime ReadBuildDateUtc(Assembly assembly)
        {
            string buildDate = ReadMetadataValue(assembly, BuildDateMetadataKey);
            DateTime parsedDate;
            if (DateTime.TryParse(
                    buildDate,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out parsedDate))
            {
                return parsedDate;
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(assembly.Location) &&
                    File.Exists(assembly.Location))
                {
                    return File.GetLastWriteTimeUtc(assembly.Location);
                }
            }
            catch (Exception)
            {
                // La date du fichier ne sert que de repli pour un binaire ancien.
            }

            return DateTime.MinValue;
        }

        private static string ReadMetadataValue(Assembly assembly, string key)
        {
            object[] metadataAttributes = assembly.GetCustomAttributes(
                typeof(AssemblyMetadataAttribute),
                false);
            foreach (AssemblyMetadataAttribute metadata in metadataAttributes)
            {
                if (string.Equals(
                        metadata.Key,
                        key,
                        StringComparison.Ordinal))
                {
                    return metadata.Value;
                }
            }

            return null;
        }
    }
}
