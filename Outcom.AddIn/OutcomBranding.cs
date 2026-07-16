using System.Drawing;
using System.IO;
using System.Reflection;

namespace Outcom.AddIn
{
    internal static class OutcomBranding
    {
        private const string IconResourceName = "Outcom.AddIn.Assets.Outcom.ico";

        internal static Icon CreateWindowIcon()
        {
            Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(
                IconResourceName);
            if (stream == null)
            {
                LocalLogger.Error("L'icône générale Outcom est absente des ressources embarquées.");
                return (Icon)SystemIcons.Application.Clone();
            }

            using (stream)
            using (var embeddedIcon = new Icon(stream))
            {
                return (Icon)embeddedIcon.Clone();
            }
        }

        internal static Bitmap CreateHeaderImage()
        {
            using (Icon icon = CreateWindowIcon())
            {
                return icon.ToBitmap();
            }
        }
    }
}
