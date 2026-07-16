using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;

namespace Outcom.AddIn
{
    /// <summary>
    /// Matérialise les fichiers virtuels exposés par Outlook pendant un glisser-déposer.
    /// </summary>
    internal static class OutlookVirtualFileDrop
    {
        private const string FileGroupDescriptorW = "FileGroupDescriptorW";
        private const string FileGroupDescriptor = "FileGroupDescriptor";
        private const string FileContents = "FileContents";
        private const int MaximumFileCount = 10;
        private const long MaximumFileSize = 25L * 1024L * 1024L;

        internal static bool IsSupported(System.Windows.Forms.IDataObject data)
        {
            if (data == null)
            {
                return false;
            }

            try
            {
                if (HasFileDescriptor(data) || data.GetDataPresent(FileContents, false))
                {
                    return true;
                }

                foreach (string format in data.GetFormats(false))
                {
                    if (format.IndexOf(
                        "RenPrivateMessages",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        internal static IList<string> Extract(System.Windows.Forms.IDataObject data)
        {
            if (!IsSupported(data) || !HasFileDescriptor(data))
            {
                return new List<string>();
            }

            IList<string> names = ReadNames(data);
            if (names.Count > MaximumFileCount)
            {
                throw new InvalidOperationException(
                    "Déposez au maximum " + MaximumFileCount + " documents à la fois.");
            }

            var paths = new List<string>();
            string directory = Path.Combine(
                Path.GetTempPath(),
                "Outcom",
                "Context",
                Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);

            try
            {
                var comData = data as System.Runtime.InteropServices.ComTypes.IDataObject;
                if (comData == null)
                {
                    throw new InvalidOperationException(
                        "Outlook n'a pas fourni les documents dans un format exploitable.");
                }

                for (int index = 0; index < names.Count; index++)
                {
                    string path = Path.Combine(directory, SanitizeFileName(names[index], index));
                    ExtractOne(comData, index, path);
                    paths.Add(path);
                }

                return paths;
            }
            catch
            {
                foreach (string path in paths)
                {
                    try { File.Delete(path); } catch (Exception) { }
                }

                try { Directory.Delete(directory, false); } catch (Exception) { }
                throw;
            }
        }

        private static IList<string> ReadNames(System.Windows.Forms.IDataObject data)
        {
            bool unicode = data.GetDataPresent(FileGroupDescriptorW, false);
            string format = unicode ? FileGroupDescriptorW : FileGroupDescriptor;
            var stream = data.GetData(format, false) as Stream;
            if (stream == null)
            {
                throw new InvalidOperationException(
                    "Outlook n'a pas fourni le nom des documents déposés.");
            }

            byte[] bytes;
            var memoryStream = stream as MemoryStream;
            if (memoryStream != null)
            {
                bytes = memoryStream.ToArray();
            }
            else
            {
                using (var copy = new MemoryStream())
                {
                    if (stream.CanSeek)
                    {
                        stream.Position = 0;
                    }

                    stream.CopyTo(copy);
                    bytes = copy.ToArray();
                }
            }
            if (bytes.Length < sizeof(uint))
            {
                throw new InvalidOperationException("La liste de documents Outlook est invalide.");
            }

            uint count = BitConverter.ToUInt32(bytes, 0);
            Type descriptorType = unicode
                ? typeof(FileDescriptorW)
                : typeof(FileDescriptorA);
            int descriptorSize = Marshal.SizeOf(descriptorType);
            if (count == 0 || count > MaximumFileCount ||
                bytes.Length < sizeof(uint) + descriptorSize * count)
            {
                throw new InvalidOperationException("La liste de documents Outlook est invalide.");
            }

            var names = new List<string>();
            GCHandle handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
            try
            {
                IntPtr current = IntPtr.Add(handle.AddrOfPinnedObject(), sizeof(uint));
                for (int index = 0; index < count; index++)
                {
                    object descriptor = Marshal.PtrToStructure(current, descriptorType);
                    names.Add(unicode
                        ? ((FileDescriptorW)descriptor).FileName
                        : ((FileDescriptorA)descriptor).FileName);
                    current = IntPtr.Add(current, descriptorSize);
                }
            }
            finally
            {
                handle.Free();
            }

            return names;
        }

        private static bool HasFileDescriptor(System.Windows.Forms.IDataObject data)
        {
            return data != null &&
                (data.GetDataPresent(FileGroupDescriptorW, false) ||
                    data.GetDataPresent(FileGroupDescriptor, false));
        }

        private static void ExtractOne(
            System.Runtime.InteropServices.ComTypes.IDataObject data,
            int index,
            string path)
        {
            var format = new FORMATETC
            {
                cfFormat = (short)DataFormats.GetFormat(FileContents).Id,
                dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = index,
                ptd = IntPtr.Zero,
                tymed = TYMED.TYMED_ISTREAM
            };
            STGMEDIUM medium;
            data.GetData(ref format, out medium);
            try
            {
                var source = (IStream)Marshal.GetObjectForIUnknown(medium.unionmember);
                using (var destination = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
                {
                    var buffer = new byte[81920];
                    IntPtr bytesReadPointer = Marshal.AllocCoTaskMem(sizeof(int));
                    long total = 0;
                    try
                    {
                        while (true)
                        {
                            source.Read(buffer, buffer.Length, bytesReadPointer);
                            int read = Marshal.ReadInt32(bytesReadPointer);
                            if (read <= 0)
                            {
                                break;
                            }

                            total += read;
                            if (total > MaximumFileSize)
                            {
                                throw new InvalidOperationException(
                                    "Un document Outlook dépasse la limite de 25 Mo.");
                            }

                            destination.Write(buffer, 0, read);
                        }
                    }
                    finally
                    {
                        Marshal.FreeCoTaskMem(bytesReadPointer);
                        if (Marshal.IsComObject(source))
                        {
                            Marshal.ReleaseComObject(source);
                        }
                    }
                }
            }
            finally
            {
                ReleaseStgMedium(ref medium);
            }
        }

        private static string SanitizeFileName(string value, int index)
        {
            string name = Path.GetFileName(value ?? string.Empty);
            foreach (char invalid in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(invalid, '_');
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                name = "Document Outlook " + (index + 1);
            }

            return index.ToString("00") + "-" + name;
        }

        [DllImport("ole32.dll")]
        private static extern void ReleaseStgMedium(ref STGMEDIUM medium);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode, Pack = 4)]
        private struct FileDescriptorW
        {
            public uint Flags;
            public Guid ClassId;
            public SizeL Size;
            public PointL Point;
            public uint FileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWriteTime;
            public uint FileSizeHigh;
            public uint FileSizeLow;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string FileName;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 4)]
        private struct FileDescriptorA
        {
            public uint Flags;
            public Guid ClassId;
            public SizeL Size;
            public PointL Point;
            public uint FileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWriteTime;
            public uint FileSizeHigh;
            public uint FileSizeLow;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string FileName;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SizeL
        {
            public int Width;
            public int Height;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct PointL
        {
            public int X;
            public int Y;
        }
    }
}
