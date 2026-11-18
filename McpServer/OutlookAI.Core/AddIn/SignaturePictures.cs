using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace OutlookAI.Core.AddIn
{
    /// <summary>Pictures a signature can carry: sniffing, sizes and the limits a new one is held to.</summary>
    public static class SignaturePictures
    {
        /// <summary>The largest picture accepted into a signature: 2 MB. A signature rides on every message.</summary>
        public const int MaxPictureBytes = 2 * 1024 * 1024;

        /// <summary>The picture formats a new picture may be: PNG, JPEG, GIF, BMP.</summary>
        public static readonly IReadOnlyList<string> AcceptedExtensions = Array.AsReadOnly(new[] { ".png", ".jpg", ".jpeg", ".gif", ".bmp" });

        /// <summary>PNG signature bytes.</summary>
        public static bool IsPng(byte[] data)
        {
            return data != null && data.Length >= 8 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47
                && data[4] == 0x0D && data[5] == 0x0A && data[6] == 0x1A && data[7] == 0x0A;
        }

        /// <summary>JPEG start-of-image.</summary>
        public static bool IsJpeg(byte[] data)
        {
            return data != null && data.Length >= 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF;
        }

        /// <summary>GIF87a / GIF89a.</summary>
        public static bool IsGif(byte[] data)
        {
            return data != null && data.Length >= 6 && data[0] == (byte)'G' && data[1] == (byte)'I' && data[2] == (byte)'F'
                && data[3] == (byte)'8' && (data[4] == (byte)'7' || data[4] == (byte)'9') && data[5] == (byte)'a';
        }

        /// <summary>A Windows bitmap.</summary>
        public static bool IsBmp(byte[] data)
        {
            return data != null && data.Length >= 26 && data[0] == (byte)'B' && data[1] == (byte)'M';
        }

        /// <summary>True when the bytes are the kind of picture the extension says.</summary>
        public static bool MatchesExtension(byte[] data, string extension)
        {
            switch ((extension ?? string.Empty).ToLowerInvariant())
            {
                case ".png":
                    return IsPng(data);
                case ".jpg":
                case ".jpeg":
                    return IsJpeg(data);
                case ".gif":
                    return IsGif(data);
                case ".bmp":
                    return IsBmp(data);
                default:
                    return false;
            }
        }

        /// <summary>The MIME type for a picture file name.</summary>
        public static string ContentType(string fileName)
        {
            switch (Path.GetExtension(fileName ?? string.Empty).ToLowerInvariant())
            {
                case ".png":
                    return "image/png";
                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";
                case ".gif":
                    return "image/gif";
                case ".bmp":
                    return "image/bmp";
                case ".svg":
                    return "image/svg+xml";
                case ".webp":
                    return "image/webp";
                case ".tif":
                case ".tiff":
                    return "image/tiff";
                case ".emf":
                case ".emz":
                    return "image/x-emf";
                case ".wmf":
                case ".wmz":
                    return "image/x-wmf";
                case ".xml":
                    return "text/xml";
                case ".thmx":
                    return "application/vnd.ms-officetheme";
                default:
                    return "application/octet-stream";
            }
        }

        /// <summary>The picture's size in pixels from its header (PNG, JPEG, GIF, BMP); false when unreadable.</summary>
        public static bool TryReadSize(byte[] data, out int width, out int height)
        {
            width = 0;
            height = 0;
            if (data == null)
            {
                return false;
            }

            if (IsPng(data) && data.Length >= 24)
            {
                width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
                height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];
                return width > 0 && height > 0;
            }

            if (IsGif(data) && data.Length >= 10)
            {
                width = data[6] | (data[7] << 8);
                height = data[8] | (data[9] << 8);
                return width > 0 && height > 0;
            }

            if (IsBmp(data))
            {
                width = BitConverter.ToInt32(data, 18);
                height = Math.Abs(BitConverter.ToInt32(data, 22));
                return width > 0 && height > 0;
            }

            if (IsJpeg(data))
            {
                int i = 2;
                while (i + 9 < data.Length)
                {
                    if (data[i] != 0xFF)
                    {
                        i++;
                        continue;
                    }

                    byte marker = data[i + 1];
                    int length = (data[i + 2] << 8) | data[i + 3];
                    bool startOfFrame = marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC;
                    if (startOfFrame)
                    {
                        height = (data[i + 5] << 8) | data[i + 6];
                        width = (data[i + 7] << 8) | data[i + 8];
                        return width > 0 && height > 0;
                    }

                    if (length < 2)
                    {
                        return false;
                    }

                    i += 2 + length;
                }
            }

            return false;
        }

        /// <summary>
        /// Reads a picture an agent named by absolute path, held to the rules every new picture is:
        /// an existing file, not empty, at most <see cref="MaxPictureBytes"/>, a PNG, JPEG, GIF or BMP
        /// whose bytes are what its extension says. Throws ArgumentException naming the path and the
        /// rule broken; nothing has been written when it does.
        /// </summary>
        public static byte[] ReadNewPicture(string path, string targetFileName)
        {
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path))
            {
                throw new ArgumentException("picture path '" + path + "' must be an absolute path to a file on this machine. Nothing was written.");
            }

            if (Directory.Exists(path))
            {
                throw new ArgumentException("picture path '" + path + "' is a folder, not a file. Nothing was written.");
            }

            if (!File.Exists(path))
            {
                throw new ArgumentException("picture file '" + path + "' does not exist. Nothing was written.");
            }

            long length = new FileInfo(path).Length;
            if (length == 0)
            {
                throw new ArgumentException("picture file '" + path + "' is empty. Nothing was written.");
            }

            if (length > MaxPictureBytes)
            {
                throw new ArgumentException("picture file '" + path + "' is " + length.ToString(CultureInfo.InvariantCulture)
                    + " bytes; a signature picture may be at most " + MaxPictureBytes.ToString(CultureInfo.InvariantCulture)
                    + " bytes, because it travels with every message. Nothing was written.");
            }

            string extension = Path.GetExtension(targetFileName);
            if (!AcceptedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
            {
                throw new ArgumentException("picture '" + targetFileName + "' must be a .png, .jpg, .jpeg, .gif or .bmp file. Nothing was written.");
            }

            byte[] data = File.ReadAllBytes(path);
            if (!MatchesExtension(data, extension))
            {
                throw new ArgumentException("picture file '" + path + "' is not a " + extension.TrimStart('.').ToUpperInvariant()
                    + " picture, which its target name '" + targetFileName + "' says it is. Name it after its real format. Nothing was written.");
            }

            return data;
        }
    }
}
