using System.Globalization;
using System.Text;

namespace OutlookAI.RemediationTools;

/// <summary>
/// The bytes of a population's attachments, built by hand so that they are byte-for-byte
/// reproducible on any machine and any runtime.
/// <para>
/// <b>Why the PNG writer does not compress.</b> A PNG's pixel data is a zlib stream, and .NET's
/// deflate implementation has changed underneath the same API before (zlib-ng replaced zlib in .NET
/// 9), so the same pixels could compress to different bytes on a different runtime - and "a
/// from-scratch rebuild reproduces it byte-for-byte" would then be true on one machine only. So the
/// image data goes into STORED deflate blocks (BTYPE 00): valid zlib that any decoder reads, whose
/// every byte is decided here. CRC-32 and Adler-32 are computed here for the same reason.
/// </para>
/// <para>
/// Text kinds are pure ASCII with CRLF line ends, which is what the formats specify and what makes
/// the byte count independent of the platform that wrote them.
/// </para>
/// </summary>
public static class CorpusAttachmentContent
{
    private static readonly byte[] PngSignature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

    private static readonly uint[] CrcTable = BuildCrcTable();

    /// <summary>
    /// An 8-bit RGB PNG of <paramref name="width"/> x <paramref name="height"/> whose channel values
    /// come from <paramref name="pixel"/>(x, y, channel).
    /// </summary>
    public static byte[] Png(int width, int height, Func<int, int, int, byte> pixel)
    {
        ArgumentNullException.ThrowIfNull(pixel);
        if (width < 1 || height < 1 || width > 64 || height > 64)
        {
            throw new ArgumentOutOfRangeException(nameof(width), "A population image is between 1x1 and 64x64.");
        }

        // Raw scanlines: filter byte 0 (None), then RGB triples.
        int stride = 1 + (width * 3);
        byte[] raw = new byte[stride * height];
        for (int y = 0; y < height; y++)
        {
            raw[y * stride] = 0;
            for (int x = 0; x < width; x++)
            {
                for (int c = 0; c < 3; c++)
                {
                    raw[(y * stride) + 1 + (x * 3) + c] = pixel(x, y, c);
                }
            }
        }

        using var png = new MemoryStream();
        png.Write(PngSignature);

        byte[] header = new byte[13];
        WriteUInt32(header, 0, (uint)width);
        WriteUInt32(header, 4, (uint)height);
        header[8] = 8;   // bit depth
        header[9] = 2;   // colour type: truecolour
        header[10] = 0;  // compression: deflate
        header[11] = 0;  // filter: adaptive
        header[12] = 0;  // interlace: none
        WriteChunk(png, "IHDR", header);
        WriteChunk(png, "IDAT", StoredZlib(raw));
        WriteChunk(png, "IEND", Array.Empty<byte>());
        return png.ToArray();
    }

    /// <summary>An iCalendar file carrying one event. <paramref name="uid"/> must be unique per attachment.</summary>
    public static byte[] Calendar(string uid, DateTime startUtc, string summary, string description)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(uid);
        DateTime start = DateTime.SpecifyKind(startUtc, DateTimeKind.Utc);
        string Stamp(DateTime t) => t.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture);
        return Ascii(
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            "PRODID:-//OutlookAI//population//EN",
            "BEGIN:VEVENT",
            "UID:" + uid + "@population.invalid",
            "DTSTAMP:" + Stamp(start.AddDays(-3)),
            "DTSTART:" + Stamp(start),
            "DTEND:" + Stamp(start.AddMinutes(45)),
            "SUMMARY:" + summary,
            "DESCRIPTION:" + description,
            "END:VEVENT",
            "END:VCALENDAR");
    }

    /// <summary>An RFC 822 message, the way a forwarded mail is attached as <c>.eml</c>.</summary>
    public static byte[] Message(
        CorpusCorrespondent from, CorpusCorrespondent to, string subject, DateTime dateUtc, string body)
    {
        ArgumentNullException.ThrowIfNull(from);
        ArgumentNullException.ThrowIfNull(to);
        return Ascii(
            "From: " + from.ToAddressSpec(),
            "To: " + to.ToAddressSpec(),
            "Subject: " + subject,
            "Date: " + DateTime.SpecifyKind(dateUtc, DateTimeKind.Utc).ToString("ddd, dd MMM yyyy HH:mm:ss '+0000'", CultureInfo.InvariantCulture),
            "MIME-Version: 1.0",
            "Content-Type: text/plain; charset=us-ascii",
            string.Empty,
            body);
    }

    /// <summary>A plain-text file.</summary>
    public static byte[] Text(string text) => Ascii(text);

    /// <summary>
    /// CRC-32 as PNG and zlib's gzip wrapper define it (polynomial 0xEDB88320, reflected). Public so
    /// T1 can check a population image the way a decoder would.
    /// </summary>
    public static uint Crc32(ReadOnlySpan<byte> data)
    {
        uint crc = 0xFFFF_FFFFu;
        foreach (byte b in data)
        {
            crc = CrcTable[(crc ^ b) & 0xFF] ^ (crc >> 8);
        }

        return crc ^ 0xFFFF_FFFFu;
    }

    /// <summary>Adler-32, the checksum that closes a zlib stream.</summary>
    public static uint Adler32(ReadOnlySpan<byte> data)
    {
        uint a = 1;
        uint b = 0;
        foreach (byte d in data)
        {
            a = (a + d) % 65521u;
            b = (b + a) % 65521u;
        }

        return (b << 16) | a;
    }

    private static byte[] StoredZlib(byte[] raw)
    {
        using var z = new MemoryStream();
        z.WriteByte(0x78);   // CMF: deflate, 32K window
        z.WriteByte(0x01);   // FLG: no dictionary, fastest; (0x7801 % 31 == 0)
        int offset = 0;
        do
        {
            int length = Math.Min(65_535, raw.Length - offset);
            bool final = offset + length >= raw.Length;
            z.WriteByte(final ? (byte)0x01 : (byte)0x00);   // BFINAL + BTYPE 00, byte-aligned
            z.WriteByte((byte)(length & 0xFF));
            z.WriteByte((byte)(length >> 8));
            z.WriteByte((byte)(~length & 0xFF));
            z.WriteByte((byte)((~length >> 8) & 0xFF));
            z.Write(raw, offset, length);
            offset += length;
        }
        while (offset < raw.Length);

        byte[] adler = new byte[4];
        WriteUInt32(adler, 0, Adler32(raw));
        z.Write(adler);
        return z.ToArray();
    }

    private static void WriteChunk(Stream png, string type, byte[] data)
    {
        byte[] length = new byte[4];
        WriteUInt32(length, 0, (uint)data.Length);
        png.Write(length);
        byte[] typeAndData = new byte[4 + data.Length];
        Encoding.ASCII.GetBytes(type, 0, 4, typeAndData, 0);
        Buffer.BlockCopy(data, 0, typeAndData, 4, data.Length);
        png.Write(typeAndData);
        byte[] crc = new byte[4];
        WriteUInt32(crc, 0, Crc32(typeAndData));
        png.Write(crc);
    }

    private static void WriteUInt32(byte[] target, int offset, uint value)
    {
        target[offset] = (byte)(value >> 24);
        target[offset + 1] = (byte)(value >> 16);
        target[offset + 2] = (byte)(value >> 8);
        target[offset + 3] = (byte)value;
    }

    private static byte[] Ascii(params string[] lines)
    {
        string text = string.Join("\r\n", lines) + "\r\n";
        foreach (char c in text)
        {
            if (c > 0x7E && c != '\r' && c != '\n')
            {
                throw new ArgumentException("Population attachment text must be ASCII.", nameof(lines));
            }
        }

        return Encoding.ASCII.GetBytes(text);
    }

    private static uint[] BuildCrcTable()
    {
        var table = new uint[256];
        for (uint n = 0; n < 256; n++)
        {
            uint c = n;
            for (int k = 0; k < 8; k++)
            {
                c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
            }

            table[n] = c;
        }

        return table;
    }
}
