using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The ONE granted temporary test signature: .htm (with a small PNG resource in its
/// _files directory, exercising Word's native image handling) + .txt (excerpt source).
/// Disposal deletes every file/directory it created. Shared by the signature-steering
/// and signature-placement live suites (D38 test regime: only names prefixed
/// <see cref="SignatureCatalog.TestSignaturePrefix"/> may ever be written).
/// </summary>
internal sealed class TestSignature : IDisposable
{
    // Minimal valid 1x1 transparent PNG.
    private static readonly byte[] TinyPng =
    {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
        0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x62, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
        0x42, 0x60, 0x82,
    };

    private readonly string _directory;
    private readonly string _resourceDir;
    private readonly string _txtPath;

    private TestSignature(string name, string directory)
    {
        Name = name;
        _directory = directory;
        BodyMarker = "OutlookAI MCP testhandtekening " + name.Substring(name.Length - 6);
        FilePath = Path.Combine(directory, name + ".htm");
        _txtPath = Path.Combine(directory, name + ".txt");
        _resourceDir = Path.Combine(directory, name + "_files");

        Directory.CreateDirectory(_resourceDir);
        File.WriteAllBytes(Path.Combine(_resourceDir, "sigimg.png"), TinyPng);
        File.WriteAllText(FilePath,
            "<html><head><meta charset=\"utf-8\"></head><body>"
            + "<p>Met vriendelijke groet,</p>"
            + "<p>" + BodyMarker + "</p>"
            + "<p><img width=\"1\" height=\"1\" src=\"" + name + "_files/sigimg.png\" alt=\"logo\"></p>"
            + "</body></html>");
        File.WriteAllText(_txtPath, "Met vriendelijke groet,\r\n" + BodyMarker + "\r\n");
    }

    public string Name { get; }

    /// <summary>Distinctive text the signature places into a draft body (order asserts).</summary>
    public string BodyMarker { get; }

    /// <summary>The .htm path (what the override inserts).</summary>
    public string FilePath { get; }

    public static TestSignature Create(string runMarker)
    {
        string name = SignatureCatalog.TestSignaturePrefix + "Sig" + runMarker;
        return new TestSignature(name, SignatureCatalog.DefaultSignatureDirectory);
    }

    public void Dispose()
    {
        TryDelete(() => File.Delete(FilePath));
        TryDelete(() => File.Delete(_txtPath));
        TryDelete(() => Directory.Delete(_resourceDir, recursive: true));

        // Belt: nothing with the test prefix may survive this instance (S3).
        foreach (string leftover in Directory.GetFileSystemEntries(_directory, Name + "*"))
        {
            TryDelete(() =>
            {
                if (Directory.Exists(leftover))
                {
                    Directory.Delete(leftover, recursive: true);
                }
                else
                {
                    File.Delete(leftover);
                }
            });
        }
    }

    private static void TryDelete(Action deletion)
    {
        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                deletion();
                return;
            }
            catch (IOException) when (attempt < 2)
            {
                Thread.Sleep(500);
            }
            catch (FileNotFoundException)
            {
                return;
            }
            catch (DirectoryNotFoundException)
            {
                return;
            }
        }
    }
}
