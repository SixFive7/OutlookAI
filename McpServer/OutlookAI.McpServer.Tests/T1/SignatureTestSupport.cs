using System.Text;
using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Fakes for the signature model's two registry seams - the settings (folder value, policy, roaming,
/// Click-to-Run) and the profiles' account lists - in the names and shapes Office writes, plus a
/// temporary signature folder. Nothing here touches the real registry or the real Signatures folder.
/// </summary>
internal sealed class FakeSignatureSettings : ISignatureSettingsRegistry
{
    public const string Office = "16.0";

    private readonly Dictionary<(SignatureRegistryHive, string, string), object> _values = new();

    /// <summary>
    /// Office LTSC 2021 by default - a build that predates roaming, so roaming is OFF and a suffixed
    /// name is just a name. Tests that need roaming say so.
    /// </summary>
    public FakeSignatureSettings()
    {
        WithBuild("16.0.14334.20848", "ProPlus2021Volume");
    }

    public FakeSignatureSettings Set(SignatureRegistryHive hive, string key, string value, object data)
    {
        _values[(hive, key.ToLowerInvariant(), value.ToLowerInvariant())] = data;
        return this;
    }

    public FakeSignatureSettings Clear(SignatureRegistryHive hive, string key, string value)
    {
        _values.Remove((hive, key.ToLowerInvariant(), value.ToLowerInvariant()));
        return this;
    }

    public FakeSignatureSettings WithBuild(string version, string product)
    {
        Set(SignatureRegistryHive.LocalMachine, OfficeBuildInfo.ClickToRunConfigurationKey, "VersionToReport", version);
        return Set(SignatureRegistryHive.LocalMachine, OfficeBuildInfo.ClickToRunConfigurationKey, "ProductReleaseIds", product);
    }

    /// <summary>Microsoft 365 Apps on a roaming build.</summary>
    public FakeSignatureSettings Microsoft365() => WithBuild("16.0.18129.20158", "O365ProPlusRetail");

    /// <summary>Office LTSC 2024, the testbed's: it roams where policy allows it (E2's R3 run 2), so off unless Outlook's own copies show it.</summary>
    public FakeSignatureSettings Ltsc2024() => WithBuild("16.0.17932.20884", "ProPlus2024Volume");

    public FakeSignatureSettings RoamingSwitchedOff(bool policy = false, bool temporaryToggleName = false)
    {
        string key = (policy ? @"Software\Policies\Microsoft\Office\" : @"Software\Microsoft\Office\") + Office + @"\Outlook\Setup";
        return Set(SignatureRegistryHive.CurrentUser, key,
            temporaryToggleName ? SignatureRoaming.DisableTemporaryToggleValueName : SignatureRoaming.DisableValueName, 1);
    }

    public FakeSignatureSettings Policy(SignaturePolicySource source, bool? disable = null, string? newMessage = null, string? reply = null)
    {
        string key = SignaturePolicy.KeyPath(source, Office);
        if (disable != null)
        {
            Set(SignatureRegistryHive.CurrentUser, key, SignaturePolicy.DisableSignaturesValueName, disable.Value ? 1 : 0);
        }

        if (newMessage != null)
        {
            Set(SignatureRegistryHive.CurrentUser, key, SignaturePolicy.NewSignatureValueName, newMessage);
        }

        if (reply != null)
        {
            Set(SignatureRegistryHive.CurrentUser, key, SignaturePolicy.ReplySignatureValueName, reply);
        }

        return this;
    }

    public FakeSignatureSettings FolderValue(string value)
    {
        return Set(SignatureRegistryHive.CurrentUser, @"Software\Microsoft\Office\" + Office + @"\Common\General", "Signatures", value);
    }

    public object? ReadValue(SignatureRegistryHive hive, string keyPath, string valueName)
    {
        return _values.TryGetValue((hive, keyPath.ToLowerInvariant(), valueName.ToLowerInvariant()), out object? value) ? value : null;
    }
}

/// <summary>
/// Every profile's account-manager key in memory, keyed the way the live reader keys entries
/// (<c>Software\Microsoft\Office\16.0\Outlook\Profiles\&lt;profile&gt;\9375CFF0...\0000000n</c>), with the
/// recorded default profile. Records every write.
/// </summary>
internal sealed class FakeProfileAccountRegistry : IProfileAccountRegistry
{
    public const string Pop3Clsid = "{ED475411-B0D6-11D2-8C3B-00104B2A6676}";
    public const string Imap4Clsid = "{ED475412-B0D6-11D2-8C3B-00104B2A6676}";
    public const string MapiServiceClsid = "{ED475414-B0D6-11D2-8C3B-00104B2A6676}";

    private readonly List<(string Key, string Profile, Dictionary<string, object?> Values)> _entries = new();
    private readonly List<string> _profiles = new();

    public string? DefaultProfile { get; set; }

    public List<(string Key, string ValueName, string Value)> Sets { get; } = new();

    public List<(string Key, string ValueName)> Deletes { get; } = new();

    public static string KeyFor(string profile, int entry)
    {
        return OutlookProfileRegistry.BuildOutlookRootKeyPath(FakeSignatureSettings.Office) + "\\"
            + OutlookProfileRegistry.ProfilesSubKeyName + "\\" + profile + "\\" + OutlookProfileRegistry.AccountsSubKeyName
            + "\\" + entry.ToString("D8", System.Globalization.CultureInfo.InvariantCulture);
    }

    public FakeProfileAccountRegistry Profile(string name, bool isDefault = false)
    {
        if (!_profiles.Contains(name, StringComparer.OrdinalIgnoreCase))
        {
            _profiles.Add(name);
        }

        if (isDefault)
        {
            DefaultProfile = name;
        }

        return this;
    }

    public FakeProfileAccountRegistry Add(string profile, int entry, Dictionary<string, object?> values)
    {
        Profile(profile);
        _entries.Add((KeyFor(profile, entry), profile, new Dictionary<string, object?>(values, StringComparer.OrdinalIgnoreCase)));
        return this;
    }

    public object? Value(string key, string valueName)
    {
        Dictionary<string, object?>? entry = Find(key);
        return entry != null && entry.TryGetValue(valueName, out object? value) ? value : null;
    }

    public IReadOnlyList<ProfileRegistryEntry> ReadEntries()
    {
        return _entries
            .Select(e => new ProfileRegistryEntry(e.Key, new Dictionary<string, object?>(e.Values, StringComparer.OrdinalIgnoreCase), e.Profile))
            .ToList();
    }

    public ProfileList ReadProfiles()
    {
        return new ProfileList(DefaultProfile, _profiles.OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToList());
    }

    public void SetString(string key, string valueName, string value)
    {
        Sets.Add((key, valueName, value));
        Dictionary<string, object?> entry = Find(key) ?? throw new InvalidOperationException("no such entry: " + key);
        entry[valueName] = value;
    }

    public void DeleteValue(string key, string valueName)
    {
        Deletes.Add((key, valueName));
        Find(key)?.Remove(valueName);
    }

    private Dictionary<string, object?>? Find(string key)
    {
        return _entries.FirstOrDefault(e => string.Equals(e.Key, key, StringComparison.OrdinalIgnoreCase)).Values;
    }

    public static Dictionary<string, object?> Pop3(string email, string? newSignature = null, string? replySignature = null, string? accountName = null)
    {
        return With(new Dictionary<string, object?>
        {
            ["clsid"] = Pop3Clsid,
            ["Account Name"] = accountName ?? email,
            ["Email"] = email,
        }, newSignature, replySignature);
    }

    public static Dictionary<string, object?> Imap(string email, string? newSignature = null, string? replySignature = null)
    {
        return With(new Dictionary<string, object?>
        {
            ["clsid"] = Imap4Clsid,
            ["Account Name"] = email,
            ["Email"] = email,
        }, newSignature, replySignature);
    }

    /// <summary>An Exchange account as measured: no Email value, the Account Name is the address it was added under.</summary>
    public static Dictionary<string, object?> Exchange(string accountName, string? newSignature = null, string? replySignature = null)
    {
        return With(new Dictionary<string, object?>
        {
            ["clsid"] = MapiServiceClsid,
            ["Service Name"] = "MSEMS",
            ["Account Name"] = accountName,
        }, newSignature, replySignature);
    }

    public static Dictionary<string, object?> DataFile(string storeName)
    {
        return new Dictionary<string, object?>
        {
            ["clsid"] = MapiServiceClsid,
            ["Service Name"] = "MSUPST MS",
            ["Account Name"] = storeName,
        };
    }

    private static Dictionary<string, object?> With(Dictionary<string, object?> values, string? newSignature, string? replySignature)
    {
        if (newSignature != null)
        {
            // REG_BINARY UTF-16, NUL-terminated - what Outlook itself writes.
            values["New Signature"] = Encoding.Unicode.GetBytes(newSignature + "\0");
        }

        if (replySignature != null)
        {
            values["Reply-Forward Signature"] = Encoding.Unicode.GetBytes(replySignature + "\0");
        }

        return values;
    }
}

/// <summary>A per-test signature folder with a backup root beside it, deleted on dispose.</summary>
internal sealed class TempSignatureFolder : IDisposable
{
    public TempSignatureFolder()
    {
        Root = Path.Combine(Path.GetTempPath(), "OutlookAI-SigTests-" + Guid.NewGuid().ToString("N"));
        Directory = Path.Combine(Root, "Signatures");
        BackupRoot = Path.Combine(Root, "backups");
        System.IO.Directory.CreateDirectory(Directory);
    }

    public string Root { get; }

    public string Directory { get; }

    public string BackupRoot { get; }

    public string PathOf(string fileName) => System.IO.Path.Combine(Directory, fileName);

    public string Write(string fileName, string content, Encoding encoding)
    {
        string path = PathOf(fileName);
        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, encoding.GetPreamble().Concat(encoding.GetBytes(content)).ToArray());
        return path;
    }

    public string WriteBytes(string fileName, byte[] content)
    {
        string path = PathOf(fileName);
        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, content);
        return path;
    }

    /// <summary>A loose file outside the signature folder (a picture to copy in).</summary>
    public string Outside(string fileName, byte[] content)
    {
        string path = System.IO.Path.Combine(Root, "outside", fileName);
        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, content);
        return path;
    }

    public void Dispose()
    {
        try
        {
            System.IO.Directory.Delete(Root, recursive: true);
        }
        catch (IOException)
        {
        }
    }
}

/// <summary>Small valid pictures, built rather than stored.</summary>
internal static class TestPictures
{
    /// <summary>A valid PNG of the given size (one IHDR, an empty-ish IDAT, IEND) - enough for any header reader.</summary>
    public static byte[] Png(int width, int height, byte marker = 0)
    {
        List<byte> png = new() { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        png.AddRange(Chunk("IHDR", new byte[]
        {
            (byte)(width >> 24), (byte)(width >> 16), (byte)(width >> 8), (byte)width,
            (byte)(height >> 24), (byte)(height >> 16), (byte)(height >> 8), (byte)height,
            8, 6, 0, 0, 0,
        }));
        png.AddRange(Chunk("IDAT", new byte[] { 0x78, 0x9C, 0x62, marker, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01 }));
        png.AddRange(Chunk("IEND", Array.Empty<byte>()));
        return png.ToArray();
    }

    /// <summary>A JPEG header with a baseline frame of the given size.</summary>
    public static byte[] Jpeg(int width, int height)
    {
        return new byte[]
        {
            0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00,
            0xFF, 0xC0, 0x00, 0x11, 0x08, (byte)(height >> 8), (byte)height, (byte)(width >> 8), (byte)width, 0x03,
            0x01, 0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01, 0xFF, 0xD9,
        };
    }

    public static string Hex(byte[] data)
    {
        return string.Concat(data.Select(b => b.ToString("x2", System.Globalization.CultureInfo.InvariantCulture)));
    }

    private static IEnumerable<byte> Chunk(string type, byte[] data)
    {
        List<byte> chunk = new()
        {
            (byte)(data.Length >> 24), (byte)(data.Length >> 16), (byte)(data.Length >> 8), (byte)data.Length,
        };
        chunk.AddRange(Encoding.ASCII.GetBytes(type));
        chunk.AddRange(data);
        chunk.AddRange(new byte[] { 0, 0, 0, 0 }); // CRC - no reader here checks it
        return chunk;
    }
}
