using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;
using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The registry half of Mailbox Safety rule 6 (Q147, research section 5.6): every registry value
/// that decides which signature Outlook uses, captured BEFORE a signature-touching live suite and
/// verified identical AFTER it. <see cref="SignatureDirectorySnapshot"/> proves the files untouched;
/// this proves the settings untouched - a test may change a value only if it puts it back.
/// <para>
/// What it covers, for the detected Office major: the recorded default profile; the folder value
/// (<c>Common\General\Signatures</c>); the deployed preference and both policy copies of
/// <c>MailSettings</c> (<c>DisableSignatures</c>, <c>NewSignature</c>, <c>ReplySignature</c>); the
/// roaming switches under <c>Outlook\Setup</c> and its policy copy; and <c>New Signature</c> /
/// <c>Reply-Forward Signature</c> on every account-manager entry of every profile - accounts, data
/// files and address books alike, because a stray value on a data file is exactly the 2026-09-24
/// defect. Values are compared by type and SHA-256 and reported by PATH only.
/// </para>
/// <para>
/// A snapshot that cannot be taken REFUSES the suite, as the directory snapshot does.
/// </para>
/// </summary>
public sealed class SignatureRegistrySnapshot
{
    private SignatureRegistrySnapshot(IReadOnlyDictionary<string, string> values)
    {
        Values = values;
    }

    /// <summary>"HKCU\path|value name" (OrdinalIgnoreCase) -> "kind:sha256" of every covered value that exists.</summary>
    public IReadOnlyDictionary<string, string> Values { get; }

    /// <summary>The keys and value names covered, outside the profiles (which are enumerated).</summary>
    public static IReadOnlyList<(string Key, string Value)> FixedLocations(string officeVersion)
    {
        List<(string Key, string Value)> locations = new()
        {
            (OutlookProfileRegistry.BuildOutlookRootKeyPath(officeVersion), LiveProfileAccountRegistry.DefaultProfileValueName),
            (@"Software\Microsoft\Office\" + officeVersion + @"\Common\General", SignatureFolder.RegistryValueName),
        };
        foreach (SignaturePolicySource source in new[] { SignaturePolicySource.Preference, SignaturePolicySource.GroupPolicy, SignaturePolicySource.CloudPolicy })
        {
            string key = SignaturePolicy.KeyPath(source, officeVersion);
            locations.Add((key, SignaturePolicy.DisableSignaturesValueName));
            locations.Add((key, SignaturePolicy.NewSignatureValueName));
            locations.Add((key, SignaturePolicy.ReplySignatureValueName));
        }

        foreach (string key in SignatureRoaming.SwitchKeyPaths(officeVersion))
        {
            locations.Add((key, SignatureRoaming.DisableValueName));
            locations.Add((key, SignatureRoaming.DisableTemporaryToggleValueName));
        }

        return locations;
    }

    /// <summary>Captures the live registry. Throws (refusing the suite) when it cannot be read.</summary>
    public static SignatureRegistrySnapshot Capture()
    {
        try
        {
            return Capture(ReadLive());
        }
        catch (Exception ex) when (ex is System.Security.SecurityException or UnauthorizedAccessException or IOException)
        {
            throw new InvalidOperationException(
                "REFUSING to run the signature suite: the signature registry snapshot could not be taken (" + ex.Message
                + "). Without it the user's signature settings cannot be proven untouched.", ex);
        }
    }

    /// <summary>Captures from any source of (key, value name, raw value) - the T1 seam.</summary>
    public static SignatureRegistrySnapshot Capture(IEnumerable<(string Key, string ValueName, object? Value)> values)
    {
        Dictionary<string, string> captured = new(StringComparer.OrdinalIgnoreCase);
        foreach ((string key, string valueName, object? value) in values)
        {
            if (value != null)
            {
                captured["HKCU\\" + key + "|" + valueName] = Fingerprint(value);
            }
        }

        return new SignatureRegistrySnapshot(captured);
    }

    /// <summary>Every difference against a later capture: CHANGED, ADDED or REMOVED, by path. Empty = untouched.</summary>
    public IReadOnlyList<string> Diff(SignatureRegistrySnapshot after)
    {
        List<string> problems = new();
        foreach (KeyValuePair<string, string> entry in Values)
        {
            if (!after.Values.TryGetValue(entry.Key, out string? now))
            {
                problems.Add("REMOVED: " + entry.Key);
            }
            else if (!string.Equals(entry.Value, now, StringComparison.Ordinal))
            {
                problems.Add("CHANGED: " + entry.Key);
            }
        }

        foreach (string key in after.Values.Keys)
        {
            if (!Values.ContainsKey(key))
            {
                problems.Add("ADDED: " + key);
            }
        }

        problems.Sort(StringComparer.OrdinalIgnoreCase);
        return problems;
    }

    /// <summary>Re-captures and throws on any difference. Call AFTER the suite (fixture disposal).</summary>
    public void VerifySettingsUntouched()
    {
        IReadOnlyList<string> diff = Diff(Capture());
        if (diff.Count > 0)
        {
            throw new InvalidOperationException(
                "SIGNATURE REGISTRY GUARD VIOLATION: signature settings changed during the suite and were not put back: "
                + string.Join("; ", diff) + ". Outlook may now insert another signature than the user chose - investigate "
                + "and restore before running again.");
        }
    }

    private static IEnumerable<(string Key, string ValueName, object? Value)> ReadLive()
    {
        string office = OutlookProfileRegistry.OfficeVersion;
        List<(string, string, object?)> values = new();
        foreach ((string key, string valueName) in FixedLocations(office))
        {
            using RegistryKey? opened = Registry.CurrentUser.OpenSubKey(key, writable: false);
            values.Add((key, valueName, opened?.GetValue(valueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames)));
        }

        string profilesPath = OutlookProfileRegistry.OutlookRootKeyPath + "\\" + OutlookProfileRegistry.ProfilesSubKeyName;
        using RegistryKey? profiles = Registry.CurrentUser.OpenSubKey(profilesPath, writable: false);
        foreach (string profile in profiles?.GetSubKeyNames() ?? Array.Empty<string>())
        {
            string accountsPath = profilesPath + "\\" + profile + "\\" + OutlookProfileRegistry.AccountsSubKeyName;
            using RegistryKey? accounts = Registry.CurrentUser.OpenSubKey(accountsPath, writable: false);
            foreach (string entry in accounts?.GetSubKeyNames() ?? Array.Empty<string>())
            {
                using RegistryKey? opened = accounts!.OpenSubKey(entry, writable: false);
                foreach (string valueName in new[] { ProfileAccountEntries.NewSignatureValueName, ProfileAccountEntries.ReplyForwardSignatureValueName })
                {
                    values.Add((accountsPath + "\\" + entry, valueName, opened?.GetValue(valueName)));
                }
            }
        }

        return values;
    }

    private static string Fingerprint(object value)
    {
        byte[] bytes = value switch
        {
            byte[] binary => binary,
            string text => Encoding.UTF8.GetBytes(text),
            string[] multi => Encoding.UTF8.GetBytes(string.Join("\0", multi)),
            _ => Encoding.UTF8.GetBytes(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty),
        };
        return value.GetType().Name + ":" + Convert.ToHexString(SHA256.HashData(bytes));
    }
}
