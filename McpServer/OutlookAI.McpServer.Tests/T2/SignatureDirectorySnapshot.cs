using System.Security.Cryptography;
using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The D38 signature-suite guard (user-ordered, ABSOLUTE): a full snapshot of the
/// Outlook Signatures directory (recursive file list + SHA-256 hashes) taken BEFORE
/// any signature-touching live suite runs, and verified byte-identical AFTER it -
/// except entries carrying the "OutlookAI-McpTest-" prefix, the only names live tests
/// may create/update/delete. The user's real signatures are untouchable: any
/// non-test-prefixed difference (changed, added or removed) fails the run loudly.
/// A snapshot that cannot be taken REFUSES the suite (the capturing fixture throws,
/// failing every test in its collection).
/// </summary>
public sealed class SignatureDirectorySnapshot
{
    private SignatureDirectorySnapshot(string root, IReadOnlyDictionary<string, string> hashesByRelativePath)
    {
        Root = root;
        HashesByRelativePath = hashesByRelativePath;
    }

    /// <summary>The snapshotted directory.</summary>
    public string Root { get; }

    /// <summary>Relative path (OrdinalIgnoreCase) -> "length:sha256hex" of every file, recursive.</summary>
    public IReadOnlyDictionary<string, string> HashesByRelativePath { get; }

    /// <summary>
    /// Captures the directory state. A missing directory is a VALID (empty) snapshot -
    /// machines without signatures stay testable; any read/hash failure throws
    /// (= the guard refuses to run the suite).
    /// </summary>
    public static SignatureDirectorySnapshot Capture(string? directory = null)
    {
        string root = directory ?? SignatureCatalog.DefaultSignatureDirectory;
        Dictionary<string, string> hashes = new(StringComparer.OrdinalIgnoreCase);
        if (!Directory.Exists(root))
        {
            return new SignatureDirectorySnapshot(root, hashes);
        }

        try
        {
            foreach (string file in Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories))
            {
                string relative = Path.GetRelativePath(root, file);
                byte[] bytes = File.ReadAllBytes(file);
                hashes[relative] = bytes.Length + ":" + Convert.ToHexString(SHA256.HashData(bytes));
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            throw new InvalidOperationException(
                "REFUSING to run the signature suite: the Signatures directory snapshot could not be taken ("
                + ex.Message + "). Without the snapshot the real signatures cannot be proven untouched.", ex);
        }

        return new SignatureDirectorySnapshot(root, hashes);
    }

    /// <summary>
    /// Compares a fresh capture against this snapshot and returns every difference
    /// that does NOT carry the test prefix (empty list = the user's real signatures
    /// are bit-identical). Differences are reported by relative path only.
    /// </summary>
    public IReadOnlyList<string> DiffIgnoringTestEntries(SignatureDirectorySnapshot after)
    {
        List<string> problems = new();

        foreach (KeyValuePair<string, string> entry in HashesByRelativePath)
        {
            if (IsTestEntry(entry.Key))
            {
                continue;
            }

            if (!after.HashesByRelativePath.TryGetValue(entry.Key, out string? afterHash))
            {
                problems.Add("REMOVED: " + entry.Key);
            }
            else if (!string.Equals(entry.Value, afterHash, StringComparison.OrdinalIgnoreCase))
            {
                problems.Add("CHANGED: " + entry.Key);
            }
        }

        foreach (string key in after.HashesByRelativePath.Keys)
        {
            if (!IsTestEntry(key) && !HashesByRelativePath.ContainsKey(key))
            {
                problems.Add("ADDED: " + key);
            }
        }

        problems.Sort(StringComparer.OrdinalIgnoreCase);
        return problems;
    }

    /// <summary>
    /// Verifies the directory is unchanged outside test-prefixed entries; throws with
    /// the full difference list otherwise. Call AFTER the suite (fixture disposal).
    /// </summary>
    public void VerifyRealSignaturesUntouched()
    {
        IReadOnlyList<string> diff = DiffIgnoringTestEntries(Capture(Root));
        if (diff.Count > 0)
        {
            throw new InvalidOperationException(
                "SIGNATURE GUARD VIOLATION: the Signatures directory changed outside the test prefix '"
                + SignatureCatalog.TestSignaturePrefix + "' during the suite: " + string.Join("; ", diff)
                + ". The user's real signatures must be bit-identical - investigate before running again.");
        }
    }

    /// <summary>An entry is test-owned when ANY path segment starts with the test prefix.</summary>
    public static bool IsTestEntry(string relativePath)
    {
        foreach (string segment in relativePath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
        {
            if (segment.StartsWith(SignatureCatalog.TestSignaturePrefix, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }
}
