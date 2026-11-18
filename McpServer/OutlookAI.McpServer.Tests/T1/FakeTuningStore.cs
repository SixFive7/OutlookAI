using System.Security;

using OutlookAI.Core.AddIn;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// A registry in a dictionary, for the add-in's tuning logic (<c>TuningReconciler</c>,
/// <c>CachedModeSettings</c>, <c>EarlierPolicyRemoval</c>, all in Core's <c>AddIn\</c>). A write
/// or a delete under a denied prefix throws what .NET throws for a key the token may not write;
/// reads never throw, like the add-in's own store.
/// </summary>
internal sealed class FakeTuningStore : ITuningStore
{
    private readonly Dictionary<string, object> _values = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Key prefixes a write or delete under throws <see cref="UnauthorizedAccessException"/> - a filtered token under HKCU\Software\Policies.</summary>
    public List<string> DeniedPrefixes { get; } = new();

    /// <summary>Key prefixes a write or delete under throws <see cref="SecurityException"/> - the other face of access denied.</summary>
    public List<string> SecurityDeniedPrefixes { get; } = new();

    /// <summary>Key prefixes a write or delete under throws <see cref="IOException"/> - any other failure.</summary>
    public List<string> BrokenPrefixes { get; } = new();

    /// <summary>Value names a write or delete of throws <see cref="IOException"/>.</summary>
    public HashSet<string> BrokenValueNames { get; } = new(StringComparer.Ordinal);

    /// <summary>Every write attempted, "key\value", in order - refused ones included.</summary>
    public List<string> WriteAttempts { get; } = new();

    /// <summary>Every delete attempted, "key\value", in order - refused ones included.</summary>
    public List<string> DeleteAttempts { get; } = new();

    private static string Address(string keyPath, string valueName) => keyPath + "\\" + valueName;

    /// <summary>Sets a value as something outside the code under test would: no checks, no record.</summary>
    public void Set(string keyPath, string valueName, object value) => _values[Address(keyPath, valueName)] = value;

    /// <summary>Removes a value as something outside the code under test would: no checks, no record.</summary>
    public void Unset(string keyPath, string valueName) => _values.Remove(Address(keyPath, valueName));

    public object? Get(string keyPath, string valueName) => _values.TryGetValue(Address(keyPath, valueName), out object? v) ? v : null;

    public bool Has(string keyPath, string valueName) => _values.ContainsKey(Address(keyPath, valueName));

    /// <summary>Every value under a key prefix, as "key\value".</summary>
    public List<string> AddressesUnder(string keyPrefix) =>
        _values.Keys.Where(a => a.StartsWith(keyPrefix + "\\", StringComparison.OrdinalIgnoreCase)).ToList();

    public int? ReadDword(string keyPath, string valueName) => Get(keyPath, valueName) is int i ? i : null;

    public string ReadString(string keyPath, string valueName) => Get(keyPath, valueName) as string ?? string.Empty;

    public void WriteDword(string keyPath, string valueName, int value) => Write(keyPath, valueName, value);

    public void WriteString(string keyPath, string valueName, string value) => Write(keyPath, valueName, value);

    public void DeleteValue(string keyPath, string valueName)
    {
        DeleteAttempts.Add(Address(keyPath, valueName));
        ThrowIfRefused(keyPath, valueName);
        _values.Remove(Address(keyPath, valueName));
    }

    private void Write(string keyPath, string valueName, object value)
    {
        WriteAttempts.Add(Address(keyPath, valueName));
        ThrowIfRefused(keyPath, valueName);
        _values[Address(keyPath, valueName)] = value;
    }

    private void ThrowIfRefused(string keyPath, string valueName)
    {
        if (DeniedPrefixes.Any(p => keyPath.StartsWith(p, StringComparison.OrdinalIgnoreCase)))
            throw new UnauthorizedAccessException("Access to the registry key 'HKEY_CURRENT_USER\\" + keyPath + "' is denied.");
        if (SecurityDeniedPrefixes.Any(p => keyPath.StartsWith(p, StringComparison.OrdinalIgnoreCase)))
            throw new SecurityException("Requested registry access is not allowed.");
        if (BrokenPrefixes.Any(p => keyPath.StartsWith(p, StringComparison.OrdinalIgnoreCase)) || BrokenValueNames.Contains(valueName))
            throw new IOException("The device is not ready.");
    }
}
