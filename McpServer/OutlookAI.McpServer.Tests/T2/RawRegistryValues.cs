using Microsoft.Win32;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Some HKCU values exactly as they stand - kind and data, or absent - so a live test that changes
/// a signature setting can put back exactly what was there. Writing back through the product would
/// not: it writes REG_SZ where Outlook wrote REG_BINARY, and clears a value Outlook left present but
/// empty, and the signature registry snapshot (<see cref="SignatureRegistrySnapshot"/>) fails the run
/// on either. A key the capture did not find is removed again on restore, once it is empty. The one
/// copy (Q154) of the live tier's HKCU reads and writes: a value read raw, a value set, a capture put back.
/// </summary>
internal sealed class RawRegistryValues
{
    private readonly string _keyPath;
    private readonly bool _keyExisted;
    private readonly List<(string Name, RegistryValueKind Kind, object? Data)> _values;

    private RawRegistryValues(string keyPath, bool keyExisted, List<(string, RegistryValueKind, object?)> values)
    {
        _keyPath = keyPath;
        _keyExisted = keyExisted;
        _values = values;
    }

    /// <summary>One value as it stands - its data, unexpanded - or null when it or its key is absent.</summary>
    public static object? Read(string keyPath, string valueName)
    {
        using RegistryKey? key = Registry.CurrentUser.OpenSubKey(keyPath, writable: false);
        return key?.GetValue(valueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
    }

    /// <summary>Sets one value, making its key when it is not there; the test's capture puts it back.</summary>
    public static void Write(string keyPath, string valueName, object value, RegistryValueKind kind)
    {
        using RegistryKey key = Registry.CurrentUser.CreateSubKey(keyPath, writable: true);
        key.SetValue(valueName, value, kind);
    }

    /// <summary>The values as they stand now; absent ones are remembered as absent.</summary>
    public static RawRegistryValues Capture(string keyPath, params string[] valueNames)
    {
        using RegistryKey? key = Registry.CurrentUser.OpenSubKey(keyPath, writable: false);
        List<(string, RegistryValueKind, object?)> values = new();
        foreach (string name in valueNames)
        {
            object? data = key?.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            values.Add((name, data == null ? RegistryValueKind.Unknown : key!.GetValueKind(name), data));
        }

        return new RawRegistryValues(keyPath, key != null, values);
    }

    /// <summary>Writes every value back exactly - or deletes it when it was absent.</summary>
    public void Restore()
    {
        using (RegistryKey key = Registry.CurrentUser.CreateSubKey(_keyPath, writable: true))
        {
            foreach ((string name, RegistryValueKind kind, object? data) in _values)
            {
                if (data == null)
                {
                    key.DeleteValue(name, throwOnMissingValue: false);
                }
                else
                {
                    key.SetValue(name, data, kind);
                }
            }
        }

        if (!_keyExisted)
        {
            using RegistryKey? left = Registry.CurrentUser.OpenSubKey(_keyPath, writable: false);
            bool empty = left != null && left.ValueCount == 0 && left.SubKeyCount == 0;
            left?.Dispose();
            if (empty)
            {
                Registry.CurrentUser.DeleteSubKey(_keyPath, throwOnMissingSubKey: false);
            }
        }
    }
}
