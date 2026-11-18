using OutlookAI.Services;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// In-memory <see cref="IPromptRegistry"/>. Case-insensitive on both key and value names,
/// like the registry, and able to fail on demand so the store's failure paths are reachable.
/// The one copy (Q154).
/// </summary>
internal sealed class FakePromptRegistry : IPromptRegistry
{
    private readonly Dictionary<string, Dictionary<string, object>> _keys =
        new(StringComparer.OrdinalIgnoreCase);

    internal bool ThrowOnRead { get; set; }
    internal bool ThrowOnWrite { get; set; }
    internal int StringWrites { get; private set; }
    internal int DwordWrites { get; private set; }
    internal int Deletes { get; private set; }

    internal int TotalWrites => StringWrites + DwordWrites;

    internal void Seed(string subKey, string valueName, object value)
    {
        Key(subKey, create: true)![valueName] = value;
    }

    internal bool Has(string subKey, string valueName)
    {
        Dictionary<string, object>? key = Key(subKey, create: false);
        return key != null && key.ContainsKey(valueName);
    }

    internal string? Raw(string subKey, string valueName)
    {
        Dictionary<string, object>? key = Key(subKey, create: false);
        if (key == null || !key.TryGetValue(valueName, out var value))
        {
            return null;
        }
        return value as string;
    }

    public bool TryReadString(string subKey, string valueName, out string value)
    {
        if (ThrowOnRead)
        {
            throw new InvalidOperationException("fabricated read failure");
        }
        value = string.Empty;
        Dictionary<string, object>? key = Key(subKey, create: false);
        if (key == null || !key.TryGetValue(valueName, out var stored) || stored is not string text)
        {
            return false;
        }
        value = text;
        return true;
    }

    public bool TryReadDword(string subKey, string valueName, out int value)
    {
        if (ThrowOnRead)
        {
            throw new InvalidOperationException("fabricated read failure");
        }
        value = 0;
        Dictionary<string, object>? key = Key(subKey, create: false);
        if (key == null || !key.TryGetValue(valueName, out var stored) || stored is not int number)
        {
            return false;
        }
        value = number;
        return true;
    }

    public void WriteString(string subKey, string valueName, string value)
    {
        if (ThrowOnWrite)
        {
            throw new InvalidOperationException("fabricated write failure");
        }
        StringWrites++;
        Key(subKey, create: true)![valueName] = value;
    }

    public void WriteDword(string subKey, string valueName, int value)
    {
        if (ThrowOnWrite)
        {
            throw new InvalidOperationException("fabricated write failure");
        }
        DwordWrites++;
        Key(subKey, create: true)![valueName] = value;
    }

    public void DeleteValue(string subKey, string valueName)
    {
        if (ThrowOnWrite)
        {
            throw new InvalidOperationException("fabricated write failure");
        }
        Deletes++;
        Key(subKey, create: false)?.Remove(valueName);
    }

    public IList<string> ListValueNames(string subKey)
    {
        if (ThrowOnRead)
        {
            throw new InvalidOperationException("fabricated read failure");
        }
        Dictionary<string, object>? key = Key(subKey, create: false);
        return key == null ? new List<string>() : new List<string>(key.Keys);
    }

    private Dictionary<string, object>? Key(string subKey, bool create)
    {
        string name = subKey ?? string.Empty;
        if (_keys.TryGetValue(name, out var key))
        {
            return key;
        }
        if (!create)
        {
            return null;
        }
        key = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        _keys[name] = key;
        return key;
    }
}
