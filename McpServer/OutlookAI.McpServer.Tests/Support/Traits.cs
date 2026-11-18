using System.Reflection;
using Xunit;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// xunit traits and attribute arguments read from metadata, for the checks that inventory the
/// suite's own tests. Read through <see cref="CustomAttributeData"/> because xunit's
/// <see cref="TraitAttribute"/> exposes no Name or Value: its constructor arguments are the only
/// place the pair exists. The one copy (Q154).
/// </summary>
internal static class Traits
{
    /// <summary>Whether <paramref name="attributes"/> carry the trait <paramref name="name"/>=<paramref name="value"/>.</summary>
    public static bool Has(IEnumerable<CustomAttributeData> attributes, string name, string value)
    {
        return attributes.Any(a => a.AttributeType == typeof(TraitAttribute)
            && a.ConstructorArguments.Count == 2
            && string.Equals(a.ConstructorArguments[0].Value as string, name, StringComparison.Ordinal)
            && string.Equals(a.ConstructorArguments[1].Value as string, value, StringComparison.Ordinal));
    }

    /// <summary>Whether <paramref name="attributes"/> carry <c>Category=Live</c>.</summary>
    public static bool IsLive(IEnumerable<CustomAttributeData> attributes)
    {
        return Has(attributes, "Category", "Live");
    }

    /// <summary>Whether <paramref name="type"/> is a class marked <c>Category=Live</c> itself.</summary>
    public static bool IsLiveClass(Type type)
    {
        return type.IsClass && IsLive(type.GetCustomAttributesData());
    }

    /// <summary>
    /// The string argument of the first attribute named <paramref name="attributeName"/> (a type
    /// name, such as <c>CollectionAttribute</c>) on <paramref name="type"/> that takes exactly one; null
    /// when there is none.
    /// </summary>
    public static string? SingleAttributeArgument(Type type, string attributeName)
    {
        return type.GetCustomAttributesData()
            .Where(a => string.Equals(a.AttributeType.Name, attributeName, StringComparison.Ordinal))
            .Where(a => a.ConstructorArguments.Count == 1)
            .Select(a => a.ConstructorArguments[0].Value as string)
            .FirstOrDefault(value => value != null);
    }
}
