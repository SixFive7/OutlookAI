using Xunit;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// Source files read as text, for the checks that pin what a file of this repository says: where a
/// live test's source is, its lines, and the body of one member cut out of them. The one copy (Q154).
/// </summary>
internal static class SourceText
{
    /// <summary>
    /// A live test's source as it is on disk now. A bare name is under <c>T2/</c>; a name with a
    /// folder (<c>T3/...</c>) is relative to the test project.
    /// </summary>
    public static string LiveSource(string fileName)
    {
        string dir = RepoPaths.TestProjectDir();
        string path = fileName.Contains('/', StringComparison.Ordinal)
            ? Path.Combine(dir, fileName.Replace('/', Path.DirectorySeparatorChar))
            : Path.Combine(dir, "T2", fileName);
        Assert.True(File.Exists(path), "live test source is missing: " + path);
        return File.ReadAllText(path);
    }

    /// <summary>A source text's lines, without the carriage returns a checkout may carry.</summary>
    public static string[] Lines(string text)
    {
        return text.Split('\n').Select(l => l.TrimEnd('\r')).ToArray();
    }

    /// <summary>
    /// The body of the member whose declaration line starts with <paramref name="declarationStart"/>:
    /// the lines between its own opening and closing brace, which stand alone on their lines at
    /// <paramref name="braceIndent"/> - four spaces for a member of a type in a file-scoped namespace,
    /// eight in a block-scoped one. <paramref name="where"/> names the file in a failure.
    /// </summary>
    public static string MemberBody(string[] lines, string declarationStart, string braceIndent, string where)
    {
        int declaration = Array.FindIndex(lines, l => l.TrimStart().StartsWith(declarationStart, StringComparison.Ordinal));
        Assert.True(declaration >= 0, declarationStart + " was not found in " + where + " - this test has stopped proving anything");
        return string.Join("\n", BodyAfter(lines, declaration, braceIndent, where, declarationStart));
    }

    /// <summary>The lines between the first brace pair at <paramref name="braceIndent"/> after line <paramref name="declaration"/>.</summary>
    public static string[] BodyAfter(string[] lines, int declaration, string braceIndent, string where, string what)
    {
        string openBrace = braceIndent + "{";
        string closeBrace = braceIndent + "}";
        int open = Array.FindIndex(lines, declaration, l => string.Equals(l, openBrace, StringComparison.Ordinal));
        int close = open < 0 ? -1 : Array.FindIndex(lines, open + 1, l => string.Equals(l, closeBrace, StringComparison.Ordinal));
        Assert.True(open > declaration && close > open, "could not find the body of " + what + " in " + where);
        return lines[(open + 1)..close];
    }
}
