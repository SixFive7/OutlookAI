using System.Text.RegularExpressions;
using System.Xml.Linq;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The add-in's ribbon callbacks bind when the add-in compiles, not when Office calls them. A ribbon control's
/// <c>Context</c> - the window the ribbon belongs to - is a VARIANT, and the add-in embeds its interop types, so
/// it is typed <c>dynamic</c>: passed on as it is, it makes the whole call bound at run time and its result
/// dynamic too. <c>Ribbon.GetAIAssistantPressed</c> did that, and its <c>pane.Visible</c> then bound against
/// VSTO's own pane class, which is internal: the binder fell back to <c>object</c>, which has no
/// <c>Visible</c>, and the callback threw a RuntimeBinderException on every refresh of the ribbon of a compose
/// window with the pane - 446 and 820 in two ten-minute loops of stream K2's trace of OUTLOOK.EXE, 2026-11-20 -
/// so the AI Assistant toggle never showed the pane as open.
///
/// <para>
/// No test here can load the add-in, a .NET Framework VSTO assembly, so this holds its source to the fix: every
/// use of a callback's <c>control.Context</c> in what <c>OutlookAI.csproj</c> compiles is cast to
/// <c>object</c>, and the pressed callback types its pane. <c>T2/LivePaneTests</c> reads the toggle itself,
/// through the add-in's automation object, which calls the callback as Office does.
/// </para>
/// </summary>
public sealed class RibbonCallbackBindingTests
{
    [Fact]
    public void EveryUseOfARibbonControlsContext_IsCastToObject()
    {
        int uses = 0;
        List<string> uncast = new List<string>();
        foreach (string file in AddInSources())
        {
            string code = WithoutComments(File.ReadAllText(Path.Combine(RepoPaths.RepoRoot(), file)));
            foreach (Match use in Regex.Matches(code, @"\bcontrol\s*\.\s*Context\b"))
            {
                uses++;
                if (!Regex.IsMatch(code.Substring(0, use.Index), @"\(\s*object\s*\)\s*$"))
                {
                    uncast.Add(file + ": ..." + code.Substring(Math.Max(0, use.Index - 60), Math.Min(code.Length - Math.Max(0, use.Index - 60), 90)).Replace('\n', ' '));
                }
            }
        }

        Assert.True(uses >= 2, "the add-in's sources use control.Context " + uses + " time(s); the ribbon's two callbacks alone use it twice, so this no longer reads what it means to.");
        Assert.True(uncast.Count == 0, "control.Context passed on as dynamic, so the call binds at run time:\n" + string.Join("\n", uncast));
    }

    [Fact]
    public void ThePressedCallback_TypesItsPane()
    {
        string ribbon = WithoutComments(File.ReadAllText(Path.Combine(RepoPaths.RepoRoot(), "Ribbon.cs")));
        Match callback = Regex.Match(ribbon, @"public bool GetAIAssistantPressed\(Office\.IRibbonControl control\)\s*\{(?<body>[^}]*)\}");
        Assert.True(callback.Success, "Ribbon.cs has no GetAIAssistantPressed(Office.IRibbonControl control) of the shape this reads.");
        Assert.Contains("CustomTaskPane pane = Globals.ThisAddIn.FindPaneForWindow((object)control.Context);", callback.Groups["body"].Value, StringComparison.Ordinal);
    }

    /// <summary>The C# files <c>OutlookAI.csproj</c> compiles, relative to the repository's root.</summary>
    private static IReadOnlyList<string> AddInSources()
    {
        XDocument project = XDocument.Load(Path.Combine(RepoPaths.RepoRoot(), "OutlookAI.csproj"));
        List<string> files = project.Descendants()
            .Where(e => e.Name.LocalName == "Compile")
            .Select(e => (string?)e.Attribute("Include"))
            .Where(f => f != null && f.EndsWith(".cs", StringComparison.OrdinalIgnoreCase))
            .Select(f => f!.Replace('\\', Path.DirectorySeparatorChar))
            .ToList();
        Assert.Contains("Ribbon.cs", files);
        return files;
    }

    private static string WithoutComments(string code) => Regex.Replace(code, @"/\*.*?\*/|//[^\n]*", " ", RegexOptions.Singleline);
}
