using System.Xml.Linq;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The add-in's reference to OutlookAI.Core (Q154 (b) and Q167 (a), decided by the maintainer
/// 2026-11-17 and 2026-11-18): a ProjectReference to Core's net48 build, which Core's build copies
/// beside the add-in; Core in the solution, so that F5 in Visual Studio builds it; and the two global
/// properties the scripted builds pass that must not reach Core's build.
///
/// <para>
/// It reads the repository's own project files, because the add-in is MSBuild-only and nothing in
/// this suite can build it. What the add-in may USE of Core - the one namespace OutlookAI.Core.AddIn,
/// which stays out of Outlook's COM session, COM threads and the search index (Q164 (a)) - is
/// <c>Tools/Checks/check-pinned-constants.ps1</c>'s check 15.
/// </para>
/// </summary>
public sealed class AddInCoreReferenceTests
{
    private const string CoreProject = @"McpServer\OutlookAI.Core\OutlookAI.Core.csproj";

    [Fact]
    public void TheSolution_HoldsTheAddInAndCore()
    {
        XDocument solution = XDocument.Load(Path.Combine(CachedModeSettingsTests.RepoRoot(), "OutlookAI.slnx"));

        Assert.Equal(
            new[] { "OutlookAI.csproj", "McpServer/OutlookAI.Core/OutlookAI.Core.csproj" },
            solution.Descendants("Project").Select(p => (string?)p.Attribute("Path")));
    }

    [Fact]
    public void TheAddIn_ReferencesCoresNet48Build_AndCopiesIt()
    {
        XElement reference = Assert.Single(AddInProjectReferences());

        Assert.Equal(CoreProject, (string?)reference.Attribute("Include"));
        Assert.Equal("TargetFramework=net48", Metadata(reference, "SetTargetFramework"));
        // The add-in compiles against Core and needs OutlookAI.Core.dll beside it at run time.
        Assert.False(string.Equals(Metadata(reference, "ReferenceOutputAssembly"), "false", StringComparison.OrdinalIgnoreCase), "the add-in must compile against Core");
        Assert.False(string.Equals(Metadata(reference, "Private"), "false", StringComparison.OrdinalIgnoreCase), "Core's assemblies must be copied beside OutlookAI.dll");

        XDocument core = XDocument.Load(Path.Combine(CachedModeSettingsTests.RepoRoot(), CoreProject));
        string frameworks = core.Descendants("TargetFrameworks").Single().Value;
        Assert.Contains("net48", frameworks.Split(';', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries));
    }

    /// <summary>
    /// Testbed/host/Publish-AddInPayload.ps1 and Tools/Switch-AddInBuild.ps1 pass the add-in's settings
    /// as GLOBAL properties, and a global property flows into every referenced project unless the
    /// reference removes it. PrepareForRunDependsOn - guard 1 of Q81 - names
    /// VisualStudioForApplicationsBuild, which Core does not have (MSB4057); DefineConstants
    /// (VSTO40;TRACE) would replace Core's own symbols, NETFRAMEWORK among them.
    /// </summary>
    [Theory]
    [InlineData("PrepareForRunDependsOn")]
    [InlineData("DefineConstants")]
    public void AGlobalPropertyOfTheScriptedBuilds_DoesNotReachCore(string property)
    {
        XElement reference = Assert.Single(AddInProjectReferences());
        string[] removed = (Metadata(reference, "GlobalPropertiesToRemove") ?? "")
            .Split(';', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

        Assert.Contains(property, removed);
    }

    private static IEnumerable<XElement> AddInProjectReferences()
    {
        XDocument addIn = XDocument.Load(Path.Combine(CachedModeSettingsTests.RepoRoot(), "OutlookAI.csproj"));
        return addIn.Descendants(addIn.Root!.Name.Namespace + "ProjectReference");
    }

    private static string? Metadata(XElement item, string name) =>
        (string?)item.Element(item.Name.Namespace + name) ?? (string?)item.Attribute(name);
}
