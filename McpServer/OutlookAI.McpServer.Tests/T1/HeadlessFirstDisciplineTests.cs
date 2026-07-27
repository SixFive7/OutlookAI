using System.Reflection;
using ModelContextProtocol.Server;
using OutlookAI.McpServer.Tools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// D33 headless-first surface pin (soak fix, 2026-07-23): Outlook must stay/end up
/// window-less for every tool except the three show-me tools and the draft tools'
/// opt-out `display` parameter (D4). This test forces every FUTURE tool through a
/// conscious classification - adding a tool without putting it in exactly one of the
/// pinned sets below fails the suite, and putting it in the wrong set is a reviewed
/// decision. The live no-new-window proof is T2 LiveHeadlessGuaranteeTests; this pin
/// keeps the tool SURFACE honest on every CI run.
/// </summary>
public sealed class HeadlessFirstDisciplineTests
{
    /// <summary>The ONLY tools allowed to create/show an Outlook window unconditionally (D33).</summary>
    private static readonly IReadOnlyList<string> ShowMeTools =
        ["open_in_outlook", "show_search_results", "goto_folder"];

    /// <summary>Tools that may open a window ONLY via their `display` parameter (D4: default true).</summary>
    private static readonly IReadOnlyList<string> DraftDisplayTools =
        ["new_draft", "reply_draft", "replyall_draft", "forward_draft", "update_draft"];

    /// <summary>Tools that must NEVER cause an Outlook window, in any argument combination.</summary>
    private static readonly IReadOnlyList<string> HeadlessSafeTools =
        ["search", "thread", "read", "save_attachment", "outlook_health", "list_accounts", "list_folders",
         "list_signatures", "manage_signature", "send", "move_mail", "archive_mail", "discard_draft"];

    private static Dictionary<string, MethodInfo> DiscoverAdvertisedTools()
    {
        Dictionary<string, MethodInfo> tools = new(StringComparer.Ordinal);
        foreach (Type type in typeof(OutlookTools).Assembly.GetTypes())
        {
            if (type.GetCustomAttribute<McpServerToolTypeAttribute>() == null)
            {
                continue;
            }

            foreach (MethodInfo method in type.GetMethods(BindingFlags.Public | BindingFlags.Static))
            {
                McpServerToolAttribute? attribute = method.GetCustomAttribute<McpServerToolAttribute>();
                if (attribute?.Name is string name)
                {
                    tools[name] = method;
                }
            }
        }

        return tools;
    }

    [Fact]
    public void EveryAdvertisedTool_IsClassifiedForD33()
    {
        Dictionary<string, MethodInfo> advertised = DiscoverAdvertisedTools();
        HashSet<string> classified = new(StringComparer.Ordinal);
        foreach (string name in ShowMeTools.Concat(DraftDisplayTools).Concat(HeadlessSafeTools))
        {
            Assert.True(classified.Add(name), $"tool '{name}' is classified twice");
        }

        HashSet<string> unclassified = new(advertised.Keys, StringComparer.Ordinal);
        unclassified.ExceptWith(classified);
        Assert.True(unclassified.Count == 0,
            "New tool(s) not classified for the D33 headless-first guarantee: "
            + string.Join(", ", unclassified)
            + ". Decide whether each may create Outlook windows and add it to the matching pinned set.");

        HashSet<string> missing = new(classified, StringComparer.Ordinal);
        missing.ExceptWith(advertised.Keys);
        Assert.True(missing.Count == 0,
            "Pinned tool(s) no longer advertised: " + string.Join(", ", missing));
    }

    [Fact]
    public void OnlyDraftTools_HaveADisplayParameter_DefaultingTrue()
    {
        Dictionary<string, MethodInfo> advertised = DiscoverAdvertisedTools();
        foreach ((string name, MethodInfo method) in advertised)
        {
            ParameterInfo? display = method.GetParameters()
                .FirstOrDefault(p => string.Equals(p.Name, "display", StringComparison.OrdinalIgnoreCase));

            if (DraftDisplayTools.Contains(name))
            {
                Assert.True(display != null, $"draft tool '{name}' must expose the display parameter (D4)");
                Assert.True(display!.HasDefaultValue && Equals(display.DefaultValue, true),
                    $"draft tool '{name}' display parameter must default to true (D4)");
            }
            else
            {
                Assert.True(display == null,
                    $"tool '{name}' exposes a display parameter but is not a draft tool - "
                    + "window creation outside the show-me set violates D33");
            }
        }
    }

    [Fact]
    public void ShowMeTools_AreExactlyTheDocumentedThree()
    {
        // D33's allow-list is part of the product contract (v3.MD): growing it is a
        // decision-table change, not a code change.
        Assert.Equal(["goto_folder", "open_in_outlook", "show_search_results"],
            ShowMeTools.OrderBy(n => n, StringComparer.Ordinal).ToArray());
    }
}
