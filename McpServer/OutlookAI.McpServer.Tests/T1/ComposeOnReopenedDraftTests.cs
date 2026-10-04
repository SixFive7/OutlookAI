using System.Reflection;
using System.Text.RegularExpressions;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// No compose inspector is ever opened on an object that <c>Items.Add</c>, <c>Reply</c>, <c>ReplyAll</c>
/// or <c>Forward</c> returned: the draft creators save that object, re-open it by EntryID, and compose on
/// the re-opened item the way update_draft revises a draft.
/// <para>
/// Why (2026-10-04, the test guests' first crash dumps - Docs/live-tier-on-the-vm.md section 4.6): closing
/// a compose inspector of the creator's own object crashed OUTLOOK.EXE - an access violation at
/// <c>OLMAPI32.DLL+0x2E411</c> on Outlook's main thread, inside the client's late-bound
/// <c>Inspector.Close</c> - about one new draft in 40 in a tight loop, and in two of fourteen
/// compose-class suite runs, whatever else was varied. Composing on the re-opened item never crashed: none
/// in 1,600 compositions, none in 400 update_draft revisions of one draft.
/// </para>
/// <para>
/// Read out of the sources, because the failure needs a real Outlook. Nothing here touches one.
/// </para>
/// </summary>
public sealed class ComposeOnReopenedDraftTests
{
    private const string Session = "OutlookAI.Core/Com/OutlookComSession.cs";

    /// <summary>
    /// Anything that opens, uses or closes a compose inspector or its editor - as member accesses and calls,
    /// so a name such as BodyPlacedViaWordEditor, or the "NoWordEditor" token, is not one.
    /// </summary>
    private static readonly string[] InspectorWork =
    {
        ".GetInspector", ".WordEditor", ".Close(", ".Activate(", "PromoteForWordEditor(", "ComposeDraft(", "CloseHiddenInspector(",
    };

    [Theory]
    [InlineData("TryCreateNewDraft")]
    [InlineData("TryCreateDerivedDraft")]
    public void TheDraftCreators_ComposeOnTheReopenedItem_NeverThroughAnInspectorOfTheirOwn(string creator)
    {
        string body = CodeOnly(MemberBody(creator));

        Assert.Equal(1, Count(body, "ComposeReopened(ref mail, "));
        foreach (string work in InspectorWork)
        {
            Assert.False(body.Contains(work, StringComparison.Ordinal), creator + " does inspector work itself (" + work + ") - on the object its creator returned, which is what crashed Outlook");
        }
    }

    [Fact]
    public void ComposeReopened_SavesReopensAndReleasesTheCreatorsObject_BeforeAnyInspector()
    {
        string body = CodeOnly(MemberBody("ComposeReopened"));

        int save = body.IndexOf("created.Save();", StringComparison.Ordinal);
        int reopen = body.IndexOf(".GetItemFromID(entryId", StringComparison.Ordinal);
        int release = body.IndexOf("Release(mail);", StringComparison.Ordinal);
        int replace = body.IndexOf("mail = reopened;", StringComparison.Ordinal);
        int revise = body.IndexOf("ReviseHeldDocument(", StringComparison.Ordinal);
        Assert.True(save >= 0, "ComposeReopened must save the creator's object as it stands");
        Assert.True(save < reopen && reopen < release && release < replace && replace < revise,
            "ComposeReopened must save, re-open by EntryID, release the creator's object and replace it - in that order - before the revision");

        // The revision is handed the re-opened item, never the creator's object.
        string firstArgument = body.Substring(revise + "ReviseHeldDocument(".Length).TrimStart().Split(',')[0].Trim();
        Assert.Equal("mail!", firstArgument);

        foreach (string work in InspectorWork)
        {
            Assert.False(body.Contains(work, StringComparison.Ordinal), "ComposeReopened does inspector work itself (" + work + ")");
        }
    }

    [Fact]
    public void InspectorsAreOpened_OnlyOnItemsOpenedByEntryId()
    {
        // Every GetInspector in the session sits in one of two members ...
        Assert.Equal(
            new[] { "ReviseHeldDocument", "TryApplySignatureOverrideToDraft" },
            MembersContaining("GetInspector").OrderBy(m => m, StringComparer.Ordinal).ToArray());

        // ... one opens its own item by EntryID ...
        Assert.Contains(".GetItemFromID(", CodeOnly(MemberBody("TryApplySignatureOverrideToDraft")), StringComparison.Ordinal);

        // ... and the other is handed one only by update_draft, which opens it by EntryID, and by
        // ComposeReopened, which re-opens it.
        Assert.Equal(
            new[] { "ComposeReopened", "TryUpdateDraft" },
            MembersContaining("ReviseHeldDocument(").Where(m => m != "ReviseHeldDocument").OrderBy(m => m, StringComparer.Ordinal).ToArray());
        Assert.Contains(".GetItemFromID(", CodeOnly(MemberBody("TryUpdateDraft")), StringComparison.Ordinal);
    }

    [Fact]
    public void TheInPlaceCompose_IsGone()
    {
        string text = File.ReadAllText(SessionPath());
        Assert.DoesNotContain("ComposeDraft(", text, StringComparison.Ordinal);
        Assert.DoesNotContain("CloseHiddenInspector(", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Control_TheChecksSeeTheShapesThatCrashed()
    {
        // The lines the 2026-10-04 fix took out of the creators - each one inspector work on the object the
        // creator had just returned - so the scans above are known to see them.
        string[] crashed =
        {
            "                        ComposeDraft((object)draft, body, signatureOverride);",
            "                        inspector = draft.GetInspector;",
            "                        ((dynamic)inspector!).Close(0); // olSave",
            "                        CloseHiddenInspector(mail!);",
            "                        document = ((dynamic)inspector!).WordEditor;",
        };
        foreach (string line in crashed)
        {
            Assert.Contains(InspectorWork, work => line.Contains(work, StringComparison.Ordinal));
        }

        // ... and does not see inspector work in a name or a token that merely mentions the editor.
        foreach (string line in new[]
        {
            "        private (bool SignatureInjected, bool BodyPlacedViaWordEditor) ComposeReopened(",
            "                        composeSurfaceError: wordPlaced ? null : overrideError ?? \"NoWordEditor\");",
        })
        {
            Assert.DoesNotContain(InspectorWork, work => line.Contains(work, StringComparison.Ordinal));
        }

        // And the member reader sees through a tuple return type, which the one-regex readers elsewhere do not.
        Assert.Equal("ComposeReopened", MemberName("        private (bool SignatureInjected, long TextBefore) ComposeReopened("));
        Assert.Equal("ReviseHeldDocument", MemberName("        private static (bool Ok, string? Error) ReviseHeldDocument("));
        Assert.Equal("TryCreateNewDraft", MemberName("        public ComDraftCreateResult? TryCreateNewDraft("));
        Assert.Null(MemberName("        private object? _namespace;"));
        Assert.Null(MemberName("            inspector = draft.GetInspector;"));
    }

    private static readonly Regex Modifiers = new Regex(
        @"^ {8}(?:(?:public|private|internal|protected|static|async|override|virtual|sealed|unsafe|extern|new)\s+)+(?<rest>\S.*)$",
        RegexOptions.CultureInvariant);

    private static readonly Regex NameBeforeParenthesis = new Regex(
        @"^[^(=;]*?\b(?<name>\w+)\s*(?:<[^>]*>)?\s*\(",
        RegexOptions.CultureInvariant);

    /// <summary>The member a line declares, or null: a tuple return type is skipped before the name is read.</summary>
    private static string? MemberName(string line)
    {
        Match declaration = Modifiers.Match(line);
        if (!declaration.Success)
        {
            return null;
        }

        string rest = declaration.Groups["rest"].Value;
        int at = 0;
        if (rest.StartsWith("(", StringComparison.Ordinal))
        {
            int depth = 0;
            for (; at < rest.Length; at++)
            {
                if (rest[at] == '(')
                {
                    depth++;
                }
                else if (rest[at] == ')' && --depth == 0)
                {
                    at++;
                    break;
                }
            }
        }

        Match name = NameBeforeParenthesis.Match(rest.Substring(at));
        return name.Success ? name.Groups["name"].Value : null;
    }

    /// <summary>A member's lines, from its declaration to the next declaration.</summary>
    private static string MemberBody(string member)
    {
        string[] lines = File.ReadAllLines(SessionPath());
        int start = Array.FindIndex(lines, line => MemberName(line) == member);
        Assert.True(start >= 0, member + " was not found in " + Session + " - this test has stopped proving anything");

        int end = start + 1;
        while (end < lines.Length && MemberName(lines[end]) == null)
        {
            end++;
        }

        return string.Join("\n", lines, start, end - start);
    }

    /// <summary>The members whose code holds the needle.</summary>
    private static IEnumerable<string> MembersContaining(string needle)
    {
        string? member = null;
        HashSet<string> found = new HashSet<string>(StringComparer.Ordinal);
        foreach (string line in File.ReadAllLines(SessionPath()))
        {
            member = MemberName(line) ?? member;
            if (member != null && !line.TrimStart().StartsWith("//", StringComparison.Ordinal) && line.Contains(needle, StringComparison.Ordinal))
            {
                found.Add(member);
            }
        }

        return found;
    }

    /// <summary>A member's lines with the whole-line comments taken out, so a sentence about a call is not read as the call.</summary>
    private static string CodeOnly(string body)
    {
        return string.Join("\n", body.Split('\n').Where(line => !line.TrimStart().StartsWith("//", StringComparison.Ordinal)));
    }

    private static int Count(string text, string needle)
    {
        int count = 0;
        for (int at = text.IndexOf(needle, StringComparison.Ordinal); at >= 0; at = text.IndexOf(needle, at + needle.Length, StringComparison.Ordinal))
        {
            count++;
        }

        return count;
    }

    private static string SessionPath()
    {
        string testProjectDir = typeof(ComposeOnReopenedDraftTests).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
        return Path.GetFullPath(Path.Combine(testProjectDir, "..", Session));
    }
}
