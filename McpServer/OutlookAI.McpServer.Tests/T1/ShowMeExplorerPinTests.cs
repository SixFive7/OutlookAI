using System.Reflection;
using OutlookAI.Core.Com;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// D49 on Office LTSC 2024: a show-me window must never BE the lifetime pin.
/// <para>
/// Measured on the unindexed test guest (16.0.17932, 2026-10-03, D49 probe v3, Explorer windows
/// only): <c>Explorers.Add</c> on the folder a non-displayed Explorer already shows returned THAT
/// Explorer - the same COM object, <c>Explorers.Count</c> unchanged - while an Add on another folder
/// made a new one, and the non-displayed Explorer then kept Outlook running when the new window was
/// closed. The pin is a non-displayed Explorer on the default Inbox, so <c>goto_folder</c> on that
/// Inbox displayed the pin itself, and closing it took Outlook down (<c>LiveDisconnectRecoveryTests</c>,
/// runs 2 to 8). <see cref="ComposeSurface.AddShowMeExplorer"/> is the fix; these fakes reproduce the
/// measured shape and the older one, and nothing here touches an Outlook.
/// </para>
/// </summary>
public sealed class ShowMeExplorerPinTests
{
    [Fact]
    public void OnTheLtsc2024Shape_AShowMeCallForThePinnedFolder_GetsANewExplorer_NeverThePin()
    {
        FakeFolder inbox = new FakeFolder("Inbox");
        FakeFolder top = new FakeFolder("top folder");
        FakeExplorers explorers = new FakeExplorers { DedupeByFolder = true };
        object? pin = ComposeSurface.TryPinProcess(new FakeApplication(explorers), new FakeNamespace(inbox), out string? pinError);
        try
        {
            Assert.Null(pinError);
            Assert.True(ComposeSurface.IsPin(pin), "the fake pin must be registered the way the product registers one");

            object? shown = ComposeSurface.AddShowMeExplorer(
                f => explorers.Add(f, 0), () => explorers.Count, inbox, () => top, out bool rerouted, out string? error);

            Assert.Null(error);
            Assert.NotNull(shown);
            Assert.NotSame(pin, shown);
            Assert.False(ComposeSurface.IsPin(shown));
            Assert.True(rerouted, "the caller must be told to navigate the window back to the folder asked for");
            Assert.Same(top, ((FakeExplorer)shown!).Folder);
            Assert.Equal(2, explorers.Count);
        }
        finally
        {
            ComposeSurface.ForgetPin(pin);
        }
    }

    [Fact]
    public void Control_WithoutTheGuard_TheLtsc2024ShapeHandsBackThePin()
    {
        // What the show-me path did before the fix: a plain Add on the pinned folder. This is the
        // measured defect, reproduced, so the test above is known to be testing something.
        FakeFolder inbox = new FakeFolder("Inbox");
        FakeExplorers explorers = new FakeExplorers { DedupeByFolder = true };
        object? pin = ComposeSurface.TryPinProcess(new FakeApplication(explorers), new FakeNamespace(inbox), out _);
        try
        {
            object plain = explorers.Add(inbox, 0);
            Assert.Same(pin, plain);
            Assert.True(ComposeSurface.IsPin(plain));
            Assert.Equal(1, explorers.Count);
        }
        finally
        {
            ComposeSurface.ForgetPin(pin);
        }
    }

    [Fact]
    public void OnTheOlderShape_AddMakesANewExplorer_SoNothingIsRerouted()
    {
        FakeFolder inbox = new FakeFolder("Inbox");
        FakeExplorers explorers = new FakeExplorers { DedupeByFolder = false };
        object? pin = ComposeSurface.TryPinProcess(new FakeApplication(explorers), new FakeNamespace(inbox), out _);
        int alternateCalls = 0;
        try
        {
            object? shown = ComposeSurface.AddShowMeExplorer(
                f => explorers.Add(f, 0),
                () => explorers.Count,
                inbox,
                () => { alternateCalls++; return new FakeFolder("top folder"); },
                out bool rerouted,
                out string? error);

            Assert.Null(error);
            Assert.False(rerouted);
            Assert.NotSame(pin, shown);
            Assert.Same(inbox, ((FakeExplorer)shown!).Folder);
            Assert.Equal(0, alternateCalls);
        }
        finally
        {
            ComposeSurface.ForgetPin(pin);
        }
    }

    [Fact]
    public void AnotherFolder_GetsItsOwnExplorer_OnTheLtsc2024ShapeToo()
    {
        FakeFolder inbox = new FakeFolder("Inbox");
        FakeFolder sent = new FakeFolder("Sent Items");
        FakeExplorers explorers = new FakeExplorers { DedupeByFolder = true };
        object? pin = ComposeSurface.TryPinProcess(new FakeApplication(explorers), new FakeNamespace(inbox), out _);
        try
        {
            object? shown = ComposeSurface.AddShowMeExplorer(
                f => explorers.Add(f, 0), () => explorers.Count, sent, () => new FakeFolder("top folder"), out bool rerouted, out string? error);

            Assert.Null(error);
            Assert.False(rerouted);
            Assert.Same(sent, ((FakeExplorer)shown!).Folder);
        }
        finally
        {
            ComposeSurface.ForgetPin(pin);
        }
    }

    [Fact]
    public void AnExistingExplorerThePinRegistryDoesNotKnow_IsCaughtByTheCount()
    {
        // Another server session's pin is reached through another apartment's proxy, whose pointer
        // this registry never held - so the count, not the registry, has to catch it.
        FakeFolder inbox = new FakeFolder("Inbox");
        FakeFolder top = new FakeFolder("top folder");
        FakeExplorers explorers = new FakeExplorers { DedupeByFolder = true };
        object unregistered = explorers.Add(inbox, 0);
        Assert.False(ComposeSurface.IsPin(unregistered));

        object? shown = ComposeSurface.AddShowMeExplorer(
            f => explorers.Add(f, 0), () => explorers.Count, inbox, () => top, out bool rerouted, out string? error);

        Assert.Null(error);
        Assert.True(rerouted);
        Assert.NotSame(unregistered, shown);
        Assert.Same(top, ((FakeExplorer)shown!).Folder);
    }

    [Fact]
    public void NoAlternateFolder_IsAnError_AndThePinIsNeverReturned()
    {
        FakeFolder inbox = new FakeFolder("Inbox");
        FakeExplorers explorers = new FakeExplorers { DedupeByFolder = true };
        object? pin = ComposeSurface.TryPinProcess(new FakeApplication(explorers), new FakeNamespace(inbox), out _);
        try
        {
            object? shown = ComposeSurface.AddShowMeExplorer(
                f => explorers.Add(f, 0), () => explorers.Count, inbox, () => null, out bool rerouted, out string? error);

            Assert.Null(shown);
            Assert.False(rerouted);
            Assert.Contains("already existed", error, StringComparison.Ordinal);
        }
        finally
        {
            ComposeSurface.ForgetPin(pin);
        }
    }

    [Fact]
    public void AnAlternateThatAlsoComesBackExisting_IsAnError()
    {
        FakeFolder inbox = new FakeFolder("Inbox");
        FakeExplorers explorers = new FakeExplorers { DedupeByFolder = true };
        object? pin = ComposeSurface.TryPinProcess(new FakeApplication(explorers), new FakeNamespace(inbox), out _);
        try
        {
            object? shown = ComposeSurface.AddShowMeExplorer(
                f => explorers.Add(f, 0), () => explorers.Count, inbox, () => inbox, out bool rerouted, out string? error);

            Assert.Null(shown);
            Assert.False(rerouted);
            Assert.Contains("every folder tried", error, StringComparison.Ordinal);
            Assert.Equal(1, explorers.Count);
        }
        finally
        {
            ComposeSurface.ForgetPin(pin);
        }
    }

    [Fact]
    public void TheShowMePath_AddsItsExplorerOnlyThroughTheGuard_AndNavigatesARerouteBack()
    {
        string source = File.ReadAllText(Path.GetFullPath(Path.Combine(
            TestProjectDir(), "..", "OutlookAI.Core", "Com", "OutlookComSession.cs")));
        int start = source.IndexOf("private object? EnsureVisibleExplorer(", StringComparison.Ordinal);
        Assert.True(start >= 0, "EnsureVisibleExplorer not found - this test has stopped proving anything");
        int end = source.IndexOf("\n        /// <summary>", start, StringComparison.Ordinal);
        string body = source.Substring(start, (end > start ? end : source.Length) - start);

        Assert.Contains("ComposeSurface.AddShowMeExplorer(", body, StringComparison.Ordinal);
        Assert.DoesNotContain(".Add(folderToShow", body, StringComparison.Ordinal);
        Assert.Contains("if (rerouted)", body, StringComparison.Ordinal);
        Assert.Contains(".CurrentFolder = target", body, StringComparison.Ordinal);
    }

    private static string TestProjectDir()
    {
        return typeof(ShowMeExplorerPinTests).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
    }

    // ------------------------------------------------------------------ fakes (public: the product binds them dynamically)

    /// <summary>A folder; compared by reference, as Outlook compares the folder an Explorer shows.</summary>
    public sealed class FakeFolder
    {
        public FakeFolder(string name)
        {
            Name = name;
        }

        public string Name { get; }
    }

    /// <summary>An Explorer and the folder it was added on.</summary>
    public sealed class FakeExplorer
    {
        public FakeExplorer(FakeFolder folder)
        {
            Folder = folder;
        }

        public FakeFolder Folder { get; }
    }

    /// <summary>
    /// <c>Explorers</c>. With <see cref="DedupeByFolder"/> it behaves as Office LTSC 2024 was
    /// measured to: an Add on a folder an existing Explorer shows returns that Explorer.
    /// </summary>
    public sealed class FakeExplorers
    {
        private readonly List<FakeExplorer> _all = new List<FakeExplorer>();

        public bool DedupeByFolder { get; set; }

        public int Count => _all.Count;

        public object Add(object folder, int displayMode)
        {
            FakeFolder f = (FakeFolder)folder;
            if (DedupeByFolder)
            {
                FakeExplorer? existing = _all.FirstOrDefault(e => ReferenceEquals(e.Folder, f));
                if (existing != null)
                {
                    return existing;
                }
            }

            FakeExplorer created = new FakeExplorer(f);
            _all.Add(created);
            return created;
        }
    }

    /// <summary><c>Application</c>, as far as <see cref="ComposeSurface.TryPinProcess"/> reads it.</summary>
    public sealed class FakeApplication
    {
        public FakeApplication(FakeExplorers explorers)
        {
            Explorers = explorers;
        }

        public FakeExplorers Explorers { get; }
    }

    /// <summary><c>NameSpace</c>, as far as <see cref="ComposeSurface.TryPinProcess"/> reads it.</summary>
    public sealed class FakeNamespace
    {
        private readonly FakeFolder _inbox;

        public FakeNamespace(FakeFolder inbox)
        {
            _inbox = inbox;
        }

        public object GetDefaultFolder(int folderType)
        {
            return folderType == SpecialFolders.OlFolderInbox
                ? _inbox
                : throw new InvalidOperationException("the fake has an Inbox only");
        }
    }
}
