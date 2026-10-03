using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;

using OutlookAI.PolicyWriter;
using OutlookAI.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The elevated helper from start to finish - <c>PolicyWriter\PolicyWriterRun.cs</c>, LINKED into
/// this assembly - against a fake Windows, so the order of its checks is pinned without an
/// administrator, a second user or a registry.
///
/// <para>
/// THE CASE IT EXISTS FOR (Q128). A standard user, Alice, approves the UAC prompt with an
/// administrator's credentials, Bob's. The helper then runs AS BOB, and Bob's HKCU is Bob's hive: a
/// write there would change Bob's Outlook and report success. So the request names Alice's SID, the
/// helper checks that Alice's process started it, and every write goes under
/// <c>HKEY_USERS\&lt;Alice's SID&gt;</c>. The fake host below records every write with the SID it was
/// aimed at, so a write anywhere else would show.
/// </para>
/// </summary>
public sealed class PolicyWriterRunTests
{
    private const string Alice = "S-1-5-21-1004336348-1177238915-682003330-1001";
    private const string Bob = "S-1-5-21-1004336348-1177238915-682003330-500";

    private sealed class FakeHost : IPolicyWriterHost
    {
        public string Requester { get; set; } = Alice;

        public string RequesterWhy { get; set; } = string.Empty;

        public HashSet<string> LoadedHives { get; } = new(StringComparer.Ordinal) { Alice, Bob };

        public Exception? WriteFails { get; set; }

        public int? ReadBackOverride { get; set; }

        public int RequesterAsked { get; private set; }

        public List<(string Sid, string KeyPath, string Name, int Value)> Writes { get; } = new();

        public string RequesterSid(out string why)
        {
            RequesterAsked++;
            why = RequesterWhy;
            return Requester;
        }

        public bool IsUserHiveLoaded(string sid) => LoadedHives.Contains(sid);

        public void WriteDword(string sid, string keyPath, string name, int value)
        {
            if (WriteFails != null && Writes.Count >= 1)
                throw WriteFails;
            Writes.Add((sid, keyPath, name, value));
        }

        public int? ReadDword(string sid, string keyPath, string name)
        {
            if (ReadBackOverride.HasValue)
                return ReadBackOverride;
            var hit = Writes.LastOrDefault(w => w.Sid == sid && w.KeyPath == keyPath && w.Name == name);
            return hit.Name == null ? null : hit.Value;
        }
    }

    private static string[] Args(string sid, params string[] values) =>
        new[] { "--sid", sid, "--office", "16.0" }.Concat(values).ToArray();

    private static readonly string[] AllFive =
    {
        "SyncWindowSetting=0", "SyncWindowSettingDays=0", "DownloadSharedFolders=1", "CacheOthersMail=1", "DisableSyncSliderForSharedMailbox=1",
    };

    private static int Run(FakeHost host, string[] args, out string log)
    {
        var writer = new StringWriter();
        int code = PolicyWriterRun.Run(args, host, writer);
        log = writer.ToString();
        return code;
    }

    [Fact]
    public void TheRequestersOwnHive_IsWritten_UnderHkeyUsers_AtTheCachedModePolicyKey_AndReadBack()
    {
        var host = new FakeHost();

        int code = Run(host, Args(Alice, AllFive), out string log);

        Assert.Equal(PolicyWriterExit.Written, code);
        Assert.Equal(5, host.Writes.Count);
        Assert.All(host.Writes, w => Assert.Equal(Alice, w.Sid));
        Assert.All(host.Writes, w => Assert.Equal(@"Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode", w.KeyPath));
        Assert.Equal(CachedModePolicy.ValueNames, host.Writes.Select(w => w.Name));
        Assert.Contains(@"HKEY_USERS\" + Alice + @"\Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode", log);
    }

    [Fact]
    public void OverTheShoulder_TheAdministratorsHiveIsNeverTheTarget()
    {
        // Alice's Outlook started the helper; the helper runs as Bob. The requester the host
        // reports is Alice - the parent process's user, not the helper's own - and so is the target.
        var host = new FakeHost { Requester = Alice };

        int code = Run(host, Args(Alice, "SyncWindowSetting=0"), out _);

        Assert.Equal(PolicyWriterExit.Written, code);
        Assert.DoesNotContain(host.Writes, w => w.Sid == Bob);
    }

    [Fact]
    public void ARequestNamingSomeoneElse_IsRefused_BeforeAnyWrite()
    {
        // Alice's process asks for Bob's hive - or Bob's for Alice's: either way, not the requester's own.
        var host = new FakeHost { Requester = Alice };

        int code = Run(host, Args(Bob, "SyncWindowSetting=0"), out string log);

        Assert.Equal(PolicyWriterExit.RefusedUser, code);
        Assert.Empty(host.Writes);
        Assert.Contains("REFUSED (user)", log);
    }

    [Fact]
    public void AnUnknownRequester_IsRefused_BeforeAnyWrite()
    {
        var host = new FakeHost { Requester = string.Empty, RequesterWhy = "the process that started this helper is no longer running." };

        int code = Run(host, Args(Alice, AllFive), out string log);

        Assert.Equal(PolicyWriterExit.RefusedUser, code);
        Assert.Empty(host.Writes);
        Assert.Contains("no longer running", log);
    }

    [Fact]
    public void AHiveThatIsNotLoaded_IsRefused_NeverLoaded()
    {
        var host = new FakeHost();
        host.LoadedHives.Remove(Alice);

        int code = Run(host, Args(Alice, AllFive), out _);

        Assert.Equal(PolicyWriterExit.RefusedHiveNotLoaded, code);
        Assert.Empty(host.Writes);
    }

    [Fact]
    public void AMalformedRequest_IsRefused_BeforeTheHostIsEvenAsked()
    {
        var host = new FakeHost();

        int code = Run(host, Args(Alice, "SyncWindowSetting=0", @"..\Run\x=1"), out string log);

        Assert.Equal(PolicyWriterExit.RefusedArguments, code);
        Assert.Equal(0, host.RequesterAsked);
        Assert.Empty(host.Writes);
        Assert.Contains("REFUSED (arguments)", log);
    }

    [Fact]
    public void AccessDenied_MeansItWasNotElevated()
    {
        var host = new FakeHost { WriteFails = new UnauthorizedAccessException("denied") };
        Assert.Equal(PolicyWriterExit.AccessDenied, Run(host, Args(Alice, AllFive), out string log));
        Assert.Contains("ACCESS DENIED", log);

        host = new FakeHost { WriteFails = new SecurityException("not allowed") };
        Assert.Equal(PolicyWriterExit.AccessDenied, Run(host, Args(Alice, AllFive), out _));
    }

    [Fact]
    public void AnyOtherWriteFailure_IsFailed_AndSaysWhatWasWrittenBeforeIt()
    {
        var host = new FakeHost { WriteFails = new IOException("disk") };

        int code = Run(host, Args(Alice, AllFive), out string log);

        Assert.Equal(PolicyWriterExit.Failed, code);
        Assert.Contains("Written before it: SyncWindowSetting", log);
    }

    [Fact]
    public void AValueThatReadsBackDifferently_IsReported()
    {
        var host = new FakeHost { ReadBackOverride = 12 };

        Assert.Equal(PolicyWriterExit.ReadBackMismatch, Run(host, Args(Alice, "SyncWindowSetting=0"), out _));
    }
}
