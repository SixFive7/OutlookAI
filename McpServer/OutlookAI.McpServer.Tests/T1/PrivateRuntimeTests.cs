using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;

using OutlookAI.Core.AddIn;
using OutlookAI.RuntimeSetup;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The helper setup runs to put the mail server's private .NET runtime in place (Q172): what it
/// installs, what it refuses, and that the runtime folder is the old copy or the whole new one and
/// never half of either. Every package here is a small zip built by the test, in the layout of
/// Microsoft's - the host resolver and the shared framework under the pinned version - with its own
/// length and SHA-512, and its "download" a file: nothing touches the network or the real
/// %LOCALAPPDATA%\OutlookAI. The real package is pinned in <see cref="ThePinnedPackageIsMicrosoftsRuntimeZip"/>
/// and fetched for real only by setup.
/// </summary>
public sealed class PrivateRuntimeTests : IDisposable
{
    private readonly string _scratch = Path.Combine(Path.GetTempPath(), "OutlookAI-PrivateRuntimeTests-" + Guid.NewGuid().ToString("N"));

    public PrivateRuntimeTests() => Directory.CreateDirectory(_scratch);

    public void Dispose() => ScratchFiles.TryDeleteTree(_scratch);

    private string Root => Path.Combine(_scratch, "OutlookAI", ServerDotnetRuntime.FolderName);

    /// <summary>The two files <see cref="ServerDotnetRuntime.IsComplete"/> looks for, and one more, as Microsoft's zip has them.</summary>
    private static readonly string[] RuntimeLayout =
    {
        "host/fxr/" + ServerDotnetRuntime.Version + "/hostfxr.dll",
        "shared/Microsoft.NETCore.App/" + ServerDotnetRuntime.Version + "/System.Private.CoreLib.dll",
        "dotnet.exe",
    };

    /// <summary>A zip of <paramref name="entries"/>, each holding its own name, written under the scratch folder.</summary>
    private string Zip(string name, params string[] entries)
    {
        string path = Path.Combine(_scratch, name);
        using (ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            foreach (string entry in entries)
            {
                using Stream s = zip.CreateEntry(entry).Open();
                byte[] bytes = Encoding.UTF8.GetBytes("content of " + entry);
                s.Write(bytes, 0, bytes.Length);
            }
        }

        return path;
    }

    /// <summary>The package a file is: its own length and SHA-512, at an address no test fetches.</summary>
    private static PrivateRuntime.Package PackageOf(string path)
    {
        byte[] bytes = File.ReadAllBytes(path);
        return new PrivateRuntime.Package(new Uri("https://example.invalid/" + Path.GetFileName(path)), bytes.Length, Convert.ToHexStringLower(SHA512.HashData(bytes)));
    }

    /// <summary>A "download" that opens <paramref name="path"/>, and counts how often it was asked.</summary>
    private sealed class FileDownload
    {
        private readonly string? _path;

        public FileDownload(string? path) => _path = path;

        public int Calls { get; private set; }

        public Stream Open(Uri uri, CancellationToken cancel)
        {
            Calls++;
            if (_path == null)
                throw new HttpRequestException("no network in this test");
            return File.OpenRead(_path);
        }
    }

    private PrivateRuntime.Outcome Ensure(PrivateRuntime.Package package, FileDownload download, string? staged = null)
    {
        using var runtime = new PrivateRuntime(Root, package, download.Open);
        return runtime.Ensure(staged, CancellationToken.None);
    }

    private string[] WorkFolders() =>
        Directory.GetDirectories(Path.GetDirectoryName(Root)!, PrivateRuntime.WorkFolderPrefix(Root) + "*");

    private string Log() => File.ReadAllText(PrivateRuntime.LogPath(Root));

    [Fact]
    public void ThePinnedPackageIsMicrosoftsRuntimeZip()
    {
        Assert.Equal("11.0.0", ServerDotnetRuntime.Version);
        Assert.Equal("dotnet-runtime-11.0.0-win-x64.zip", PrivateRuntime.PackageFileName);
        Assert.Equal("https://builds.dotnet.microsoft.com/dotnet/Runtime/11.0.0/dotnet-runtime-11.0.0-win-x64.zip", PrivateRuntime.PinnedUrl);
        Assert.Equal(39856058, PrivateRuntime.PinnedLength);
        Assert.Matches("^[0-9a-f]{128}$", PrivateRuntime.PinnedSha512);
        Assert.Equal(PrivateRuntime.PinnedUrl, PrivateRuntime.Package.Pinned.Url.AbsoluteUri);
        Assert.Equal(PrivateRuntime.PinnedLength, PrivateRuntime.Package.Pinned.Length);
        Assert.Equal(PrivateRuntime.PinnedSha512, PrivateRuntime.Package.Pinned.Sha512);
    }

    [Fact]
    public void AnEmptyFolderGetsTheRuntime_AndNoWorkFolderStays()
    {
        string zip = Zip("good.zip", RuntimeLayout);
        var download = new FileDownload(zip);

        Assert.Equal(PrivateRuntime.Outcome.Installed, Ensure(PackageOf(zip), download));

        Assert.True(ServerDotnetRuntime.IsComplete(Root));
        Assert.Equal(1, download.Calls);
        Assert.Empty(WorkFolders());
        Assert.Contains("SHA-512 as pinned", Log(), StringComparison.Ordinal);
        Assert.Contains("installed", Log(), StringComparison.Ordinal);
    }

    [Fact]
    public void ARuntimeAlreadyThereIsLeftAlone_AndNothingIsFetched()
    {
        string zip = Zip("good.zip", RuntimeLayout);
        Assert.Equal(PrivateRuntime.Outcome.Installed, Ensure(PackageOf(zip), new FileDownload(zip)));
        var second = new FileDownload(null);

        Assert.Equal(PrivateRuntime.Outcome.AlreadyInstalled, Ensure(PackageOf(zip), second));

        Assert.Equal(0, second.Calls);
        Assert.Contains("already installed", Log(), StringComparison.Ordinal);
    }

    [Fact]
    public void APackageBesideTheInstallerIsUsedInsteadOfTheNetwork()
    {
        string zip = Zip("staged.zip", RuntimeLayout);
        var download = new FileDownload(null);

        Assert.Equal(PrivateRuntime.Outcome.Installed, Ensure(PackageOf(zip), download, staged: zip));

        Assert.Equal(0, download.Calls);
        Assert.True(ServerDotnetRuntime.IsComplete(Root));
        Assert.Contains("package from " + zip, Log(), StringComparison.Ordinal);
    }

    [Fact]
    public void AStagedPathWithNoFileThereFallsBackToTheDownload()
    {
        string zip = Zip("good.zip", RuntimeLayout);
        var download = new FileDownload(zip);

        Assert.Equal(PrivateRuntime.Outcome.Installed, Ensure(PackageOf(zip), download, staged: Path.Combine(_scratch, "missing.zip")));

        Assert.Equal(1, download.Calls);
    }

    [Fact]
    public void APackageOfAnotherLengthIsRefused_AndNothingIsUnpacked()
    {
        string zip = Zip("good.zip", RuntimeLayout);
        PrivateRuntime.Package pinned = PackageOf(zip);
        var longer = new PrivateRuntime.Package(pinned.Url, pinned.Length + 1, pinned.Sha512);

        Assert.Equal(PrivateRuntime.Outcome.PackageRejected, Ensure(longer, new FileDownload(zip)));

        Assert.False(Directory.Exists(Root));
        Assert.Empty(WorkFolders());
        Assert.Contains("rejected", Log(), StringComparison.Ordinal);
    }

    [Fact]
    public void APackageOfTheRightLengthAndAnotherHashIsRefused()
    {
        string zip = Zip("good.zip", RuntimeLayout);
        PrivateRuntime.Package pinned = PackageOf(zip);
        var otherHash = new PrivateRuntime.Package(pinned.Url, pinned.Length, new string('0', 128));

        Assert.Equal(PrivateRuntime.Outcome.PackageRejected, Ensure(otherHash, new FileDownload(zip)));

        Assert.False(Directory.Exists(Root));
        Assert.Contains("rejected: SHA-512", Log(), StringComparison.Ordinal);
    }

    [Fact]
    public void ADownloadLongerThanPinnedIsCutAtOneByteMore_AndRefused()
    {
        string zip = Zip("good.zip", RuntimeLayout);
        PrivateRuntime.Package pinned = PackageOf(zip);
        var shorter = new PrivateRuntime.Package(pinned.Url, 10, pinned.Sha512);

        Assert.Equal(PrivateRuntime.Outcome.PackageRejected, Ensure(shorter, new FileDownload(zip)));

        Assert.Contains("rejected: 11 bytes, where 10 are pinned", Log(), StringComparison.Ordinal);
    }

    [Fact]
    public void AnEntryThatWouldLandOutsideTheFolderRefusesThePackage_AndWritesNothing()
    {
        // Hashed as it is, so it passes the hash check: the zip-slip check is the one that must hold.
        string zip = Zip("slip.zip", RuntimeLayout.Concat(new[] { "../../escaped.txt" }).ToArray());

        Assert.Equal(PrivateRuntime.Outcome.PackageRejected, Ensure(PackageOf(zip), new FileDownload(zip)));

        Assert.False(Directory.Exists(Root));
        Assert.Empty(Directory.GetFiles(_scratch, "escaped.txt", SearchOption.AllDirectories));
        Assert.Empty(WorkFolders());
        Assert.Contains("would land outside", Log(), StringComparison.Ordinal);
    }

    [Fact]
    public void APackageWithoutTheRuntimeIsNotPutInPlace()
    {
        string zip = Zip("hollow.zip", "dotnet.exe", "LICENSE.txt");

        Assert.Equal(PrivateRuntime.Outcome.UnpackFailed, Ensure(PackageOf(zip), new FileDownload(zip)));

        Assert.False(Directory.Exists(Root));
        Assert.Contains("holds no runtime " + ServerDotnetRuntime.Version, Log(), StringComparison.Ordinal);
    }

    [Fact]
    public void AnEarlierRuntimeIsReplacedWhole_AndRemoved()
    {
        Directory.CreateDirectory(Path.Combine(Root, "shared", "Microsoft.NETCore.App", "10.9.9"));
        File.WriteAllText(Path.Combine(Root, "shared", "Microsoft.NETCore.App", "10.9.9", "old.dll"), "old");
        string zip = Zip("good.zip", RuntimeLayout);

        Assert.Equal(PrivateRuntime.Outcome.Installed, Ensure(PackageOf(zip), new FileDownload(zip)));

        Assert.True(ServerDotnetRuntime.IsComplete(Root));
        Assert.False(Directory.Exists(Path.Combine(Root, "shared", "Microsoft.NETCore.App", "10.9.9")));
        Assert.Empty(WorkFolders());
    }

    [Fact]
    public void AFailedDownloadLeavesTheEarlierFolderAsItWas()
    {
        string old = Path.Combine(Root, "old.txt");
        Directory.CreateDirectory(Root);
        File.WriteAllText(old, "old");

        Assert.Equal(PrivateRuntime.Outcome.DownloadFailed, Ensure(PackageOf(Zip("good.zip", RuntimeLayout)), new FileDownload(null)));

        Assert.Equal("old", File.ReadAllText(old));
        Assert.Empty(WorkFolders());
        Assert.Contains("download failed: HttpRequestException", Log(), StringComparison.Ordinal);
    }

    [Fact]
    public void AWorkFolderAnEarlierRunLeftIsRemoved()
    {
        string left = Root + ".setup-20260101000000000";
        Directory.CreateDirectory(Path.Combine(left, "runtime"));
        string zip = Zip("good.zip", RuntimeLayout);

        Assert.Equal(PrivateRuntime.Outcome.Installed, Ensure(PackageOf(zip), new FileDownload(zip)));

        Assert.False(Directory.Exists(left));
        Assert.Contains("which an earlier run left behind", Log(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("host/fxr/11.0.0/hostfxr.dll", true)]
    [InlineData("dotnet.exe", true)]
    [InlineData("a/../b.txt", true)]
    [InlineData("../escaped.txt", false)]
    [InlineData("a/../../escaped.txt", false)]
    [InlineData("..\\escaped.txt", false)]
    [InlineData("C:\\escaped.txt", false)]
    [InlineData("\\escaped.txt", false)]
    [InlineData("/escaped.txt", false)]
    [InlineData("file.txt:stream", false)]
    [InlineData("", false)]
    [InlineData(".", false)]
    public void AnEntryLandsInsideTheFolderOrNowhere(string entryName, bool inside)
    {
        string target = Path.Combine(_scratch, "target");
        string? place = PrivateRuntime.PlaceOf(target, entryName);

        Assert.Equal(inside, place != null);
        if (place != null)
            Assert.StartsWith(Path.GetFullPath(target) + Path.DirectorySeparatorChar, place, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetupReadsReadyForBothSuccessesAndACodeForEachFailure()
    {
        Assert.Equal(PrivateRuntime.ExitCode.Ready, PrivateRuntime.ToExitCode(PrivateRuntime.Outcome.Installed));
        Assert.Equal(PrivateRuntime.ExitCode.Ready, PrivateRuntime.ToExitCode(PrivateRuntime.Outcome.AlreadyInstalled));
        Assert.Equal(0, (int)PrivateRuntime.ExitCode.Ready);
        Assert.Equal(10, (int)PrivateRuntime.ToExitCode(PrivateRuntime.Outcome.DownloadFailed));
        Assert.Equal(11, (int)PrivateRuntime.ToExitCode(PrivateRuntime.Outcome.PackageRejected));
        Assert.Equal(12, (int)PrivateRuntime.ToExitCode(PrivateRuntime.Outcome.UnpackFailed));
        Assert.Equal(13, (int)PrivateRuntime.ToExitCode(PrivateRuntime.Outcome.SwapFailed));
    }
}
