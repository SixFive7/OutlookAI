using System;
using OutlookAI.Core.Diagnostics;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 unit tier (v3.MD section 0.6): pure logic, no index, no Outlook, CI-safe.
/// </summary>
public sealed class PingTests
{
    [Fact]
    public void Echo_WrapsMessageInEnvelope()
    {
        Assert.Equal("OutlookAI.Core echo: hello", Ping.Echo("hello"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("  leading and trailing spaces  ")]
    [InlineData("unicode: ✉ € é")]
    [InlineData("{\"looks\":\"like json\"}")]
    public void Echo_PreservesMessageVerbatim(string message)
    {
        string result = Ping.Echo(message);

        Assert.StartsWith(Ping.EchoPrefix, result, StringComparison.Ordinal);
        Assert.Equal(message, result.Substring(Ping.EchoPrefix.Length));
    }

    [Fact]
    public void Echo_NullMessage_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => Ping.Echo(null!));
    }

    [Fact]
    public void TargetFramework_ReportsNet10ForThisTestHost()
    {
        // This test project targets net10.0-windows, so it must load Core's net10 build.
        // (The net48 build is compile-gated by CI; it is consumed at runtime only in v3.1.)
        Assert.Contains(".NETCoreApp,Version=v10.0", Ping.TargetFramework, StringComparison.Ordinal);
    }
}
