using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins Q74 C1 + C3: the short-decoded-id check is two halves - REJECTED on a cached Exchange store
/// (the Phase-1 finding, workstation only, <c>Requires=CachedExchange</c>) and OPENS AS ITSELF on a PST
/// (inferred, asserted on the guests) - and which half applies to which store is decided in one pure
/// place. The live halves themselves need Outlook; what is pinned here is the decision they assert
/// against and the traits that send each half to the machine it can pass on.
/// </summary>
public sealed class ShortDecodedIdExpectationTests
{
    [Theory]
    [InlineData(3, false, DecodedIdStoreFormat.Pst)]
    [InlineData(3, null, DecodedIdStoreFormat.Pst)]
    [InlineData(0, true, DecodedIdStoreFormat.CachedExchange)]
    [InlineData(1, true, DecodedIdStoreFormat.CachedExchange)]
    [InlineData(2, true, DecodedIdStoreFormat.CachedExchange)]
    [InlineData(0, false, DecodedIdStoreFormat.OnlineExchange)]
    [InlineData(null, true, DecodedIdStoreFormat.Unknown)]
    [InlineData(0, null, DecodedIdStoreFormat.Unknown)]
    [InlineData(null, null, DecodedIdStoreFormat.Unknown)]
    public void TheStoreFormat_IsReadFromTheStoreTypeAndTheCachedFlag(int? exchangeStoreType, bool? isCachedExchange, DecodedIdStoreFormat expected)
    {
        Assert.Equal(expected, ShortDecodedIdExpectation.FormatOf(exchangeStoreType, isCachedExchange));
    }

    [Theory]
    [InlineData(DecodedIdStoreFormat.CachedExchange, DecodedIdOutcome.RejectedAsAnInvalidEntryId)]
    [InlineData(DecodedIdStoreFormat.Pst, DecodedIdOutcome.OpensAsTheItemItself)]
    [InlineData(DecodedIdStoreFormat.OnlineExchange, DecodedIdOutcome.NotApplicable)]
    [InlineData(DecodedIdStoreFormat.Unknown, DecodedIdOutcome.NotApplicable)]
    public void EachFormat_HasExactlyOneExpectedOutcome(DecodedIdStoreFormat format, DecodedIdOutcome expected)
    {
        Assert.Equal(expected, ShortDecodedIdExpectation.ExpectedOutcome(format));
    }

    [Fact]
    public void TheRejectionIsTheInvalidEntryIdHResult()
    {
        // The Phase-1 finding as recorded in v3.MD 0.8 and asserted since: 0x80040107.
        Assert.Equal("80040107", ShortDecodedIdExpectation.InvalidEntryIdHResult);
        Assert.Equal(3, ShortDecodedIdExpectation.NotExchange);
    }

    [Fact]
    public void TheCachedExchangeHalf_RunsOnlyWhereThereIsCachedExchange()
    {
        // Control: before Q74 the one test carried SearchIndex and MultipleStores only, so the guest
        // filter selected it and it failed on every guest by design.
        List<string> requires = Requires(nameof(LiveDecodeVerifyTests.ShortDecodedId_IsRejectedByGetItemFromID_DiscoveryRecorded));

        Assert.Contains(LiveRunFilters.CachedExchange, requires);
        Assert.Contains(LiveRunFilters.CachedExchange, LiveRunFilters.ExchangeOnlyCapabilities);
    }

    [Fact]
    public void ThePstHalf_Exists_AndTheGuestFilterSelectsIt()
    {
        // Control: before Q74 there was no PST half at all, so no guest ever exercised the decoded id.
        List<string> requires = Requires(nameof(LiveDecodeVerifyTests.ShortDecodedId_OpensAsTheItemItself_OnAPstStore));

        Assert.Contains(LiveRunFilters.SearchIndex, requires);
        Assert.DoesNotContain(requires, LiveRunFilters.ExchangeOnlyCapabilities.Contains);
    }

    private static List<string> Requires(string method)
    {
        MethodInfo test = typeof(LiveDecodeVerifyTests).GetMethod(method)
            ?? throw new InvalidOperationException(method + " is gone - this pin proves nothing.");
        return test.GetCustomAttributesData()
            .Where(a => a.AttributeType == typeof(TraitAttribute) && a.ConstructorArguments.Count == 2)
            .Where(a => string.Equals(a.ConstructorArguments[0].Value as string, "Requires", StringComparison.Ordinal))
            .Select(a => (string)a.ConstructorArguments[1].Value!)
            .ToList();
    }
}
