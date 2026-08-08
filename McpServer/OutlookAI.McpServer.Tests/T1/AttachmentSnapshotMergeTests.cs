using System;
using System.Collections.Generic;
using OutlookAI.Core.Com;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Soak fix 21. The draft tools take TWO attachment snapshots of the same saved draft -
/// one inside the composing COM call, one as a plain follow-up call - because the first
/// one lies: an attachment Outlook materialized during the composition (a signature's
/// inline logo) reports Size = 0 there, and in the HTMLBody-fallback shape does not appear
/// at all. These rules decide which snapshot is reported, and they are deliberately
/// MONOTONE: a re-read may only ever improve a result, never replace known bytes with
/// unknown ones or drop an attachment that was already seen.
/// </summary>
public class AttachmentSnapshotMergeTests
{
    private static IReadOnlyList<ComAttachmentInfo> Snapshot(params (string Name, long? Size)[] items)
    {
        var list = new List<ComAttachmentInfo>();
        for (int i = 0; i < items.Length; i++)
        {
            list.Add(new ComAttachmentInfo(i + 1, items[i].Name, items[i].Size));
        }

        return list;
    }

    [Fact]
    public void HasUnsizedAttachment_IsTrue_ForTheZeroByteShapeThatWasReported()
    {
        Assert.True(AttachmentSnapshotMerge.HasUnsizedAttachment(Snapshot(("image001.png", 0))));
    }

    [Fact]
    public void HasUnsizedAttachment_IsTrue_WhenTheSizeIsUnknown()
    {
        Assert.True(AttachmentSnapshotMerge.HasUnsizedAttachment(Snapshot(("image001.png", null))));
    }

    [Fact]
    public void HasUnsizedAttachment_IsTrue_WhenOnlyOneOfSeveralIsUnsized()
    {
        Assert.True(AttachmentSnapshotMerge.HasUnsizedAttachment(Snapshot(("a.pdf", 4096), ("image001.png", 0))));
    }

    [Fact]
    public void HasUnsizedAttachment_IsFalse_WhenEveryByteCountIsKnown()
    {
        Assert.False(AttachmentSnapshotMerge.HasUnsizedAttachment(Snapshot(("a.pdf", 4096), ("image001.png", 3035))));
    }

    [Fact]
    public void HasUnsizedAttachment_IsFalse_ForAnEmptySnapshot()
    {
        Assert.False(AttachmentSnapshotMerge.HasUnsizedAttachment(Array.Empty<ComAttachmentInfo>()));
    }

    [Fact]
    public void KnownBytes_CountsOnlyPositiveSizes()
    {
        Assert.Equal(3035, AttachmentSnapshotMerge.KnownBytes(Snapshot(("image001.png", 3035), ("x.png", 0), ("y.png", null))));
    }

    [Fact]
    public void IsBetter_TheReportedDefect_ZeroBytesReplacedByRealBytes()
    {
        IReadOnlyList<ComAttachmentInfo> composeCall = Snapshot(("image001.png", 0));
        IReadOnlyList<ComAttachmentInfo> followUpCall = Snapshot(("image001.png", 3035));
        Assert.True(AttachmentSnapshotMerge.IsBetter(followUpCall, composeCall));
    }

    [Fact]
    public void IsBetter_TheFallbackShape_AnUnseenAttachmentAppears()
    {
        Assert.True(AttachmentSnapshotMerge.IsBetter(Snapshot(("image001.png", 3035)), Array.Empty<ComAttachmentInfo>()));
    }

    [Fact]
    public void IsBetter_IsFalse_WhenTheReReadLostAnAttachment()
    {
        IReadOnlyList<ComAttachmentInfo> good = Snapshot(("image001.png", 3035));
        Assert.False(AttachmentSnapshotMerge.IsBetter(Array.Empty<ComAttachmentInfo>(), good));
    }

    [Fact]
    public void IsBetter_IsFalse_WhenTheReReadWouldReplaceKnownBytesWithZero()
    {
        IReadOnlyList<ComAttachmentInfo> good = Snapshot(("image001.png", 3035));
        IReadOnlyList<ComAttachmentInfo> worse = Snapshot(("image001.png", 0));
        Assert.False(AttachmentSnapshotMerge.IsBetter(worse, good));
    }

    [Fact]
    public void IsBetter_IsFalse_ForAnIdenticalSnapshot()
    {
        Assert.False(AttachmentSnapshotMerge.IsBetter(Snapshot(("image001.png", 3035)), Snapshot(("image001.png", 3035))));
    }

    [Fact]
    public void IsBetter_IsFalse_ForTwoEmptySnapshots()
    {
        Assert.False(AttachmentSnapshotMerge.IsBetter(Array.Empty<ComAttachmentInfo>(), Array.Empty<ComAttachmentInfo>()));
    }

    [Fact]
    public void IsBetter_PrefersTheSnapshotThatSeesMoreAttachments_EvenWithFewerKnownBytes()
    {
        IReadOnlyList<ComAttachmentInfo> one = Snapshot(("a.pdf", 999999));
        IReadOnlyList<ComAttachmentInfo> two = Snapshot(("a.pdf", 0), ("image001.png", 0));
        Assert.True(AttachmentSnapshotMerge.IsBetter(two, one));
    }
}
