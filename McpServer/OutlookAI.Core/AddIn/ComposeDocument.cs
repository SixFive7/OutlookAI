using System;
using System.IO;
using System.Text;

namespace OutlookAI.Core.AddIn
{
    /// <summary>How a plain-text body is laid out in the compose document.</summary>
    public enum DraftTextLayout
    {
        /// <summary>
        /// Every line break a soft line break (vertical tab) inside one paragraph, and an empty
        /// paragraph between the body and a signature or quote below it: what the MCP server's draft
        /// tools write, matching the &lt;br&gt; of their HTML fallback.
        /// </summary>
        LineBreaks,

        /// <summary>
        /// Every line its own paragraph, and one paragraph mark after the body: what the add-in's task
        /// pane writes.
        /// </summary>
        Paragraphs,
    }

    /// <summary>
    /// The regions of an Outlook compose document and the writes into them, on an already-acquired
    /// Word document: the draft the user or the agent writes, the signature region Outlook marks with
    /// the hidden bookmark <c>_MailAutoSig</c>, and the quoted original it marks with
    /// <c>_MailOriginal</c>. The one copy the add-in's task pane and the MCP server both use (the
    /// code-sharing refactor's stage S4, survey groups A02 to A04); until then the server kept a port
    /// of the pane's technique to late-bound COM, which had since gained fixes the pane lacked.
    /// <para>
    /// Every member takes the document as an object and works late-bound, and nothing here finds a
    /// document, an inspector or Outlook itself: the caller hands over a document it holds, on the
    /// thread that holds it. The HTML write puts a temporary .htm file under the user's
    /// <c>%LOCALAPPDATA%\OutlookAI\tmp</c> for Word's converter and deletes it at once.
    /// </para>
    /// </summary>
    public static class ComposeDocument
    {
        /// <summary>The hidden bookmark over the signature region.</summary>
        public const string SignatureBookmark = "_MailAutoSig";

        /// <summary>The hidden bookmark over the quoted original of a reply or forward.</summary>
        public const string OriginalBookmark = "_MailOriginal";

        // ------------------------------------------------------------------ reading the regions

        /// <summary>
        /// Where the draft region ends: the start of the signature region, else of the quoted original,
        /// else the end of the document.
        /// </summary>
        public static int FindDraftEnd(object document)
        {
            dynamic doc = document;
            object? bookmarks = null;
            try
            {
                bookmarks = doc.Bookmarks;
                dynamic bm = (dynamic)bookmarks!;

                // Both markers are HIDDEN bookmarks - invisible to Exists() unless ShowHidden is on.
                bm.ShowHidden = true;
                foreach (string name in new[] { SignatureBookmark, OriginalBookmark })
                {
                    if ((bool)bm.Exists(name))
                    {
                        return BookmarkSpan(bookmarks!, name).Start;
                    }
                }
            }
            finally
            {
                ComObjects.Release(bookmarks);
            }

            return GetDocumentEnd(document);
        }

        /// <summary>The draft region's text, without the line breaks around it.</summary>
        public static string ReadDraftText(object document)
        {
            dynamic doc = document;
            int draftEnd = FindDraftEnd(document);
            object? range = null;
            try
            {
                range = doc.Range(0, draftEnd);
                string text = (string?)((dynamic)range!).Text ?? string.Empty;
                return text.Trim('\r', '\n');
            }
            finally
            {
                ComObjects.Release(range);
            }
        }

        /// <summary>The signature region's text, or empty when the draft has none.</summary>
        public static string ReadSignatureText(object document) => ReadBookmarkText(document, SignatureBookmark);

        /// <summary>The quoted original's text, or empty when the draft has none.</summary>
        public static string ReadThreadText(object document) => ReadBookmarkText(document, OriginalBookmark);

        private static string ReadBookmarkText(object document, string name)
        {
            dynamic doc = document;
            object? bookmarks = null;
            object? bookmark = null;
            object? range = null;
            try
            {
                bookmarks = doc.Bookmarks;
                dynamic bm = (dynamic)bookmarks!;
                bm.ShowHidden = true;
                if (!(bool)bm.Exists(name))
                {
                    return string.Empty;
                }

                bookmark = bm.Item(name);
                range = ((dynamic)bookmark!).Range;
                string text = (string?)((dynamic)range!).Text ?? string.Empty;
                return text.TrimEnd('\r', '\n');
            }
            finally
            {
                ComObjects.Release(range);
                ComObjects.Release(bookmark);
                ComObjects.Release(bookmarks);
            }
        }

        private static (int Start, int End) BookmarkSpan(object bookmarks, string name)
        {
            object? bookmark = null;
            object? range = null;
            try
            {
                bookmark = ((dynamic)bookmarks).Item(name);
                range = ((dynamic)bookmark!).Range;
                return ((int)((dynamic)range!).Start, (int)((dynamic)range).End);
            }
            finally
            {
                ComObjects.Release(range);
                ComObjects.Release(bookmark);
            }
        }

        // ------------------------------------------------------------------ the signature region

        /// <summary>
        /// The bookmark dance itself: replace the <c>_MailAutoSig</c> region (or insert above
        /// <c>_MailOriginal</c>, or at the end of the document when neither exists) with the
        /// signature file's content, and recreate the <c>_MailAutoSig</c> bookmark over it, so that
        /// Outlook and the draft/signature/thread split keep working on this draft. The draft region and
        /// the quoted thread are never modified. A missing file inserts nothing.
        /// </summary>
        public static (bool Applied, string? Error) ApplySignature(object document, string signatureFilePath)
        {
            if (string.IsNullOrEmpty(signatureFilePath) || !File.Exists(signatureFilePath))
            {
                return (false, "SignatureFileMissing");
            }

            dynamic doc = document;
            object? bookmarks = null;
            try
            {
                bookmarks = doc.Bookmarks;
                dynamic bm = (dynamic)bookmarks!;

                // _MailAutoSig/_MailOriginal are HIDDEN bookmarks - invisible to
                // Exists() unless ShowHidden is on.
                bm.ShowHidden = true;

                int insertAt;
                if ((bool)bm.Exists(SignatureBookmark))
                {
                    // Replace: drop the marker, then the signature content itself.
                    object? sigBookmark = null;
                    object? sigRange = null;
                    try
                    {
                        sigBookmark = bm.Item(SignatureBookmark);
                        sigRange = ((dynamic)sigBookmark!).Range;
                        insertAt = (int)((dynamic)sigRange!).Start;
                        ((dynamic)sigBookmark).Delete();
                        ((dynamic)sigRange).Delete();
                    }
                    finally
                    {
                        ComObjects.Release(sigRange);
                        ComObjects.Release(sigBookmark);
                    }
                }
                else if ((bool)bm.Exists(OriginalBookmark))
                {
                    // No signature yet - a reply or forward made without an inspector never has one: it goes into
                    // the empty paragraph directly above the quoted message's header block, NOT at _MailOriginal's
                    // start, which sits inside that block's bordered first paragraph. There it took the block's top
                    // border and, with the body written above it, put the agent's whole text under the separator line
                    // (Q149, 12 of 12 measured replies and forwards; Q209 (a), decided 2026-11-18).
                    insertAt = PrepareSlotAboveQuote(document, BookmarkSpan(bookmarks!, OriginalBookmark).Start);
                }
                else
                {
                    // Plain new draft: end of document (before the final paragraph mark).
                    insertAt = Math.Max(0, GetDocumentEnd(document) - 1);
                }

                int endBefore = GetDocumentEnd(document);
                object? insertRange = null;
                try
                {
                    insertRange = doc.Range(insertAt, insertAt);
                    ((dynamic)insertRange!).InsertFile(signatureFilePath, Type.Missing, false, false, false);
                }
                finally
                {
                    ComObjects.Release(insertRange);
                }

                int endAfter = GetDocumentEnd(document);
                int newEnd = insertAt + Math.Max(0, endAfter - endBefore);

                // Recreate the marker over the inserted content so Outlook (and the
                // draft/signature/quote split) keep working on this draft.
                object? newRange = null;
                object? newMark = null;
                try
                {
                    newRange = doc.Range(insertAt, newEnd);

                    // Bookmarks.Add hands back the Bookmark it made: released here, while the
                    // document is still open, never by the garbage collector after it closed.
                    newMark = bm.Add(SignatureBookmark, newRange);
                }
                finally
                {
                    ComObjects.Release(newMark);
                    ComObjects.Release(newRange);
                }

                return endAfter > endBefore ? (true, null) : (false, "InsertFileAddedNothing");
            }
            catch (Exception ex) when (ComObjects.IsCallFailure(ex))
            {
                return (false, ComObjects.DescribeFailure(ex));
            }
            finally
            {
                ComObjects.Release(bookmarks);
            }
        }

        /// <summary>
        /// Removes the signature region - the <c>_MailAutoSig</c> bookmark and its content. Ok with
        /// nothing removed when the draft has no such region.
        /// </summary>
        public static (bool Ok, bool Removed, string? Error) RemoveSignature(object document)
        {
            dynamic doc = document;
            object? bookmarks = null;
            object? sigBookmark = null;
            object? sigRange = null;
            try
            {
                bookmarks = doc.Bookmarks;
                dynamic bm = (dynamic)bookmarks!;

                // _MailAutoSig is a HIDDEN bookmark - invisible to Exists() unless ShowHidden is on.
                bm.ShowHidden = true;
                if (!(bool)bm.Exists(SignatureBookmark))
                {
                    return (true, false, null);
                }

                sigBookmark = bm.Item(SignatureBookmark);
                sigRange = ((dynamic)sigBookmark!).Range;
                ((dynamic)sigBookmark).Delete();
                ((dynamic)sigRange!).Delete();
                return (true, true, null);
            }
            catch (Exception ex) when (ComObjects.IsCallFailure(ex))
            {
                return (false, false, ComObjects.DescribeFailure(ex));
            }
            finally
            {
                ComObjects.Release(sigRange);
                ComObjects.Release(sigBookmark);
                ComObjects.Release(bookmarks);
            }
        }

        /// <summary>
        /// Turns every LINKED picture in the compose document into an EMBEDDED one, and
        /// returns how many it converted. Best-effort by design: a failure here must never
        /// fail a composition, so every COM problem is swallowed and simply not counted.
        /// <para>
        /// ⚠ THE MECHANISM THIS EXISTS FOR (live-measured, D47). Word's
        /// <c>Range.InsertFile</c> of a signature .htm does NOT embed the signature's
        /// images: it inserts each one as a LINKED <c>InlineShape</c> pointing at the file
        /// on disk, and Outlook stores that link verbatim, so the saved draft carries
        /// <c>&lt;img src="file:///…\Signatures\&lt;name&gt;_files\logo.png"&gt;</c> and
        /// ZERO attachments. The link only renders because the file happens to sit on this
        /// machine. Re-rendering such a document - which is exactly what
        /// <c>update_draft</c> does when the held inspector re-materializes the saved HTML
        /// - makes Word re-serialize the unresolved linked picture as its VML placeholder
        /// AutoShape (<c>&lt;v:rect … alt="logo"&gt;</c>), and the <c>&lt;img&gt;</c>
        /// disappears. Embedding the picture at composition time removes the link, so
        /// Outlook writes real image bytes as an inline <c>cid:</c> attachment that
        /// survives any number of re-renders - and reaches the recipient, which a
        /// <c>file:///</c> link never could.
        /// </para>
        /// <c>SavePictureWithDocument</c> must be set BEFORE <c>BreakLink</c>: breaking a
        /// link whose picture is not stored with the document leaves nothing behind.
        /// </summary>
        public static int EmbedLinkedPictures(object document)
        {
            dynamic doc = document;
            int embedded = 0;
            object? shapes = null;
            try
            {
                shapes = doc.InlineShapes;
                dynamic collection = (dynamic)shapes!;
                int count = (int)collection.Count;
                for (int i = 1; i <= count; i++)
                {
                    object? shape = null;
                    object? link = null;
                    try
                    {
                        shape = collection[i];

                        // An already-embedded picture has no LinkFormat: late-bound COM
                        // answers either null or a binding failure, and both mean skip.
                        link = ((dynamic)shape!).LinkFormat;
                        if (link == null)
                        {
                            continue;
                        }

                        ((dynamic)link!).SavePictureWithDocument = true;
                        ((dynamic)link!).BreakLink();
                        embedded++;
                    }
                    catch (Exception ex) when (ComObjects.IsCallFailure(ex))
                    {
                    }
                    finally
                    {
                        ComObjects.Release(link);
                        ComObjects.Release(shape);
                    }
                }
            }
            catch (Exception ex) when (ComObjects.IsCallFailure(ex))
            {
            }
            finally
            {
                ComObjects.Release(shapes);
            }

            return embedded;
        }

        // ------------------------------------------------------------------ the draft region

        /// <summary>
        /// Writes a plain-text body into the DRAFT region, laid out as <paramref name="layout"/> says.
        /// <para>
        /// The draft region ends where the signature (<c>_MailAutoSig</c>) or, absent that, the quoted
        /// original (<c>_MailOriginal</c>) begins. The boundary marker is DELETED before the write and
        /// RECREATED over the (shifted, byte-identical) region afterwards. That dance is load-bearing,
        /// not cosmetic: Word absorbs text inserted at a bookmark's start INTO that bookmark, so writing
        /// straight to Range(0,0) while <c>_MailAutoSig</c> starts at 0 made the agent body part of the
        /// signature region (the A1 defect - live-proven on an account without a default signature).
        /// Recreating the marker keeps Outlook's signature handling and the draft/signature/thread split
        /// working on the draft. The signature and quote regions themselves are never modified.
        /// </para>
        /// <para>
        /// ⚠ WITH NO BOUNDARY MARKER the draft region has no upper bound, and callers need OPPOSITE
        /// things (live-caught by the batch-C reply test): on a CREATE the document holds only empty
        /// boilerplate, so writing at 0 is correct and safe; on an UPDATE - and in the task pane, which
        /// rewrites what the user wrote - the document holds the PREVIOUS body, so writing at 0 would
        /// PREPEND. <paramref name="replaceToEndWhenNoBoundary"/> makes the draft region run to the end
        /// of the document then (minus Word's final paragraph mark, which cannot be deleted).
        /// </para>
        /// </summary>
        public static (bool Inserted, string? Error) WriteDraftText(
            object document, string text, DraftTextLayout layout, bool replaceToEndWhenNoBoundary)
        {
            return WriteDraftRegion(document, text ?? string.Empty, null, layout, replaceToEndWhenNoBoundary);
        }

        /// <summary>
        /// Writes an already-normalized HTML body into the DRAFT region: the region is CLEARED and the
        /// markup is inserted with <c>Range.InsertFile</c> from a temporary .htm file - the same verb
        /// <see cref="ApplySignature"/> uses, and the only rich-content mechanism proven in this
        /// codebase. Word's own HTML converter does the rendering, so headings, lists and tables arrive
        /// as real Word structures inside the message's WordSection1, and the boundary marker is
        /// re-anchored by the SAME <c>Content.End</c> delta arithmetic - a collapsed insert range does
        /// not span what was inserted, so the length must be measured, not read off the range.
        /// Otherwise as <see cref="WriteDraftText"/>.
        /// </summary>
        public static (bool Inserted, string? Error) WriteDraftHtml(
            object document, string html, bool replaceToEndWhenNoBoundary)
        {
            return WriteDraftRegion(document, null, html ?? string.Empty, DraftTextLayout.LineBreaks, replaceToEndWhenNoBoundary);
        }

        private static (bool Inserted, string? Error) WriteDraftRegion(
            object document, string? text, string? html, DraftTextLayout layout, bool replaceToEndWhenNoBoundary)
        {
            dynamic doc = document;
            object? bookmarks = null;
            string? boundary = null;
            int boundaryStart = -1;
            int boundaryEnd = -1;
            bool aboveQuote = false;
            try
            {
                bookmarks = doc.Bookmarks;
                dynamic bm = (dynamic)bookmarks!;

                // _MailAutoSig/_MailOriginal are HIDDEN bookmarks - invisible to
                // Exists() unless ShowHidden is on.
                bm.ShowHidden = true;

                if ((bool)bm.Exists(SignatureBookmark))
                {
                    boundary = SignatureBookmark;
                }
                else if ((bool)bm.Exists(OriginalBookmark))
                {
                    // A reply or forward with no signature (Q209 (a)): the text goes into the empty paragraphs above
                    // the quoted message's header block. _MailOriginal itself is never touched: nothing is written at
                    // its start, so Word has nothing to absorb into it.
                    aboveQuote = true;
                }

                if (boundary != null)
                {
                    object? marker = null;
                    object? markerRange = null;
                    try
                    {
                        marker = bm.Item(boundary);
                        markerRange = ((dynamic)marker!).Range;
                        boundaryStart = (int)((dynamic)markerRange!).Start;
                        boundaryEnd = (int)((dynamic)markerRange).End;

                        // Marker only - the signature/quote CONTENT stays untouched.
                        ((dynamic)marker).Delete();
                    }
                    finally
                    {
                        ComObjects.Release(markerRange);
                        ComObjects.Release(marker);
                    }
                }

                // The DRAFT region [0, boundary) is REPLACED, so the body starts at the top instead of
                // below a template's blank paragraphs.
                int insertAt = boundary != null ? boundaryStart : 0;
                if (aboveQuote)
                {
                    // The draft region runs to the empty paragraph kept above the header block, which stays as the
                    // blank line between the text and the quote - so the text ends in one paragraph mark, not two,
                    // and an HTML body needs no blank paragraph of its own after it.
                    insertAt = PrepareSlotAboveQuote(document, BookmarkSpan(bookmarks!, OriginalBookmark).Start);
                }
                else if (boundary == null && replaceToEndWhenNoBoundary)
                {
                    insertAt = Math.Max(0, GetDocumentEnd(document) - 1);
                }

                int writtenEnd;
                if (html != null)
                {
                    writtenEnd = InsertHtmlIntoDraftRegion(document, html, insertAt, boundary != null);
                }
                else
                {
                    string written = layout == DraftTextLayout.Paragraphs
                        ? text + "\r\n"
                        : text!.Replace("\r\n", "\n").Replace('\n', '\v') + (boundary != null ? "\r\r" : "\r");
                    object? range = null;
                    try
                    {
                        range = doc.Range(0, insertAt);
                        ((dynamic)range!).Text = written;
                        writtenEnd = (int)((dynamic)range).End;
                    }
                    finally
                    {
                        ComObjects.Release(range);
                    }
                }

                if (boundary != null)
                {
                    // The region moved by (new draft length - old draft length).
                    int shifted = boundaryEnd + (writtenEnd - insertAt);
                    if (shifted < writtenEnd)
                    {
                        shifted = writtenEnd;
                    }

                    // Word's HTML converter may merge the last inserted paragraph with the
                    // signature's first, leaving the document a character or two shorter
                    // than the arithmetic predicts. An out-of-range Range() would throw and
                    // silently demote the whole composition to the HTML fallback, so clamp.
                    int documentEnd = GetDocumentEnd(document);
                    if (shifted > documentEnd)
                    {
                        shifted = documentEnd;
                    }

                    if (writtenEnd > shifted)
                    {
                        writtenEnd = shifted;
                    }

                    object? restoreRange = null;
                    object? restoredMark = null;
                    try
                    {
                        restoreRange = doc.Range(writtenEnd, shifted);
                        restoredMark = bm.Add(boundary, restoreRange);
                    }
                    finally
                    {
                        ComObjects.Release(restoredMark);
                        ComObjects.Release(restoreRange);
                    }
                }

                return (true, null);
            }
            catch (Exception ex) when (ComObjects.IsCallFailure(ex))
            {
                // The marker was deleted up front; restore it over the ORIGINAL range so
                // a failed write cannot silently strip the signature/quote anchor.
                TryRestoreBookmark(bookmarks, boundary, boundaryStart, boundaryEnd);
                return (false, ComObjects.DescribeFailure(ex));
            }
            finally
            {
                ComObjects.Release(bookmarks);
            }
        }

        /// <summary>
        /// Places an already-normalized HTML fragment into the draft region and returns the document
        /// offset just PAST it (which is where the signature/quote region now begins).
        /// <para>
        /// Mechanism: clear the compose boilerplate <c>[0, draftEnd)</c>, then
        /// <c>Range.InsertFile</c> a temporary .htm file at a COLLAPSED range at 0 - the
        /// same call <see cref="ApplySignature"/> makes for signatures, i.e. the
        /// one HTML-into-Word route this codebase has proven. Word's HTML converter turns
        /// the markup into real Word structures inside the message body, so the result
        /// inherits the message style instead of sitting outside WordSection1 (the batch-A
        /// A1(ii) defect). The inserted LENGTH must be measured as a <c>Content.End</c>
        /// delta: a collapsed insert range does not grow to span what was inserted.
        /// </para>
        /// </summary>
        private static int InsertHtmlIntoDraftRegion(object document, string html, int draftEnd, bool hasBoundary)
        {
            dynamic doc = document;
            string path = WriteTemporaryHtmlFile(html, hasBoundary);
            try
            {
                if (draftEnd > 0)
                {
                    // The draft region's old content - replaced, not appended to,
                    // exactly like the plain-text path's Range(0, draftEnd).Text write.
                    object? clearRange = null;
                    try
                    {
                        clearRange = doc.Range(0, draftEnd);
                        ((dynamic)clearRange!).Delete();
                    }
                    finally
                    {
                        ComObjects.Release(clearRange);
                    }
                }

                int endBefore = GetDocumentEnd(document);
                object? insertRange = null;
                try
                {
                    insertRange = doc.Range(0, 0);
                    ((dynamic)insertRange!).InsertFile(path, Type.Missing, false, false, false);
                }
                finally
                {
                    ComObjects.Release(insertRange);
                }

                int endAfter = GetDocumentEnd(document);
                return Math.Max(0, endAfter - endBefore);
            }
            finally
            {
                try
                {
                    File.Delete(path);
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }
            }
        }

        /// <summary>
        /// Q209 (a), decided 2026-11-18: the start of the empty paragraph directly above a reply's or forward's quoted
        /// header block - where the text and signature go, and the paragraph that then stays as the blank line
        /// between them and the quote. <c>Reply()</c> and <c>Forward()</c> leave two empty paragraphs above the header
        /// block for exactly this (Q149 cell C4, 6 of 6); the header block's first paragraph, where
        /// <c>_MailOriginal</c> starts, carries the separator line as its top border, and text written at its start
        /// took that border (12 of 12 measured replies and forwards).
        /// <para>
        /// When the paragraph above the header block is not empty - a draft someone typed into right above the quote -
        /// an empty one is split off its end, with its formatting; when the header block is the document's first
        /// paragraph, an empty one is put before it and stripped of the formatting it copied from the block
        /// (<see cref="QuoteSlotRules"/>). Either way nothing is written into the bordered paragraph.
        /// </para>
        /// </summary>
        private static int PrepareSlotAboveQuote(object document, int originalStart)
        {
            dynamic doc = document;
            int headerStart = ParagraphStartAt(document, originalStart);
            bool lastIsBlank = false;
            int lastStart = 0;
            if (headerStart > 0)
            {
                object? probe = null;
                object? paragraphs = null;
                object? paragraph = null;
                object? range = null;
                try
                {
                    probe = doc.Range(headerStart - 1, headerStart - 1);
                    paragraphs = ((dynamic)probe!).Paragraphs;
                    paragraph = ((dynamic)paragraphs!).Item(1);
                    range = ((dynamic)paragraph!).Range;
                    lastStart = (int)((dynamic)range!).Start;
                    lastIsBlank = QuoteSlotRules.IsBlankParagraphText(TryGetString(() => (string?)((dynamic)range!).Text));
                }
                finally
                {
                    ComObjects.Release(range);
                    ComObjects.Release(paragraph);
                    ComObjects.Release(paragraphs);
                    ComObjects.Release(probe);
                }
            }

            switch (QuoteSlotRules.Decide(headerStart, lastIsBlank))
            {
                case QuoteSlotAction.UseLastParagraph:
                    return lastStart;

                case QuoteSlotAction.SplitLastParagraph:
                {
                    // A paragraph mark just before the last paragraph's own: its text keeps its paragraph, and an empty
                    // paragraph with the same formatting - outside the header block - is left above the block.
                    object? at = null;
                    try
                    {
                        at = doc.Range(headerStart - 1, headerStart - 1);
                        ((dynamic)at!).InsertBefore("\r");
                    }
                    finally
                    {
                        ComObjects.Release(at);
                    }

                    return headerStart;
                }

                default:
                {
                    // The header block is the first paragraph: an empty paragraph before it copies its formatting, the
                    // top border with it, so that formatting is taken off again.
                    object? at = null;
                    object? paragraphs = null;
                    object? first = null;
                    object? borders = null;
                    try
                    {
                        at = doc.Range(0, 0);
                        ((dynamic)at!).InsertBefore("\r");
                        paragraphs = doc.Paragraphs;
                        first = ((dynamic)paragraphs!).Item(1);
                        ((dynamic)first!).Reset();
                        borders = ((dynamic)first).Borders;
                        ((dynamic)borders!).Enable = 0;
                    }
                    catch (Exception ex) when (ComObjects.IsCallFailure(ex))
                    {
                        // Best-effort: the text still goes above the block, perhaps with its border.
                    }
                    finally
                    {
                        ComObjects.Release(borders);
                        ComObjects.Release(first);
                        ComObjects.Release(paragraphs);
                        ComObjects.Release(at);
                    }

                    return 0;
                }
            }
        }

        /// <summary>The start of the paragraph that holds document position <paramref name="position"/>.</summary>
        private static int ParagraphStartAt(object document, int position)
        {
            dynamic doc = document;
            object? probe = null;
            object? paragraphs = null;
            object? paragraph = null;
            object? range = null;
            try
            {
                probe = doc.Range(position, position);
                paragraphs = ((dynamic)probe!).Paragraphs;
                paragraph = ((dynamic)paragraphs!).Item(1);
                range = ((dynamic)paragraph!).Range;
                return (int)((dynamic)range!).Start;
            }
            finally
            {
                ComObjects.Release(range);
                ComObjects.Release(paragraph);
                ComObjects.Release(paragraphs);
                ComObjects.Release(probe);
            }
        }

        private static string? TryGetString(Func<string?> getter)
        {
            try
            {
                return getter();
            }
            catch (Exception ex) when (ComObjects.IsCallFailure(ex))
            {
                return null;
            }
        }

        /// <summary>
        /// Writes the fragment as a self-contained utf-8 .htm file for Word's converter.
        /// It lives under the shared state root (v3.MD section 0.5.2), never in a mailbox
        /// path, and is deleted immediately after the insert; a trailing empty paragraph
        /// is added when a signature/quote region follows, so Word cannot merge the last
        /// body paragraph into it.
        /// </summary>
        private static string WriteTemporaryHtmlFile(string html, bool hasBoundary)
        {
            string directory = Path.Combine(SharedState.Root, "tmp");
            Directory.CreateDirectory(directory);
            PurgeStaleTemporaryFiles(directory);

            string path = Path.Combine(directory, "draft-body-" + Guid.NewGuid().ToString("N") + ".htm");
            string document = SignatureText.EnsureHtmlDocument(html + (hasBoundary ? "\r\n<p>&nbsp;</p>" : string.Empty));
            File.WriteAllText(path, document, new UTF8Encoding(false));
            return path;
        }

        /// <summary>Best-effort cleanup of temp bodies a crashed run could have left behind.</summary>
        private static void PurgeStaleTemporaryFiles(string directory)
        {
            try
            {
                DateTime cutoff = DateTime.UtcNow.AddHours(-1);
                foreach (string stale in Directory.GetFiles(directory, "draft-body-*.htm"))
                {
                    if (File.GetLastWriteTimeUtc(stale) < cutoff)
                    {
                        File.Delete(stale);
                    }
                }
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        private static void TryRestoreBookmark(object? bookmarksObject, string? name, int start, int end)
        {
            if (bookmarksObject == null || name == null || start < 0 || end < start)
            {
                return;
            }

            object? range = null;
            object? parentDoc = null;
            object? restored = null;
            try
            {
                dynamic bm = (dynamic)bookmarksObject;
                if ((bool)bm.Exists(name))
                {
                    return;
                }

                parentDoc = bm.Parent;
                range = ((dynamic)parentDoc!).Range(start, end);
                restored = bm.Add(name, range);
            }
            catch (Exception ex) when (ComObjects.IsCallFailure(ex))
            {
            }
            finally
            {
                ComObjects.Release(restored);
                ComObjects.Release(range);
                ComObjects.Release(parentDoc);
            }
        }

        /// <summary>The end of the document's content.</summary>
        public static int GetDocumentEnd(object document)
        {
            dynamic doc = document;
            object? content = null;
            try
            {
                content = doc.Content;
                return (int)((dynamic)content!).End;
            }
            finally
            {
                ComObjects.Release(content);
            }
        }
    }
}
