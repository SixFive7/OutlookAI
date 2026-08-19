using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// One name and how many of ITS LOWEST-INDEXED instances this call must delete.
    /// <para>
    /// The count is taken before any addition runs, which is what makes "remove and add
    /// the same name in one call means replace" survive the additions being executed
    /// first: <c>Attachments.Add</c> appends, so the instances that existed before the
    /// additions are exactly the lowest-indexed ones.
    /// </para>
    /// </summary>
    public sealed class DraftAttachmentRemoval
    {
        /// <summary>Creates a removal instruction.</summary>
        public DraftAttachmentRemoval(string fileName, int count)
        {
            FileName = fileName;
            Count = count;
        }

        /// <summary>Attachment file name to match, case-insensitively.</summary>
        public string FileName { get; }

        /// <summary>How many instances of that name to delete, lowest index first.</summary>
        public int Count { get; }
    }

    /// <summary>
    /// The attachment work one <c>update_draft</c> call still has to do, computed from
    /// what the caller asked for and what the draft actually carries right now.
    /// <para>
    /// <b>Why this is a plan and not two loops.</b> <c>update_draft</c> used to remove by
    /// name and then add by path, unconditionally. That is correct exactly once. When the
    /// COM host is killed part-way through the sequence and the caller repeats the same
    /// request, running it again removes what the first attempt added, or adds a second
    /// copy of what the first attempt already attached - the two failures that make a
    /// repeat unsafe, and the reason the retry could not be recommended.
    /// </para>
    /// <para>
    /// <b>The rule.</b> Converge on the END STATE rather than replay the STEPS, because
    /// the killed attempt reported no steps. The end state is stated per file name:
    /// </para>
    /// <list type="bullet">
    /// <item><description>a name the caller asked to REMOVE ends with exactly the
    /// instances the caller also asked to ADD - so every current instance is deleted and
    /// every requested path is attached. Redoing that is harmless even when it already
    /// ran, because an attachment added from a path can be added from the same path
    /// again, and by name alone the old copy and the new copy are indistinguishable. This
    /// is the one branch that deliberately repeats work rather than risk keeping the wrong
    /// copy.</description></item>
    /// <item><description>a name the caller only asked to ADD ends with what the draft
    /// held before plus what was requested, so the shortfall against that total is what
    /// still has to go on, taking the LAST paths of that name because additions run in
    /// request order.</description></item>
    /// <item><description>a name in neither list is not touched.</description></item>
    /// </list>
    /// <para>
    /// <b>A first attempt is the identity case.</b> When nothing has been applied yet,
    /// "before" and "now" are the same list and the plan reduces to remove-every-match
    /// plus add-everything - byte for byte the behaviour this method replaced. So the
    /// re-entrant path is not a second mode with its own semantics; it is the same rule
    /// evaluated against a draft that has moved.
    /// </para>
    /// <para>
    /// Pure, so T1 can pin every state a kill can leave behind without an Outlook.
    /// </para>
    /// </summary>
    public static class DraftAttachmentPlan
    {
        /// <summary>
        /// Builds the plan for one ATTEMPT: the pre-image is the recorded one when this call
        /// is a repeat, and the draft's own current names when it is not.
        /// <para>
        /// It exists so that choice is a line T1 can revert. The call site is inside the COM
        /// sequence, where nothing without a real Outlook can reach it, and getting it
        /// backwards - using the live names on a repeat - is silent: every plan still looks
        /// reasonable, and the attachment the first attempt added is simply attached again.
        /// </para>
        /// </summary>
        public static DraftAttachmentWork BuildForAttempt(
            ComDraftUpdateResume? resume,
            IReadOnlyList<string> namesNow,
            IReadOnlyList<string> addPaths,
            IReadOnlyList<string> removeNames)
        {
            return Build(resume?.AttachmentNamesBefore ?? namesNow, namesNow, addPaths, removeNames);
        }

        /// <summary>
        /// Builds the plan.
        /// <paramref name="namesBefore"/> is the pre-image the parent recorded before the
        /// FIRST attempt; pass <paramref name="namesNow"/> for a first attempt, and the
        /// recorded list for a repeat.
        /// </summary>
        public static DraftAttachmentWork Build(
            IReadOnlyList<string> namesBefore,
            IReadOnlyList<string> namesNow,
            IReadOnlyList<string> addPaths,
            IReadOnlyList<string> removeNames)
        {
            if (namesBefore == null)
            {
                throw new ArgumentNullException(nameof(namesBefore));
            }

            if (namesNow == null)
            {
                throw new ArgumentNullException(nameof(namesNow));
            }

            if (addPaths == null)
            {
                throw new ArgumentNullException(nameof(addPaths));
            }

            if (removeNames == null)
            {
                throw new ArgumentNullException(nameof(removeNames));
            }

            Dictionary<string, int> before = Count(namesBefore);
            Dictionary<string, int> now = Count(namesNow);
            HashSet<string> remove = new HashSet<string>(removeNames, StringComparer.OrdinalIgnoreCase);

            List<string> pathsToAdd = new List<string>();
            List<string> alreadyOn = new List<string>();
            List<DraftAttachmentRemoval> removals = new List<DraftAttachmentRemoval>();

            // Additions are grouped by file name because that is the only identity an
            // attachment has here: two different paths ending in "report.pdf" are one name
            // to Outlook, and the caller reads back names rather than paths.
            foreach (IGrouping<string, string> group in addPaths
                .GroupBy(p => Path.GetFileName(p) ?? string.Empty, StringComparer.OrdinalIgnoreCase))
            {
                List<string> paths = group.ToList();
                if (remove.Contains(group.Key))
                {
                    // Replace: the old copies go, all requested copies come back.
                    pathsToAdd.AddRange(paths);
                    continue;
                }

                int target = Lookup(before, group.Key) + paths.Count;
                int shortfall = target - Lookup(now, group.Key);
                int outstanding = Math.Max(0, Math.Min(shortfall, paths.Count));
                for (int i = paths.Count - outstanding; i < paths.Count; i++)
                {
                    pathsToAdd.Add(paths[i]);
                }

                for (int i = 0; i < paths.Count - outstanding; i++)
                {
                    alreadyOn.Add(group.Key);
                }
            }

            // Removals are counted against the draft AS IT IS NOW, before this call adds
            // anything, so executing them after the additions still takes the pre-existing
            // copies and never the ones just attached.
            List<string> alreadyOff = new List<string>();
            foreach (string name in removeNames.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                int present = Lookup(now, name);
                if (present > 0)
                {
                    removals.Add(new DraftAttachmentRemoval(name, present));
                }
                else if (Lookup(before, name) > 0)
                {
                    // It was on the draft when this operation was first asked for and it is
                    // not now, so the interrupted attempt removed it. The caller asked for
                    // an end state and that part of it holds - reporting it as "not found"
                    // would describe this ATTEMPT rather than the outcome.
                    alreadyOff.Add(name);
                }
            }

            return new DraftAttachmentWork(pathsToAdd, removals, alreadyOn, alreadyOff);
        }

        private static Dictionary<string, int> Count(IReadOnlyList<string> names)
        {
            Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (string name in names)
            {
                if (name == null)
                {
                    continue;
                }

                counts[name] = Lookup(counts, name) + 1;
            }

            return counts;
        }

        private static int Lookup(Dictionary<string, int> counts, string name)
        {
            return counts.TryGetValue(name, out int count) ? count : 0;
        }
    }

    /// <summary>The plan itself: what to attach, what to delete, and what a previous attempt already did.</summary>
    public sealed class DraftAttachmentWork
    {
        /// <summary>Creates the plan.</summary>
        public DraftAttachmentWork(
            IReadOnlyList<string> pathsToAdd,
            IReadOnlyList<DraftAttachmentRemoval> removals,
            IReadOnlyList<string> alreadyAttached,
            IReadOnlyList<string> alreadyRemoved)
        {
            PathsToAdd = pathsToAdd;
            Removals = removals;
            AlreadyAttached = alreadyAttached;
            AlreadyRemoved = alreadyRemoved;
        }

        /// <summary>Requested files that still have to be attached, in request order.</summary>
        public IReadOnlyList<string> PathsToAdd { get; }

        /// <summary>Names to delete, with how many of their lowest-indexed instances.</summary>
        public IReadOnlyList<DraftAttachmentRemoval> Removals { get; }

        /// <summary>Requested files an interrupted attempt already attached - reported as added, because they are.</summary>
        public IReadOnlyList<string> AlreadyAttached { get; }

        /// <summary>Requested names an interrupted attempt already removed - reported as removed, for the same reason.</summary>
        public IReadOnlyList<string> AlreadyRemoved { get; }
    }
}
