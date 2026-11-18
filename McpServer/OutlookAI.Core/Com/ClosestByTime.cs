using System;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// The candidate nearest in time to an index row, among those within the tolerance - what the index-hit
    /// locator's probe returns, instead of the first candidate within the tolerance. Items of one subject made
    /// in one batch sit tenths of a second apart (the Exchange guest's recurring holiday appointments, measured
    /// 2026-11-18), so the first one within five seconds can be another item than the row's. Pure.
    /// </summary>
    /// <typeparam name="T">The candidate's snapshot.</typeparam>
    public sealed class ClosestByTime<T>
        where T : class
    {
        private readonly int _toleranceSeconds;

        /// <summary>Creates an empty choice over candidates at most <paramref name="toleranceSeconds"/> away.</summary>
        public ClosestByTime(int toleranceSeconds)
        {
            if (toleranceSeconds < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(toleranceSeconds));
            }

            _toleranceSeconds = toleranceSeconds;
        }

        /// <summary>The nearest candidate offered so far within the tolerance, or null.</summary>
        public T? Best { get; private set; }

        /// <summary>How far <see cref="Best"/> is from the row's time, in seconds; null without one.</summary>
        public double? BestSecondsApart { get; private set; }

        /// <summary>
        /// Offers a candidate <paramref name="secondsApart"/> from the row's time (null: it has no time). Returns
        /// true when it is the row's own time to within <see cref="HitLocator.ExactMatchSeconds"/>, so looking
        /// further cannot find a nearer one; an earlier candidate keeps a tie.
        /// </summary>
        public bool Offer(T candidate, double? secondsApart)
        {
            if (candidate == null)
            {
                throw new ArgumentNullException(nameof(candidate));
            }

            if (!secondsApart.HasValue || secondsApart.Value > _toleranceSeconds)
            {
                return false;
            }

            if (Best == null || secondsApart.Value < BestSecondsApart!.Value)
            {
                Best = candidate;
                BestSecondsApart = secondsApart.Value;
            }

            return secondsApart.Value <= HitLocator.ExactMatchSeconds;
        }
    }
}
