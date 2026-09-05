using System;

namespace CommonLib.Utility
{
    public sealed class UnitUtility
    {
        private static readonly Lazy<UnitUtility> _lazyInstance = new(() => new UnitUtility());
        public static UnitUtility Instance { get { return _lazyInstance.Value; } }

        private UnitUtility() { }

        /// <summary>
        /// Gets the numeric scale for a unit string. All unassigned strings return 0.
        /// </summary>
        public static int GetScale(string unit)
        {
            if (string.IsNullOrEmpty(unit))
            {
                return 0;
            }

            return unit switch
            {
                "f" => -15,
                "p" => -12,
                "n" => -9,
                "u" => -6,
                "m" => -3,
                "%" => -2,
                "k" or "K" => 3,
                "M" => 6,
                "G" => 9,
                "T" => 12,
                _ => 0,
            };
        }
    }
}
