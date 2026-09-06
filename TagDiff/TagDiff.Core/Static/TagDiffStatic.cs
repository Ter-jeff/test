using System.Collections.Generic;

using CommonLib.Extension;

namespace TagDiff.Core.Static
{
    public class TagDiffStatic
    {
        public static readonly HashSet<string> NotTracks = new(StringExtensions.IgnoreCase) { "Output_hashes", "ExecInfo", "Versions" };
        private static readonly object _lock = new();
        private static int _returnValue;
        public static int ReturnValue
        {
            get
            {
                lock (_lock)
                {
                    return _returnValue;
                }
            }
            set
            {
                lock (_lock)
                {
                    _returnValue = value;
                }
            }
        }

        public static bool IsNotTracked(string sheetName)
        {
            lock (_lock)
            {
                return NotTracks.Contains(sheetName);
            }
        }

        public static void AddNotTracked(string sheetName)
        {
            lock (_lock)
            {
                NotTracks.Add(sheetName);
            }
        }

        public static void Initialize()
        {
            _returnValue = 0;
            lock (_lock)
            {
                NotTracks.Clear();
                NotTracks.UnionWith(["Output_hashes", "ExecInfo", "Versions"]);
            }
        }
    }
}
