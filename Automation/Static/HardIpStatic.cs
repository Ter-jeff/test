using System.Collections.Generic;

namespace Automation.Static
{
    public static class HardIpStatic
    {
        public static List<string> FlowUsedInteger { get; set; } = new List<string>();

        internal static void Clear()
        {
            FlowUsedInteger = new List<string>();
        }
    }
}
