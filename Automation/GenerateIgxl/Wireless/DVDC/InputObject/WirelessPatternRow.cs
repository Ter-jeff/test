using System.Diagnostics.CodeAnalysis;

using Automation.GenerateIgxl.HardIp.InputObject;

namespace Automation.GenerateIgxl.Wireless.DVDC.InputObject
{
    [ExcludeFromCodeCoverage]
    public class WirelessPatternRow : PatternRow
    {
        #region Property
        public WirelessData WirelessData = new WirelessData();
        #endregion

        public WirelessPatternRow(string patternName = "")
            : base(patternName)
        {
            IsRf = true;
        }
    }
}
