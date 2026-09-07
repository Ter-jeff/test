using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.InputObject;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    internal class JitterGenerator
    {
#nullable enable
        public JitterSheet? GenJitterSheet(Dictionary<string, HardIpSheet> planDic, string sheetName = "JitterSheet")
#nullable restore
        {
            var allMeasDPins = new List<string>();

            foreach (string planKey in planDic.Keys)
            {
                HardIpSheet hipPatternList = planDic[planKey];
                foreach (HardIpPattern pins in hipPatternList.Rows.Where(pat => pat.MeasPins.Any(p => p.MeasType.Equals(MeasType.MeasD))))
                {
                    foreach (MeasPin pin in pins.MeasPins)
                    {
                        allMeasDPins.Add(pin.PinName);
                    }
                }
            }
            if (!allMeasDPins.Any())
            {
                return null;
            }

            var jitterSheet = new JitterSheet(sheetName);
            allMeasDPins = allMeasDPins.Distinct().ToList();
            foreach (string pin in allMeasDPins)
            {
                var row = new JitterRow
                {
                    JitterSet = "Duty_Jitter",
                    PinOrGroup = pin,
                    Mode = "Disabled",
                    SampleType = "One Level",
                    TargetResolution = "=_J_TargetResolution",//"500E-15"
                    TargetPatternReps = "=_J_RepeatTimes",//"500"
                    BitsPerPattern = "2",
                    BitPeriod = "5E-9"
                };
                jitterSheet.AddRow(row);
            }
            return jitterSheet;

        }
    }
}
