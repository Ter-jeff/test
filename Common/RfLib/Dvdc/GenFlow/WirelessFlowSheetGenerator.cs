using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlow;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using CommonLib.Enums;
using CommonLib.Extension;

using RfLib.InstrumentSetup;
using RfLib.Static;

namespace RfLib.Dvdc.GenFlow
{
    public partial class WirelessFlowSheetGenerator : HardIpFlowSheetGenerator
    {
        [GeneratedRegex("IsReadCapTrim")]
        private static partial Regex MyRegex();

        private static readonly Dictionary<string, int> _testCountDictionary = new()
        {
            { "psat", 3 }, { "iqmm", 1 }, { "loft", 1 }
            , { "pwrdiff",3 }
            , { "pn",999 }
            , { "spur",999 }
        };
        public WirelessFlowSheetGenerator(HardIpInputData hardIpInputData, string sheetName, string label, List<HardIpPattern> hardIpPatterns)
            : base(hardIpInputData, sheetName, hardIpPatterns)
        {
            if (label.Length != 0)
            {
                SheetName = sheetName + "_" + label;
            }

            FlowRowGenerator = new WirelessHardIpFlowRowGenerator(hardIpInputData, SheetName);
            Volatage = label;
        }

        //protected override void Initial()
        //{
        //    FlowRowGenerator = new WirelessHardIpFlowRowGenerator(SheetName);
        //}

        protected override List<HardIpPattern> DividePatterns()
        {
            foreach (HardIpPattern pattern in PatternList)
            {
                foreach (MeasPin pin in pattern.MeasPins)
                {
                    if (Volatage == "HV" || Volatage.Length == 0)
                    {
                        UpdateLimits(pattern.UseLimitsH, pin, "H");
                    }

                    if (Volatage == "LV" || Volatage.Length == 0)
                    {
                        UpdateLimits(pattern.UseLimitsL, pin, "L");
                    }

                    if (Volatage == "NV" || Volatage.Length == 0)
                    {
                        UpdateLimits(pattern.UseLimitsN, pin, "N");
                    }
                }
                PostProcess(pattern);
            }
            return PatternList;
        }

        protected static void PostProcess(HardIpPattern hardIpPattern)
        {
            if (hardIpPattern.WirelessData != null)
            {
                if (!string.IsNullOrEmpty(hardIpPattern.WirelessData.TrimTarget) || MyRegex().IsMatch(hardIpPattern.MiscInfo))
                {
                    var ignoreList = new List<string> { "BSTC", "BSTV", "VRFV", "VRFC" };
                    hardIpPattern.UseLimitsH.RemoveAll(p => ignoreList.All(q => !p.TestName.Contains(q)));
                    hardIpPattern.UseLimitsN.RemoveAll(p => ignoreList.All(q => !p.TestName.Contains(q)));
                    hardIpPattern.UseLimitsL.RemoveAll(p => ignoreList.All(q => !p.TestName.Contains(q)));
                }
            }
        }

        public static void UpdateLimits(List<MeasPin> measPins, MeasPin measPin, string type)
        {
            var excludedTypes = new HashSet<string>(StringExtensions.IgnoreCase)
            {
                "MeasLoopBack",
                "MeasLoopBack_Dis"
            };
            var copy = new MeasPin();
            if (copy.PinName.Split('=').Length == 2)
            {
                IEnumerable<string> relatedJobs = Automation.Static.LocalSpecs.AllJobs.ToList()
                        .Where(p => Regex.IsMatch(p, measPin.PinName.Split('=')[0], RegexOptions.IgnoreCase));
                copy.Job = string.Join(",", relatedJobs);
            }
            else
            {
                copy.Job = measPin.Job;
            }

            copy.MeasType = measPin.MeasType;
            copy.RowNumForMergeMeas = measPin.RowNumForMergeMeas;
            copy.SequenceIndex = measPin.SequenceIndex;
            copy.ForceConditions = measPin.ForceConditions;
            copy.RowNum = measPin.RowNum;
            if (Automation.Static.LocalSpecs.Options.Device == EnumDevice.RF)
            {
                string[] spiTn = measPin.TestName.Split('_');
                if (spiTn.Length > 8)
                {
                    spiTn[8] = type + "V";
                }

                measPin.TestName = string.Join("_", spiTn);
            }
            copy.TestName = measPin.TestName;
            copy.PinName = measPin.PinName;
            copy.MiscInfo = measPin.MiscInfo;
            copy.CalcEqn = measPin.CalcEqn;
            List<MeasLimit> measlimits = [];
            switch (type.ToUpper())
            {
                case "H":
                    measlimits = measPin.MeasLimitsH;
                    break;
                case "L":
                    measlimits = measPin.MeasLimitsL;
                    break;
                case "N":
                    measlimits = measPin.MeasLimitsN;
                    break;
            }
            if (measlimits.Count > 0)
            {
                copy.LowLimit = measlimits[0].LoLimit;
                copy.HighLimit = measlimits[0].HiLimit;
            }
            if (measPin.MeasType != MeasType.WiSrc)
            {
                foreach (string testtype in measPin.InterPoseFunc.Split(';'))
                {
                    int count = 1;
                    if (_testCountDictionary.ContainsKey(testtype))
                    {
                        if (_testCountDictionary[testtype] != 999)
                        {
                            count = _testCountDictionary[testtype];
                        }
                        else
                        {
                            count = GetSpecialLimitFromInstrument(measPin.RfInstrumentSetup);
                        }
                    }
                    for (int i = 0; i < count; i++)
                    {
                        if (!excludedTypes.Contains(measPin.MeasType))
                        {
                            measPins.Add(copy);
                        }
                    }
                }
            }
        }

        public static int GetSpecialLimitFromInstrument(string setup)
        {
            int count = 0;
            string[] parts = setup.Split('=');
            InstrumentConfigRow? instruments = null;
            if (parts.Length > 1 && SettingStatic.InstrumentSheet != null)
            {
                instruments = SettingStatic.InstrumentSheet.GetInstrumentType(parts[1].Trim());
            }
            if (instruments != null)
            {
                Dictionary<string, string>? info = instruments.GetTypeDictionary(instruments.Instrumenttype);
                if (info != null && info.TryGetValue("offsetfreq", out string? value) && value != null)
                {
                    count = value.Split('&').Length;
                }
            }

            return count;
        }
    }
}
