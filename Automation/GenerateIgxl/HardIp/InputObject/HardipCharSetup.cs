using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class HardipCharSetup : CharSetup
    {
        public string TestNameInFlow { set; get; }

        public bool IsSplitByVoltage { set; get; }

        public HardipCharSetup()
        {
        }

        // Copy Constructor
        public HardipCharSetup(HardipCharSetup other) : base(other)
        {
            if (other == null)
            {
                return;
            }

            TestNameInFlow = other.TestNameInFlow;
            IsSplitByVoltage = other.IsSplitByVoltage;
        }

        public static string GetShmooName(TestPlanSheet planSheet, HardIpPattern pattern, HardipCharSetup shmoo, string subBlockName, bool isSweep, string originalParameterName)
        {
            string[] shmooNameArr = new string[10] { "X", "X", "X", "X", "X", "X", "X", "X", "X", "X_" };
            string pinName1D = "";
            string pinName2D = "";
            CharStep shmooStep = shmoo.CharSteps[0];
            CharStep xshmoo = shmoo.CharSteps.Where(x => x.Mode == CharSetupConst.ModeXShmoo).FirstOrDefault(y => y.ApplyToPins != "");
            if (xshmoo != null)
            {
                pinName1D = xshmoo.ApplyToPins;
            }

            CharStep yshmoo = shmoo.CharSteps.Where(x => x.Mode == CharSetupConst.ModeYShmoo).FirstOrDefault(y => y.ApplyToPins != "");
            if (yshmoo != null)
            {
                pinName2D = yshmoo.ApplyToPins;
            }

            if (isSweep)  //HAC
            {
                shmooNameArr[0] = "HAC";
            }
            else //HFH,HFL,HIO
            {

                shmooNameArr[0] = ResolveShmooType(planSheet, pattern, shmoo, pinName1D);

                shmooNameArr[1] = GetMeasType(pattern);
                shmooNameArr[3] = subBlockName;
                shmooNameArr[4] = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
                shmooNameArr[5] = GetXPinName(shmoo, pinName1D);
                shmooNameArr[7] = GetYOrParamName(pinName2D, originalParameterName);
            }

            #region UD9
            string block2Name = "";//CommonGenerator.GetSubBlock2Name(pattern.MiscInfo);
            if (block2Name != "")
            {
                shmooNameArr[8] = block2Name;
            }

            #endregion

            for (int i = 0; i < shmooNameArr.Length; i++)
            {
                shmooNameArr[i] = shmooNameArr[i].Replace("_", "");
            }

            return string.Join("_", shmooNameArr);
        }

        public static List<HardipCharSetup> GetShmoo(HardIpPattern pattern)
        {
            List<HardipCharSetup> result = new List<HardipCharSetup>();
            if (pattern.Shmoo.IsSplitByVoltage)
            {
                //NV
                HardipCharSetup newSetupNv = new HardipCharSetup
                {
                    SetupName = CommonGenerator.GetSubBlockNameWithoutMinus(pattern.Shmoo.SetupName) + "_" + "NV",
                    TestMethod = pattern.Shmoo.TestMethod,
                    CharSteps = pattern.Shmoo.CharSteps.Where(x => x.VoltageType == "NV" || x.VoltageType == "").ToList()
                };
                result.Add(newSetupNv);
                //LV
                HardipCharSetup newSetupLv = new HardipCharSetup
                {
                    SetupName = CommonGenerator.GetSubBlockNameWithoutMinus(pattern.Shmoo.SetupName) + "_" + "LV",
                    TestMethod = pattern.Shmoo.TestMethod,
                    CharSteps = pattern.Shmoo.CharSteps.Where(x => x.VoltageType == "LV" || x.VoltageType == "").ToList()
                };
                result.Add(newSetupLv);
                //HV
                HardipCharSetup newSetupHv = new HardipCharSetup
                {
                    SetupName = CommonGenerator.GetSubBlockNameWithoutMinus(pattern.Shmoo.SetupName) + "_" + "HV",
                    TestMethod = pattern.Shmoo.TestMethod,
                    CharSteps = pattern.Shmoo.CharSteps.Where(x => x.VoltageType == "HV" || x.VoltageType == "").ToList()
                };
                result.Add(newSetupHv);
            }
            else
            {
                HardipCharSetup newSetup = new HardipCharSetup();
                string charName = CommonGenerator.GetSubBlockNameWithoutMinus(pattern.Shmoo.SetupName);
                newSetup.SetupName = charName;
                newSetup.TestMethod = pattern.Shmoo.TestMethod;
                newSetup.CharSteps = pattern.Shmoo.CharSteps.Select(x => x.Copy()).ToList();
                result.Add(newSetup);
            }
            return result;
        }

        public static string GetShmooParameterName(string name)
        {
            Dictionary<string, string> hadrCodedic = new Dictionary<string, string> { { "d0", "On" }, { "d1", "Data" }, { "d2", "Retrun" }, { "d3", "Off" } };
            if (CharSetupConst.ParameterName.TryGetValue(name, out string value))
            {
                name = value;
            }

            if (hadrCodedic.ContainsKey(name.ToLower()))
            {
                return hadrCodedic[name.ToLower()];
            }

            return name;
        }

        public static string GetShmooTimeSets(string name)
        {
            if (name.Contains(","))
            {
                List<string> arr = name.Split(',').ToList();
                arr.RemoveAt(0);
                return string.Join(",", arr);
            }
            return "";
        }

        private static string ResolveShmooType(
            TestPlanSheet planSheet,
            HardIpPattern pattern,
            HardipCharSetup shmoo,
            string pinName1D)
        {
            PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            if (pinMap == null)
            {
                return "UNKNOWN";
            }

            bool allIo = true;
            bool hasPower = false;

            foreach (CharStep step in shmoo.CharSteps)
            {
                Pin pin = GetCharPin(pinMap, step.ApplyToPins);

                if (pin == null)
                {
                    continue;
                }

                string type = pin.PinType;

                if (type.Equals(PinMapConst.TypePower, StringComparison.CurrentCultureIgnoreCase) ||
                    step.StepName.StartsWith("VDD"))
                {
                    hasPower = true;
                }

                if (!type.Equals(PinMapConst.TypeIo, StringComparison.CurrentCultureIgnoreCase))
                {
                    allIo = false;
                }
                else
                {
                    step.PostStepArguments = "CorePower";
                }
            }

            if (allIo)
            {
                return "HIO";
            }

            if (hasPower)
            {
                return ResolvePowerDirection(planSheet, pattern, shmoo, pinName1D);
            }

            ReportUnknown(planSheet, pattern, "UNKNOWN");
            return "UNKNOWN";
        }

        private static Pin GetCharPin(PinMapSheet pinMap, string pinName)
        {
            if (pinMap.IsPinExist(pinName))
            {
                return pinMap.GetPin(pinName);
            }

            if (pinMap.IsGroupExist(pinName))
            {
                return pinMap.GetPinsFromGroup(pinName).FirstOrDefault();
            }

            return null;
        }

        private static string ResolvePowerDirection(
            TestPlanSheet planSheet,
            HardIpPattern pattern,
            HardipCharSetup shmoo,
            string pinName)
        {
            CharStep step = shmoo.CharSteps[0];

            string from = DataConvertor.ConvertUnits(step.RangeFrom);
            string to = DataConvertor.ConvertUnits(step.RangeTo);

            bool okFrom = double.TryParse(from, out double fromValue);
            bool okTo = double.TryParse(to, out double toValue);

            if (okFrom && okTo)
            {
                return fromValue > toValue ? "HFL" : "HFH";
            }

            int idx = planSheet.PlanHeaderIdx["forceIndex"];
            ErrorReportManager.AddError(
                HardIpErrorType.E_WrongForceCondition_09,
                pattern.SheetName,
                pattern.RowNum,
                idx,
                [pinName]
            );

            return "UNKNOWN";
        }

        private static void ReportUnknown(
            TestPlanSheet planSheet,
            HardIpPattern pattern,
            string pinType)
        {
            int forceIndex = planSheet.PlanHeaderIdx["forceIndex"];

            ErrorReportManager.AddError(
                HardIpErrorType.E_WrongForceCondition_10,
                pattern.SheetName,
                pattern.RowNum,
                forceIndex,
                [pinType]
            );
        }

        private static string GetMeasType(HardIpPattern pattern)
        {
            return pattern.MeasPins.Select(x => x.MeasType).FirstOrDefault() ?? "X";
        }

        private static string GetXPinName(HardipCharSetup shmoo, string pin)
        {
            return shmoo.CharSteps.Count(x => x.Mode == CharSetupConst.ModeXShmoo) > 1
                ? "MULTI"
                : pin.Replace(",", "");
        }

        private static string GetYOrParamName(string yPin, string param)
        {
            if (!string.IsNullOrEmpty(yPin))
            {
                return yPin.Replace("_", "");
            }

            return param.Replace("_", "").Replace(",", "");
        }

    }
}
