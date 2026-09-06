using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow
{
    public class HardIpFlowRowGenerator : FlowRowGeneratorBase
    {
        public HardIpFlowRowGenerator(HardIpInputData hardIpInputData, string sheetName) : base(hardIpInputData, sheetName)
        {
        }

        public override List<FlowRow> GenRunScenarioRows(string bootinst = "", bool isFirstScenario = false)
        {
            return Enumerable.Empty<FlowRow>().ToList();
        }

        protected override void SetBasicInfoByPattern(HardIpPattern pattern)
        {
            BlockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
        }
        [Flags]
        public enum BinningType
        {
            None = 0,
            Low = 1,
            High = 2
        }
        private static readonly Dictionary<BinningType, string> _binningMap = new()
        {
            { BinningType.Low, "BinningL" },
            { BinningType.High, "BinningH" },
            { BinningType.Low | BinningType.High, "BinningHL" }
        };

        public override List<FlowRow> GenTestRows(bool isCz2Only = false)
        {
            var flowRows = new List<FlowRow>();
            var testFlowRow = new FlowRow();
            bool noBinoutForAll = HardIpInputData.ConfigData.NoBinOutForAll;
            testFlowRow.Job = CreateTestJob();
            testFlowRow.Part = CreateTestPart();
            testFlowRow.Opcode = CreateTestOpcode();
            testFlowRow.Env = CreateTestEnv();
            testFlowRow.Enable = GenEnable(CreateTestEnable(isCz2Only), testFlowRow.Env);
            testFlowRow.Parameter = CreateTestParameter();
            testFlowRow.Comment1 = SubBlockName + Pat.ForceVoltageFlag;

            if (Regex.IsMatch(Pat.MiscInfo, HardIpConstData.MappingTtrBranch, RegexOptions.IgnoreCase))
            {
                string mapTtrKey = HardIpService.GetMappingTtrBranchItem(Pat.MiscInfo);
                if (InitOriginalItems.TryGetValue(mapTtrKey, out string item))
                {
                    testFlowRow.FailAction = item;
                }
            }
            else
            {
                testFlowRow.FailAction = CreateTestFailAction();
            }

            List<FlowRow> limitRows = SortFlowRows(GetLimitRows(testFlowRow));
            List<FlowRow> ifFlowRows = GetIfFlowRows(CreateIfConditionByTestPanDefine(), testFlowRow, limitRows);
            flowRows.AddRange(ifFlowRows);

            if (Regex.IsMatch(Pat.MiscInfo, HardIpConstData.OriginalTtrBranch, RegexOptions.IgnoreCase))
            {
                string oriTtrKey = HardIpService.GetOriginalTtrBranchItem(Pat.MiscInfo);

                var oriItemClearRows = new List<FlowRow>
                {
                    new FlowRow { Opcode = CreateIfOpcode(), Parameter = CreateTestFailAction() },
                    new FlowRow { Opcode = CreateFlagClearOpcode(), Parameter = CreateTestFailAction() },
                    new FlowRow { Opcode = CreateEnableFlowWdOpcode(), Parameter = oriTtrKey },
                    new FlowRow { Opcode = CreateEndIfOpcode() }
                };
                flowRows.AddRange(oriItemClearRows);

                InitOriginalItems[oriTtrKey] = CreateTestFailAction(); //Overwrite to follow NV/LV/HV
            }

            if (noBinoutForAll || Pat.NoBinoutStr.Contains(ActualLabelVoltage))
            {
                testFlowRow.Comment1 = testFlowRow.Comment1 + "\t" + testFlowRow.FailAction;
                testFlowRow.FailAction = "";
            }

            return flowRows;
        }

        public override List<FlowRow> GetIfFlowRows(string dataRow, FlowRow testFlowRow, List<FlowRow> uselimitFlowRows)
        {
            var ifFlowRows = new List<FlowRow>();
            ifFlowRows.Add(testFlowRow);
            if (uselimitFlowRows != null)
            {
                ifFlowRows.AddRange(uselimitFlowRows);
            }
            if (!string.IsNullOrEmpty(dataRow))
            {
                if (!string.IsNullOrEmpty(testFlowRow.DeviceName) || dataRow.Contains("&&") || dataRow.Contains("||"))
                {
                    ifFlowRows.Clear();
                    ifFlowRows.Add(FlowRow.GenIfCondition(dataRow, testFlowRow.Job));
                    ifFlowRows.Add(testFlowRow);
                    if (uselimitFlowRows != null)
                    {
                        ifFlowRows.AddRange(uselimitFlowRows);
                    }

                    ifFlowRows.Add(FlowRow.GenEndIf(testFlowRow.Job));
                    return ifFlowRows;
                }

                if (string.IsNullOrEmpty(testFlowRow.DeviceName))
                {
                    AddCondition(dataRow, ref testFlowRow, ref uselimitFlowRows);
                }
            }
            return ifFlowRows;
        }

        protected void AddCondition(string siteVar, ref FlowRow testFlowRow, ref List<FlowRow> useLimitFlowRows)
        {
            string devName = siteVar.TrimStart('!');
            testFlowRow.DeviceName = devName;
            useLimitFlowRows.ForEach(ulfr => ulfr.DeviceName = devName);
            if (!siteVar.Trim().EndsWith("False", StringComparison.CurrentCultureIgnoreCase))
            {
                string flagStatus = siteVar.StartsWith("!") ? "Flag-false" : "Flag-true";
                testFlowRow.DeviceCondition = flagStatus;
                useLimitFlowRows.ForEach(ulfr => ulfr.DeviceCondition = flagStatus);
            }
            else
            {
                string flagStatus = siteVar.StartsWith("!") ? "Flag-true" : "Flag-false";
                testFlowRow.DeviceCondition = flagStatus;
                useLimitFlowRows.ForEach(ulfr => ulfr.DeviceCondition = flagStatus);
            }
        }

        protected virtual List<FlowRow> GetLimitRows(FlowRow testRow, bool forShmoo = false)
        {
            string useLimitFailAction = "";
            if (Regex.IsMatch(Pat.MiscInfo, HardIpConstData.MappingTtrBranch, RegexOptions.IgnoreCase))
            {
                string mapTtrKey = HardIpService.GetMappingTtrBranchItem(Pat.MiscInfo);
                if (InitOriginalItems.TryGetValue(mapTtrKey, out string item))
                {
                    useLimitFailAction = item;
                }
            }
            else
            {
                //use limit rows
                useLimitFailAction = CreateUseLimitFailAction(forShmoo);
            }
            return GenUseLimitRows(testRow.Parameter, useLimitFailAction);
        }

        public override List<FlowRow> GenShmooRows(string labelVoltage = "")
        {
            var testRows = new List<FlowRow>();
            if (!HasShmoo)
            {
                return testRows;
            }

            var testFlowRow = new FlowRow
            {
                Env = CreateTestEnv(),
                Enable = Pat.ForceCondition.IsVtShmoo ? "" : "!TestOnly",
                Opcode = CreateCharacterizeOpcode(),
                Parameter = Pat.Shmoo.IsSplitByVoltage ? CreateShmooParameter() + "_" + labelVoltage : CreateShmooParameter(),
                TName = CreateShmooTName()
            };
            testRows.Add(testFlowRow);
            List<FlowRow> limitRows = SortFlowRows(GetLimitRows(testFlowRow, forShmoo: true));
            testRows.AddRange(limitRows);
            return testRows;
        }

        public override FlowRow GenBinTableRow(string voltage = "")
        {
            var bintableRow = new FlowRow
            {
                Job = CreateBinTableJob(),
                Env = CreateBinTableEnv(),
                Opcode = CreateBinTableOpcode(),
                Enable = CreateBinTableEnable(),
                Parameter = CreateBinTableParameter()
            };
            return bintableRow;
        }

        public override List<FlowRow> GenPreRetestRows()
        {
            var preRestRows = new List<FlowRow>();
            if (!Pat.MiscInfoDict.ContainsKey(HardIpConstData.ReTest))
            {
                return preRestRows;
            }

            // Clear Flag
            preRestRows.Add(GenRetestClearFlagRow());
            // Test Row 
            var testRow = new FlowRow
            {
                Job = CreateTestJob(),
                Opcode = CreateTestOpcode(),
                Enable = EnableWord,
                Env = CreatePreTestEnv(),
                Parameter = CreatePreTestParemeter(),
                FailAction = CreatePreTestFailAction()
            };
            preRestRows.Add(testRow);
            // Use Limit Rows
            preRestRows.AddRange(GenUseLimitRows(testRow.Parameter, testRow.FailAction));
            return preRestRows;
        }

        protected internal virtual List<FlowRow> GenUseLimitRows(string parameter, string binFail)
        {
            bool noBinoutForAll = HardIpInputData.ConfigData.NoBinOutForAll;
            var flowUseLimitRows = new List<FlowRow>();
            #region Generate Use-Limit rows
            string repeatStr = HardIpService.GetRepeatMapping(Pat.MiscInfo);
            List<MeasPin> useLimits;
            switch (ActualLabelVoltage)
            {
                case HardIpConstData.LabelHv:
                    useLimits = Pat.UseLimitsH;
                    break;
                case HardIpConstData.LabelLv:
                    useLimits = Pat.UseLimitsL;
                    break;
                case HardIpConstData.LabelNv:
                    useLimits = Pat.UseLimitsN;
                    break;
                default:
                    useLimits = Pat.UseLimitsN;
                    break;
            }
            //Gets the result that whether the pingroup need to be decomposed for V,F,I or V,I,R
            SearchInfo.GetTestLimitPerMeasType(Pat);
            bool isSingleLimit = SearchInfo.GetFlagSingleLimit(Pat, ActualLabelVoltage);

            var dicUseLimits = useLimits.GroupBy(p => p.SequenceIndex).ToDictionary(p => p.Key, p => p.ToList());
            SearchInfo.GetStoreNameOri(Pat).Split('+').ToList();
            foreach (KeyValuePair<int, List<MeasPin>> item in dicUseLimits)
            {
                int maxForceCnt = 1;

                SortLimitPin useLimitPin = BuildUseLimitPin(item.Value, isSingleLimit);
                List<SortLimitPin> sortUseLimitPins = SortUseLimitPins(useLimitPin, item.Value);

                for (int cnt = 0; cnt < maxForceCnt; cnt++)
                {
                    foreach (SortLimitPin pin in sortUseLimitPins)
                    {
                        foreach (string repeat in repeatStr.Split(','))
                        {
                            flowUseLimitRows.Add(BuildUseLimitRow(pin, repeat, cnt, parameter, binFail, noBinoutForAll));
                        }
                    }
                }
            }
            #endregion
            return flowUseLimitRows;
        }

        private SortLimitPin BuildUseLimitPin(List<MeasPin> pins, bool isSingleLimit)
        {
            var useLimitPin = new SortLimitPin();
            foreach (MeasPin pin in pins)
            {
                if ((pin.MeasType.Equals(MeasType.MeasC, StringComparison.OrdinalIgnoreCase) &&
                     pin.TestName.Equals("skip", StringComparison.OrdinalIgnoreCase)) ||
                    pin.MeasType.Equals(MeasType.MeasN, StringComparison.OrdinalIgnoreCase) ||
                    pin.MiscInfo.ContainsIgnoreCase(HardIpConstData.IgnoreFlowLimit))
                {
                    continue;
                }

                string name = pin.PinName.Replace(" ", "");
                if (name == "")
                {
                    pin.PinName = "";
                }

                var nameList = new List<string>();
                if ((pin.MeasType.ToLower() == "measf" && name.Contains("::")) || pin.MeasType.ToLower() == "measfdiff")
                {
                    string groupName = SearchInfo.GenDiffGroupName(name, true);
                    nameList.Add(groupName);
                }
                else if (name.Contains("::") && !pin.MeasType.ContainsIgnoreCase("diff") &&
                         pin.MeasType.ToLower() != "measvocm" && pin.MeasType.ToLower() != "measvdm")
                {
                    nameList = Regex.Split(name, "::").ToList();
                }
                else if (pin.MeasType.ContainsIgnoreCase("diff2"))
                {
                    string[] diff2PinPairs = name.Split(new[] { "::" }, StringSplitOptions.None);
                    for (int diffIdx = 0; diffIdx < diff2PinPairs.Length; diffIdx += 2)
                    {
                        string diff2PinPair = $"{diff2PinPairs[diffIdx]}::{(diffIdx + 1 < diff2PinPairs.Length ? diff2PinPairs[diffIdx + 1] : "")}";
                        nameList.Add(SearchInfo.GenDiffGroupName(diff2PinPair, true));
                    }
                }
                else
                {
                    if (isSingleLimit && LocalSpecs.Options.Device != EnumDevice.LCD && LocalSpecs.Options.Device != EnumDevice.RF)
                    {
                        nameList.Add(name);
                    }
                    else
                    {
                        List<string> deComposedGroup = TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(name);
                        if (deComposedGroup.Count == 2 && deComposedGroup.Any(x => x.Equals($"{name}_DCVI", StringComparison.OrdinalIgnoreCase)) && deComposedGroup.Any(x => x.Equals($"{name}_DCVS", StringComparison.OrdinalIgnoreCase)))
                        {
                            nameList.Add(name);
                        }
                        else
                        {
                            nameList.AddRange(TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(name));
                        }
                    }
                }

                foreach (string pinName in nameList)
                {
                    useLimitPin.AddData(pinName, pin);
                }
            }

            return useLimitPin;
        }

        private List<SortLimitPin> SortUseLimitPins(SortLimitPin useLimitPin, List<MeasPin> itemValue)
        {
            var sortUseLimitPins = new List<SortLimitPin>();
            int forceCount = itemValue.Exists(p => p.MeasType == MeasType.MeasR2) ?
                    1 : itemValue.Max(p => p.ForceConditions.Count) > 1 ?
                    itemValue.Max(p => p.ForceConditions.Count) : 1;
            int originPinCount = useLimitPin.UseLimitPins.Count / forceCount;
            for (int j = 0; j < forceCount; j++)
            {

                //Calc items
                var calcLimits =
                    useLimitPin.UseLimitPins.Where(measPin =>
                    measPin.MeasPinData.MeasType == MeasType.MeasCalc ||
                    measPin.MeasPinData.MeasType == MeasType.MeasC ||
                    measPin.MeasPinData.MeasType == MeasType.MeasLimit).ToList();

                //Sorting
                var tmpSortLimits = useLimitPin.UseLimitPins.GetRange(originPinCount * j, originPinCount)
                    .GroupBy(p => p.PinNameWoStage() + p.MeasPinData.MeasType + p.MeasPinData.HighLimit + p.MeasPinData.LowLimit)
                    .Select(p => p.Any(q => !q.PinName.Contains('=')) ? p.First(q => !q.PinName.Contains("=")) : p.First())
                    .OrderBy(p => p.PinNameWoStage())
                    .ToList();

                //Vdiff items
                var measVDiffLimits = tmpSortLimits.Where(measPin =>
                    measPin.MeasPinData.MeasType == MeasType.MeasVdiff)
                    .OrderBy(measPin => measPin.PinNameWoStage()).ToList();

                //Vocm items
                var measVocmLimits = tmpSortLimits.Where(measPin =>
                    measPin.MeasPinData.MeasType == MeasType.MeasVocm)
                    .OrderBy(measPin => measPin.PinNameWoStage()).ToList();

                tmpSortLimits.RemoveAll(calcLimits.Contains);
                tmpSortLimits.RemoveAll(measVDiffLimits.Contains);
                tmpSortLimits.RemoveAll(measVocmLimits.Contains);

                sortUseLimitPins.AddRange(tmpSortLimits);
                sortUseLimitPins.AddRange(measVDiffLimits);
                sortUseLimitPins.AddRange(measVocmLimits);
                sortUseLimitPins.AddRange(calcLimits);
            }

            return sortUseLimitPins;
        }

        private FlowRow BuildUseLimitRow(SortLimitPin pin, string repeat, int cnt, string parameter, string binFail, bool noBinoutForAll)
        {
            var row = new FlowRow();
            string lowUnit;
            string lowScale;
            string highUnit;
            string highScale;
            row.Opcode = CreateUseLimitOpcode();
            row.Job = pin.MeasPinData.Job;
            row.Parameter = parameter;
            if (!string.IsNullOrEmpty(pin.MeasPinData.TestName))
            {
                if (cnt >= pin.MeasPinData.TestName.Split(',').Length)
                {
                    row.TName = pin.MeasPinData.TestName;
                }
                else
                {
                    row.TName = pin.MeasPinData.TestName.Split(',')[cnt];
                }

                if (pin.MeasPinData.TestName.Split('_').Length == 10 && LocalSpecs.Options.Device == EnumDevice.RF)
                {
                    string sgmt8 = pin.MeasPinData.TestName.Split('_')[8];
                    if (!sgmt8.Equals(ActualLabelVoltage, StringComparison.OrdinalIgnoreCase))
                    {
                        pin.MeasPinData.TestName = pin.MeasPinData.TestName.Replace("_NV_", "_" + ActualLabelVoltage + "_");
                    }

                    string sgmt2 = pin.MeasPinData.TestName.Split('_')[2];
                    row.TName = pin.MeasPinData.TestName.Replace(sgmt2, pin.PinName.Replace("_", ""));
                }
            }

            string pinNameString = pin.PinName == "" ? "NoPinName" : pin.PinName;
            pinNameString = pin.MeasPinData.MeasType == MeasType.MeasCalc ? pin.MeasPinData.CalcEqn.Split(':').Last() : pinNameString;
            row.Comment = Pat.Pattern.GetLastPayload() + "_" + pin.MeasPinData.MeasType + "_" +
                          pinNameString;

            if (Regex.IsMatch(Pat.MiscInfo, "isFW", RegexOptions.IgnoreCase))
            {
                row.LoLim = DataConvertor.ConvertUseLimitFw(pin.MeasPinData.LowLimit, out lowUnit, out lowScale);
            }
            else
            {
                row.LoLim = DataConvertor.ConvertUseLimit(pin.MeasPinData.LowLimit, out lowUnit, out lowScale);
            }

            if (repeat != "")
            {
                if (row.LoLim.Contains("="))
                {
                    row.LoLim = row.LoLim.Replace("=", "=(") + ")" + "*" + repeat.Replace("x", "");
                }
                else
                {
                    row.LoLim = HardIpService.CalcuteLimit(row.LoLim, repeat);
                }
            }
            if (Regex.IsMatch(Pat.MiscInfo, "isFW", RegexOptions.IgnoreCase))
            {
                row.HiLim = DataConvertor.ConvertUseLimitFw(pin.MeasPinData.HighLimit, out highUnit, out highScale);
            }
            else
            {
                row.HiLim = DataConvertor.ConvertUseLimit(pin.MeasPinData.HighLimit, out highUnit, out highScale);
            }

            if (repeat != "")
            {
                if (row.HiLim.Contains("="))
                {
                    row.HiLim = row.HiLim.Replace("=", "=(") + ")" + "*" + repeat.Replace("x", "");
                }
                else
                {
                    row.HiLim = HardIpService.CalcuteLimit(row.HiLim, repeat);
                }
            }
            row.Scale = lowScale == "" ? highScale : lowScale;

            row.Units = lowUnit == "" ? highUnit : lowUnit;
            if (noBinoutForAll || Pat.NoBinoutStr.Contains(ActualLabelVoltage))
            {
                row.Comment1 = "\t" + binFail;
            }
            else
            {
                row.FailAction = binFail;
            }

            row.PatName = pin.MeasPinData.PatternName;
            row.PatIndex = pin.MeasPinData.PatternIndex;
            return row;
        }


        protected List<FlowRow> GenUseLimits(
            string parameter,
            HardIpPattern pattern,
            string voltage,
            Func<string, string> createFailAction,
            List<string> jobs = null
        )
        {
            var useLimits = new List<FlowRow>();
            string[] testJobs = SearchInfo.GetTtrEnable(Pat.TtrStr, LabelVoltage, jobs)
                .ToUpper()
                .Split(',');
            string repeatStr = HardIpService.GetRepeatMapping(pattern.MiscInfo);
            var corePowers = new List<string>();
            if (TestProgram.IgxlWorkBk.PinMapPair.Value.TryGetGroup("COREPOWER", out PinGroup groupPin))
            {
                corePowers = groupPin.PinList.Select(x => x.PinName).ToList();
            }

            //Gets the result that whether the pingroup need to be decomposed for V,F,I or V,I,R
            Dictionary<string, string> dicTestLimitPerMeasType = SearchInfo.GetTestLimitPerMeasType(pattern);
            foreach (MeasPin pin in pattern.UseLimitsN)
            {
                var finalLimitJobs = new List<string>();
                if (pin.Job != "")
                {
                    string[] limitJobs = pin.Job.ToUpper().Split(',');
                    finalLimitJobs = testJobs.Intersect(limitJobs).ToList();
                    if (!finalLimitJobs.Any())
                    {
                        continue;
                    }
                }

                if (pin.MeasType.Equals(MeasType.MeasC, StringComparison.OrdinalIgnoreCase) && pin.TestName.Equals("skip", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                List<string> nameList = BuildUseLimitsNameList(pattern, pin, dicTestLimitPerMeasType);

                foreach (string pinName in nameList)
                {
                    string repeatSetting = "";
                    //Only repeat limit for CorePower Pins
                    if (corePowers.Contains(pinName) || pinName.Equals("COREPOWER"))
                    {
                        repeatSetting = repeatStr;
                    }

                    foreach (string repeat in repeatSetting.Split(','))
                    {
                        FlowRow row = BuildUseLimitsRow(pattern, pin, pinName, finalLimitJobs, parameter, repeat, createFailAction);
                        #region Add flag to Mbist

                        string repStr = repeat == "" ? "1x" : repeat;

                        IdsRepairSingleton.Instance().AddIdsFlag(row.FailAction, repStr);
                        #endregion
                        //Set flow row comment
                        useLimits.Add(row);
                    }
                }
            }
            var sortedFlowRows = new List<FlowRow>();
            List<FlowRow> limitRows = SortFlowRows(useLimits);
            sortedFlowRows.AddRange(limitRows);

            return sortedFlowRows;
        }

        private List<string> BuildUseLimitsNameList(HardIpPattern pattern, MeasPin pin, Dictionary<string, string> dicTestLimitPerMeasType)
        {
            var nameList = new List<string>();
            string name = pin.PinName.Replace(" ", "");

            #region Change by laura at 20161006
            if (!TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(name) || (TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(name) && !SearchInfo.NeedDecompGroups(pattern, pin, dicTestLimitPerMeasType) && LocalSpecs.Options.Device != EnumDevice.RF))
            {
                nameList.Add(name);
            }
            else
            {
                if (TestProgram.IgxlWorkBk.PinMapPair.Value.TryGetGroup(name, out PinGroup group))
                {
                    IEnumerable<string> pins = group.PinList.Select(x => x.PinName);
                    nameList.AddRange(pins);
                }
            }

            #endregion
            return nameList;
        }

        private FlowRow BuildUseLimitsRow(HardIpPattern pattern, MeasPin pin, string pinName, List<string> finalLimitJobs, string parameter, string repeat, Func<string, string> createFailAction)
        {
            var row = new FlowRow
            {
                Opcode = CreateUseLimitOpcode(),
                Job = string.Join(",", finalLimitJobs),
                Parameter = parameter,
                Comment = pattern.Pattern.GetLastPayload() + "_" + pin.MeasType + "_" + pinName
            };
            if (row.Comment.Contains("__"))
            {
            }
            if (!string.IsNullOrEmpty(pin.TestName))
            {
                row.TName = pin.TestName;
            }

            bool isLow = !string.IsNullOrEmpty(pin.LowLimit) && pin.LowLimit.StartsWith("Binning", StringComparison.CurrentCultureIgnoreCase);

            bool isHigh = !string.IsNullOrEmpty(pin.HighLimit) && pin.HighLimit.StartsWith("Binning", StringComparison.CurrentCultureIgnoreCase);

            BinningType type = BinningType.None;
            if (isLow)
            {
                type |= BinningType.Low;
            }
            if (isHigh)
            {
                type |= BinningType.High;
            }

            if (_binningMap.TryGetValue(type, out string suffix))
            {
                row.Comment = (row.Comment ?? "") + "_" + suffix;
            }


            row.LoLim = DataConvertor.ConvertUseLimit(pin.LowLimit, out string lowUnit, out string lowScale);
            if (repeat != "")
            {
                if (row.LoLim.Contains("="))
                {
                    row.LoLim = row.LoLim.Replace("=", "=(") + ")" + "*" + repeat.Replace("x", "");
                }
                else
                {
                    row.LoLim = HardIpService.CalcuteLimit(row.LoLim, repeat);
                }
            }
            row.HiLim = DataConvertor.ConvertUseLimit(pin.HighLimit, out string highUnit, out string highScale);
            if (repeat != "")
            {
                if (row.HiLim.Contains("="))
                {
                    row.HiLim = row.HiLim.Replace("=", "=(") + ")" + "*" + repeat.Replace("x", "");
                }
                else
                {
                    row.HiLim = HardIpService.CalcuteLimit(row.HiLim, repeat);
                }
            }
            row.Scale = lowScale == "" ? highScale : lowScale;
            row.Units = lowUnit == "" ? highUnit : lowUnit;
            row.FailAction = createFailAction(repeat);
            return row;
        }

        #region Create Test Row Column Methods

        protected internal virtual string CreateTestParameter(bool forShmoo = false)
        {
            if (!string.IsNullOrEmpty(Pat.TestName))
            {
                if (Pat.WirelessData.IsNeedPostBurn)
                {
                    return Pat.TestName + "_PostBurn" + "_" + ActualLabelVoltage;
                }

                return Pat.TestName + "_" + ActualLabelVoltage;
            }

            string patternName = Pat.Pattern.GetPatternName();
            string para = CommonGenerator.GenHardIpInsTestName(
                BlockName,
                Pat.ForceCondition.IsVtShmoo && forShmoo ? SubBlockName + "VTSHMOO" : SubBlockName,
                patternName,
                Pat.PatternIndexFlag,
                TimingAc,
                Pat.ForceVoltageFlag,
                InstNameSubStr,
                ActualLabelVoltage,
                NoPattern,
                Pat.WirelessData.IsNeedPostBurn,
                true,
                Pat.WirelessData.IsDoMeasure);

            if (Pat.MiscInfoDict.ContainsKey(HardIpConstData.ReTest))
            {
                para = para.Replace(LabelVoltage, HardIpConstData.PrefixReTest + ActualLabelVoltage);
            }

            return para;
        }

        protected string CreateShmooParameter()
        {
            string charName = "";
            HardipCharSetup shmoo = Pat.Shmoo;
            if (shmoo != null)
            {
                charName = shmoo.SetupName;
            }

            return CreateTestParameter(true) + " " + CommonGenerator.GetSubBlockNameWithoutMinus(charName);
        }

        protected internal string CreateShmooTName()
        {
            string tName = "";
            HardipCharSetup shmooName = Pat.Shmoo;
            if (shmooName != null)
            {
                if (shmooName.TestNameInFlow.StartsWith("HAC"))
                {
                    return tName;
                }

                tName = shmooName.TestNameInFlow;
                string[] charArr = tName.Split('_');
                if (!string.IsNullOrEmpty(LabelVoltage))
                {
                    charArr[2] = LabelVoltage.Substring(0, 1);
                    tName = string.Join("_", charArr);
                }
            }
            return tName;
        }

        public override string CreateIfConditionByTestPanDefine()
        {
            string condition = "";
            if (!string.IsNullOrEmpty(Pat.SiteFlag))
            {
                List<string> conditionVoltageList = Pat.SiteFlag.Split(new[] { ';', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                if (conditionVoltageList.Exists(x => x.StartsWith(LabelVoltage)))
                {
                    condition = conditionVoltageList.Find(x => x.StartsWith(LabelVoltage)).Split('@').Last();
                }
            }
            return condition;
        }

        public override string CreateTestFailActionByTestPlanDefine()
        {
            string flag = "";
            List<string> flagVoltageList = Pat.Failflag.Split(new[] { ';', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            if (flagVoltageList.Exists(x => x.StartsWith(LabelVoltage)))
            {
                flag = flagVoltageList.Find(x => x.StartsWith(LabelVoltage)).Split('@').Last();
            }
            else if (flagVoltageList.Count() == 1)
            {
                flag = flagVoltageList.First();
            }

            return string.IsNullOrEmpty(flag) ? CreateTestFailAction() : flag;
        }

        protected internal virtual string CreateTestFailAction()
        {
            if (Pat.IsIgnorePatBinOut() ||
                VbtFunctionLibShared.EfusePrewriteFunctionList.Exists(f => f.Equals(Pat.FunctionName, StringComparison.CurrentCultureIgnoreCase)) ||
                (HardIpConstData.RegOpcodeInPatt.IsMatch(Pat.Pattern.RealPatternName) && string.IsNullOrEmpty(Pat.Failflag)))
            {
                return "";
            }

            if (!string.IsNullOrEmpty(Pat.Failflag))
            {
                return CreateTestFailActionByTestPlanDefine();
            }

            if (VbtFunctionLibShared.EfuseReadFunctionList.Exists(f => f.Equals(Pat.FunctionName, StringComparison.CurrentCultureIgnoreCase)))
            {
                return CommonGenerator.GenHipEfuseReadTestFailAction();
            }

            string patternName = Pat.Pattern.GetPatternName();
            if (Pat.Pattern.IsMultiple())
            {
                patternName = patternName.Replace(PatternClass.MultipleConst, "");
                string failFlag = CommonGenerator.GenHardIpFlowFailAction(SheetName, BlockName,
                    SubBlockName.ToUpper().Replace("-MERGE", ""), patternName, TimingAc, InstNameSubStr,
                    LabelVoltage, Pat.MiscInfo, NoPattern);
                ErrorReportManager.AddError(
                    HardIpErrorType.W_BurstPatJudgeFlag_02,
                    Pat.SheetName,
                    Pat.RowNum,
                    1,
                    ["Test Fail Action", failFlag]
                );
                return failFlag;
            }
            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                return CommonGenerator.GenWirelessFlowFailFlag(BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), LabelVoltage);
            }

            return CommonGenerator.GenHardIpFlowFailAction(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), patternName, TimingAc, InstNameSubStr,
                LabelVoltage, Pat.MiscInfo, NoPattern);
        }

        protected virtual string CreateUseLimitFailAction(bool forShmoo)
        {
            if ((Pat.ForceCondition.IsVtShmoo && forShmoo) ||
                Pat.IsIgnorePatBinOut())
            {
                return "";
            }

            if (!string.IsNullOrEmpty(Pat.Failflag))
            {
                return CreateTestFailActionByTestPlanDefine();
            }

            string patternName = Pat.Pattern.GetPatternName();
            if (Pat.Pattern.IsMultiple())
            {
                patternName = patternName.Replace(PatternClass.MultipleConst, "");
                string failFlag = CommonGenerator.GenHardIpFlowFailAction(SheetName, BlockName,
                    SubBlockName.ToUpper().Replace("-MERGE", ""), patternName, TimingAc, InstNameSubStr,
                    LabelVoltage, Pat.MiscInfo, NoPattern);
                ErrorReportManager.AddError(
                    HardIpErrorType.W_BurstPatJudgeFlag_02,
                    Pat.SheetName, Pat.RowNum,
                    1,
                    ["UseLimit Fail Action", failFlag]
                );
                return failFlag;
            }
            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                return CommonGenerator.GenWirelessFlowFailFlag(BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), LabelVoltage);
            }

            return CommonGenerator.GenHardIpFlowFailAction(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), patternName, TimingAc, InstNameSubStr,
                LabelVoltage, Pat.MiscInfo, NoPattern);
        }

        private string CreatePreTestEnv()
        {
            string env = string.Empty;
            return string.IsNullOrEmpty(CreateTestJob()) ? HardIpConstData.EnvTtr : env;
        }

        private string CreatePreTestParemeter()
        {
            return CommonGenerator.GenHardIpInsTestName(BlockName, SubBlockName, Pat.Pattern.GetLastPayload(), Pat.PatternIndexFlag, TimingAc,
                 Pat.ForceVoltageFlag, InstNameSubStr, ActualLabelVoltage, NoPattern, Pat.WirelessData.IsNeedPostBurn, true, Pat.WirelessData.IsDoMeasure);
        }

        private string CreatePreTestFailAction()
        {
            return Pat.IsIgnorePatBinOut() ? "" : HardIpConstData.ReTestFlag;
        }

        protected virtual string CreateBinTableParameter()
        {
            if (VbtFunctionLibShared.EfuseReadFunctionList.Exists(f => f.Equals(Pat.FunctionName, StringComparison.CurrentCultureIgnoreCase)))
            {
                return CommonGenerator.GenHardIpEfuseReadBinParameter(SheetName);
            }

            return CommonGenerator.GenHardIpFlowBinParameter(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""));
        }

        #endregion

    }

}
