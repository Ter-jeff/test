using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using CommonReaderLib;

using LogLib.Utility;

using OfficeOpenXml;

using TestPlanLib.EVS;
using TestPlanLib.Utility;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public partial class BinCutInstanceSheetReader : MySheetReader<BinCutInstanceSheet>
    {
        protected const string ConHeaderFlowName = "Flow name";
        private const string ConHeaderSkip = "Skip";
        private const string ConHeaderInstance = "Instance";
        private const string ConHeaderOpcode = "Opcode";
        private const string ConHeaderEnbable = "Enable";
        private const string ConHeaderBinTableEnbable = @"^BinTable Enable\s*\(all sites\)";
        private const string ConHeaderPattern = @"^Pattern\d*$";
        private const string ConHeaderComment = "Comment";
        private const string ConHeaderPatSetOrange = "PatSetName(Orange)";
        private const string ConHeaderSubFlow = @"Sub\s*Flow";
        private const string ConHeaderEnableNew = @"^Enable\s*\(all\s*sites\)";
        private const string ConHeaderJobTestStage1 = @"Job\s*/\s*Testing\s*Stage\s*\(all\s*sites\)";
        private const string ConHeaderJobTestStage2 = @"\s*Testing\s*Stage\s*\(all\s*sites\)";
        private const string ConHeaderTestingStage = "Testing Stage";
        private const string ConHeaderSiteVar = @"SiteFlag\s*\(per\s*site\)";
        private const string ConHeaderSiteVar1 = "Condition";
        private const string ConHeaderFailFlag = "FailFlag";
        private const string ConHeaderPassFlag = "PassFlag";
        private const string ConHeaderBinOutStage = "Bin_out_stage";
        private const string ConHeaderVoltage = @"Voltage\s*Category";
        private const string ConHeaderHarvPinGrpEnable = "HarvPinGrp_Enable";
        private const string ConHeaderHarvestBinningFlag = "HarvestBinningFlag";
        private const string ConHeaderIsHarvesting = "IsHarvesting";
        private const string ConHeaderLevels = "Pin Levels";
        public const string ConHeaderTempMon = "TempMon";
        #region EVS
        private const string ConHeaderEvsVoltage = @"^EVS\s*Voltage\s*Category";
        private const string ConHeaderEvsRampingCount = "Multi_EVS_ramping_counts";
        private const string ConHeaderEvsStressTime = "EVS_Stress_time";
        private const string ConHeaderEvsCoolingTime = "EVS_Cooling_Time";
        private const string ConHeaderEvsPwrPin1 = "PowerPin1";
        private const string ConHeaderEvsPwrPin2 = "PowerPin2";
        private const string ConHeaderEvsParallelSetting = "Parallel Setting";
        private const string ConHeaderEvsTotalPwrLimit = "TotalPWRLimit";
        private const string ConHeaderEvsPatternVoltage = "Pattern Voltage";
        private const string ConHeaderEvsCondition = @"^(CP|FT)\d_EVS\d_";
        private const string ConHeaderType = "Type";
        private const string ConHeaderDvsCoolingPatIndex = "DvsCoolingPatIndex";
        private const string ConHeaderDvsCoolingPatSetLoopNumber = "DvsCoolingPatSetLoopNumber";
        #endregion
        private const string ConHeaderTimeSet = "Timeset";
        private const string ConHeaderShiftSpeed = "ShiftSpeed";
        private const string ConHeaderPatternPinGroup = @"Pattern\s*pin\s*groups\s*$";
        private const string ConHeaderPatternPinGroup1 = @"Test\s*Result";
        private const string ConHeaderPinGroupBinout = @"Pattern\s*pin\s*groups\s*bin\s*out\s*$";
        private const string ConHeaderMultiFstpEnable = "MultiFSTP_Enable";
        private const string ConHeaderUserFunction = "UserFunction";
        private const string ConHeaderFunctionName = "FunctionName";
        private const string ConHeaderSpecialSetting = "Special Setting";
        private const string ConHeaderOverlay = "Overlay";
        private const string ConHeaderChar = "Char";
        public const string ConHeaderPerfMode = "PerfMode";
        public const string ConHeaderPmOrder = "PM Order";
        public const string ConHeaderCallFlow = "Call Flow";
        public const string ConHeaderBurst = "Burst";

        [GeneratedRegex(@"(?<key>[a-zA-Z]+\d*)(?<codes>\[.+\])", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"\n*\s*\t*")]
        private static partial Regex MyRegex8();
        [GeneratedRegex(ConHeaderBinTableEnbable, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex10();
        [GeneratedRegex(ConHeaderPattern, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex11();
        [GeneratedRegex(ConHeaderSubFlow, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex12();
        [GeneratedRegex(ConHeaderEnableNew, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex13();
        [GeneratedRegex(ConHeaderJobTestStage1, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex14();
        [GeneratedRegex(ConHeaderJobTestStage2, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex15();
        [GeneratedRegex(ConHeaderSiteVar, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex16();
        [GeneratedRegex(ConHeaderSiteVar1, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex17();
        [GeneratedRegex(ConHeaderEvsVoltage, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex18();
        [GeneratedRegex(ConHeaderEvsRampingCount, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex19();
        [GeneratedRegex(ConHeaderEvsStressTime, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex20();
        [GeneratedRegex(ConHeaderEvsCoolingTime, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex21();
        [GeneratedRegex(ConHeaderEvsPwrPin1, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex22();
        [GeneratedRegex(ConHeaderEvsPwrPin2, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex23();
        [GeneratedRegex(ConHeaderEvsParallelSetting, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex24();
        [GeneratedRegex(ConHeaderEvsTotalPwrLimit, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex25();
        [GeneratedRegex(ConHeaderVoltage, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex26();
        [GeneratedRegex(ConHeaderPatternPinGroup, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex27();
        [GeneratedRegex(ConHeaderPatternPinGroup1, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex28();
        [GeneratedRegex(ConHeaderPinGroupBinout, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex29();
        [GeneratedRegex(ConHeaderMultiFstpEnable, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex30();
        [GeneratedRegex(ConHeaderUserFunction, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex31();
        [GeneratedRegex(ConHeaderPerfMode, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex32();
        [GeneratedRegex(ConHeaderPmOrder, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex33();
        [GeneratedRegex(ConHeaderCallFlow, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex34();
        [GeneratedRegex(ConHeaderSpecialSetting, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex35();
        [GeneratedRegex(ConHeaderOverlay, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex36();
        [GeneratedRegex(ConHeaderChar, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex37();
        [GeneratedRegex(ConHeaderFunctionName, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex38();
        [GeneratedRegex(ConHeaderBurst, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex39();
        [GeneratedRegex(ConHeaderEvsCondition, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex40();
        [GeneratedRegex(ConHeaderType, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex41();
        [GeneratedRegex(ConHeaderDvsCoolingPatIndex, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex42();
        [GeneratedRegex(ConHeaderDvsCoolingPatSetLoopNumber, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex43();

        private readonly BasicColumnIndices _basic = new();
        private readonly EvsColumnIndices _evs = new();
        private readonly MiscColumnIndices _misc = new();
        public EnumInstanceSheetType Type = EnumInstanceSheetType.Bincut;

        public BinCutInstanceSheetReader()
        {
        }

        public override BinCutInstanceSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            GetDimensions();

            GetFirstHeaderPosition(ConHeaderFlowName);

            GetHeaders();

            if (_basic.PatternStartColNumber == 0)
            {
                ErrorMessageBox.Show("Can't find any pattern titles in " + sheetName + " sheet", "");
                throw new Exception();
            }

            int totalHeaderCol = _basic.PatternCounlumCnt + _basic.PatternStartColNumber - 1;
            int maxColNum = totalHeaderCol;
            for (int row = StartRow + 1; row <= EndRow; row++)
            {
                for (int col = totalHeaderCol; col <= EndCol; col++)
                {
                    if (!string.IsNullOrEmpty(ExcelWorksheet.GetCellValue(row, col)) && col > maxColNum)
                    {
                        maxColNum = col;
                    }
                }
            }

            // find the max pattern count
            if (maxColNum > totalHeaderCol)
            {
                string error = "The number of pattern columns (" + maxColNum + ") is different from the pattern title (" + totalHeaderCol + ") in sheet (" + sheetName + ")";
                ErrorMessageBox.Show(error, "");
                throw new Exception(error);
            }

            BinCutInstanceSheet sheet = ReadSheet(sheetName);

            CheckTtr(sheet);

            ClassificationPattern(ref sheet);

            sheet = SplitByCode(sheet);

            CheckPattern(ref sheet, BinCutErrorType.E_Pattern_05, BinCutErrorType.E_Pattern_06, BinCutErrorType.W_Pattern_06);

            return sheet;
        }

        protected static BinCutInstanceSheet SplitByCode(BinCutInstanceSheet binCutInstanceSheet)
        {
            var rows = new List<BinCutInstanceRow>();
            foreach (BinCutInstanceRow row in binCutInstanceSheet.Rows)
            {
                if (MyRegex().IsMatch(row.Instance))
                {
                    Match regexResult = MyRegex().Match(row.Instance);
                    string codes = regexResult.Groups["codes"].ToString();
                    string splitKey = regexResult.Groups["key"].ToString();
                    List<string> iterates = BinCutInstanceSheetReaderHelpers.GetCodes(codes);
                    for (int index = 0; index < iterates.Count; index++)
                    {
                        string iterate = iterates[index];
                        BinCutInstanceRow newRow = row.Copy();
                        newRow.FlowName = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.FlowName, binCutInstanceSheet.FlowNameColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.Instance = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.Instance, binCutInstanceSheet.InstanceColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.Opcode = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.Opcode, binCutInstanceSheet.OpcodeColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.EnableAndDevice = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.EnableAndDevice, binCutInstanceSheet.EnableColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.BintableEnableWd = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.BintableEnableWd, binCutInstanceSheet.BinTableEnableColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.SubFlow = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.SubFlow, binCutInstanceSheet.SubFlowColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.EnableFlow = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.EnableFlow, binCutInstanceSheet.EnableNewColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.SiteVar = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.SiteVar, binCutInstanceSheet.SiteVarColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.FailFlag = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.FailFlag, binCutInstanceSheet.FailFlagColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.PassFlag = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.PassFlag, binCutInstanceSheet.PassFlagColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.DCcategory = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.DCcategory, binCutInstanceSheet.VoltageColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.PatternPinGroup = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.PatternPinGroup, binCutInstanceSheet.PinGroupBinoutColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.PinGroupBinoutFlag = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.PinGroupBinoutFlag, binCutInstanceSheet.PinGroupBinoutColNumber, binCutInstanceSheet, row, index, iterates.Count);
                        newRow.CallFlow = BinCutInstanceSheetReaderHelpers.SplitByCode(newRow.CallFlow, binCutInstanceSheet.CallFlowColNumber, binCutInstanceSheet, row, index, iterates.Count);

                        newRow.SplitKey = splitKey + iterate;
                        newRow.PatternList = newRow.PatternList.ConvertAll(x => x.Replace(codes, iterate));
                        newRow.InitList = newRow.InitList.ConvertAll(x => x.Replace(codes, iterate));
                        newRow.PayloadList = newRow.PayloadList.ConvertAll(x => x.Replace(codes, iterate));
                        newRow.PatWithIndex = newRow.PatWithIndex.ToDictionary(item => item.Key,
                            item => BinCutInstanceSheetReaderHelpers.SplitByCode(item.Value, binCutInstanceSheet.PatternStartColNumber + item.Key, binCutInstanceSheet, row, index, iterates.Count));
                        rows.Add(newRow);
                    }
                }
                else
                {
                    rows.Add(row);
                }
            }
            binCutInstanceSheet.Rows = rows;

            List<BinCutInstanceRow> rowList = BinCutInstanceSheetReaderHelpers.MergebyInstanceName(binCutInstanceSheet.Rows);

            rowList = [.. rowList.Where(x => !string.IsNullOrEmpty(x.FlowName))];

            binCutInstanceSheet.Rows = rowList;

            return binCutInstanceSheet;
        }

        protected virtual void CheckPattern(ref BinCutInstanceSheet binCutInstanceSheet, ErrorCode emptyStringType, ErrorCode firstPatternEmptyType, ErrorCode patternEmptyType)
        {
            BinCutInstanceSheetReaderHelpers.CheckPatternBase(ref binCutInstanceSheet, emptyStringType, firstPatternEmptyType, patternEmptyType);

            foreach (BinCutInstanceRow row in binCutInstanceSheet.Rows)
            {
                row.PatternList = [.. row.PatternList.Where(x => !string.IsNullOrEmpty(x))];
                row.PayloadList = [.. row.PayloadList.Where(x => !string.IsNullOrEmpty(x))];
                row.InitList = [.. row.InitList.Where(x => !string.IsNullOrEmpty(x))];
            }
        }

        private void CheckTtr(BinCutInstanceSheet binCutInstanceSheet)
        {
            var groups = new List<List<BinCutInstanceRow>>();
            for (int i = 0; i < binCutInstanceSheet.Rows.Count; i++)
            {
                BinCutInstanceRow row1 = binCutInstanceSheet.Rows[i];
                var rows = new List<BinCutInstanceRow> { row1 };
                for (int j = i + 1; j < binCutInstanceSheet.Rows.Count; j++)
                {
                    BinCutInstanceRow row2 = binCutInstanceSheet.Rows[j];
                    if (!row1.Instance.EqualsIgnoreCase(row2.Instance))
                    {
                        break;
                    }
                    if (!row1.SubFlow.EqualsIgnoreCase(row2.SubFlow))
                    {
                        break;
                    }

                    rows.Add(row2);
                    i = j;
                }
                groups.Add(rows);
            }

            foreach (List<BinCutInstanceRow> group in groups)
            {
                var flowNames = group.Select(x => x.FlowName).Distinct().ToList();
                if (flowNames.Count > 1)
                {
                    BinCutInstanceRow row = group[0];
                    binCutInstanceSheet.AddError(BinCutErrorType.W_FlowName_01, binCutInstanceSheet.SheetName, row.RowNum, _basic.FlowNameColNumber == -1 ? _basic.SubFlowColNumber : _basic.FlowNameColNumber, $"Can not use multi flow name {string.Join(",", flowNames)} in one TTR instance !!!", [string.Join(",", flowNames)]);

                    string lastFlowName = group[0].FlowName;
                    foreach (BinCutInstanceRow item in group)
                    {
                        if (!item.FlowName.EqualsIgnoreCase(lastFlowName))
                        {
                            if (string.IsNullOrEmpty(item.PatternList[0]))
                            {
                                binCutInstanceSheet.AddError(BinCutErrorType.E_FlowName_02, binCutInstanceSheet.SheetName, item.RowNum, _basic.PatternStartColNumber, "The first init can not be empty in one TTR instance !!!");
                            }
                        }
                        lastFlowName = item.FlowName;
                    }
                }

                if (group.Count > 1)
                {
                    if (group.Select(x => x.EnableAndDevice).Distinct().Count() > 1)
                    {
                        IEnumerable<string> items = group.Select(x => x.EnableAndDevice).Where(x => !string.IsNullOrEmpty(x)).Distinct();
                        binCutInstanceSheet.AddError(BinCutErrorType.E_TTR_01, binCutInstanceSheet.SheetName, group[0].RowNum, _basic.EnableColNumber, $"Can not use multi enable {string.Join(",", items)} in one TTR instance !!!", [string.Join(",", items)]);
                    }
                    if (group.Select(x => x.EnableFlow).Distinct().Count() > 1)
                    {
                        IEnumerable<string> items = group.Select(x => x.EnableFlow).Where(x => !string.IsNullOrEmpty(x)).Distinct();
                        binCutInstanceSheet.AddError(BinCutErrorType.W_TTR_01, binCutInstanceSheet.SheetName, group[0].RowNum, _basic.EnableNewColNumber, $"Can not use multi enable {string.Join(",", items)} in one TTR instance !!!", [string.Join(",", items)]);
                    }
                    if (group.Select(x => x.JobTestStage).Distinct().Count() > 1)
                    {
                        IEnumerable<string> items = group.Select(x => x.JobTestStage).Where(x => !string.IsNullOrEmpty(x)).Distinct();
                        binCutInstanceSheet.AddError(BinCutErrorType.W_TTR_02, binCutInstanceSheet.SheetName, group[0].RowNum, _basic.JobTestStageColNumber, $"Can not use multi JobTestStage {string.Join(",", items)} in one TTR instance !!!", [string.Join(",", items)]);
                    }
                    if (group.Select(x => x.SiteVar).Distinct().Count() > 1)
                    {
                        IEnumerable<string> items = group.Select(x => x.SiteVar).Where(x => !string.IsNullOrEmpty(x)).Distinct();
                        binCutInstanceSheet.AddError(BinCutErrorType.W_TTR_03, binCutInstanceSheet.SheetName, group[0].RowNum, _basic.SiteVarColNumber, $"Can not use multi SiteVar {string.Join(",", items)} in one TTR instance !!!", [string.Join(",", items)]);
                    }
                    if (group.Select(x => x.ShiftSpeed).Distinct().Count() > 1)
                    {
                        IEnumerable<string> items = group.Select(x => x.ShiftSpeed).Where(x => !string.IsNullOrEmpty(x)).Distinct();
                        binCutInstanceSheet.AddError(BinCutErrorType.W_TTR_04, binCutInstanceSheet.SheetName, group[0].RowNum, _basic.ShiftSpeedColNumber, $"Can not use multi ShiftSpeed {string.Join(",", items)} in one TTR instance !!!", [string.Join(",", items)]);
                    }
                    if (group.Select(x => x.DCcategory).Distinct().Count() > 1)
                    {
                        IEnumerable<string> items = group.Select(x => x.DCcategory).Where(x => !string.IsNullOrEmpty(x)).Distinct();
                        binCutInstanceSheet.AddError(BinCutErrorType.W_TTR_05, binCutInstanceSheet.SheetName, group[0].RowNum, _basic.ShiftSpeedColNumber, $"Can not use multi Voltage Category {string.Join(",", items)} in one TTR instance !!!", [string.Join(",", items)]);
                    }
                    if (group.Select(x => x.TimeSet).Distinct().Count() > 1)
                    {
                        IEnumerable<string> items = group.Select(x => x.TimeSet).Where(x => !string.IsNullOrEmpty(x)).Distinct();
                        binCutInstanceSheet.AddError(BinCutErrorType.W_TTR_06, binCutInstanceSheet.SheetName, group[0].RowNum, _basic.ShiftSpeedColNumber, $"Can not use multi timeSet {string.Join(",", items)} in one TTR instance !!!", [string.Join(",", items)]);
                    }
                    if (group.Select(x => x.FailFlag).Distinct().Count() > 1)
                    {
                        IEnumerable<string> items = group.Select(x => x.FailFlag).Where(x => !string.IsNullOrEmpty(x)).Distinct();
                        binCutInstanceSheet.AddError(BinCutErrorType.W_TTR_07, binCutInstanceSheet.SheetName, group[0].RowNum, _basic.ShiftSpeedColNumber, $"Can not use multi FailFlag {string.Join(",", items)} in one TTR instance !!!", [string.Join(",", items)]);
                    }
                    if (group.Select(x => x.PatSetNameOrange).Distinct().Count() > 1)
                    {
                        IEnumerable<string> items = group.Select(x => x.PatSetNameOrange).Where(x => !string.IsNullOrEmpty(x)).Distinct();
                        binCutInstanceSheet.AddError(BinCutErrorType.W_TTR_08, binCutInstanceSheet.SheetName, group[0].RowNum, _basic.ShiftSpeedColNumber, $"Can not use multi PatSetName(Orange) {string.Join(",", items)} in one TTR instance !!!", [string.Join(",", items)]);
                    }
                }
            }
        }

        protected void ClassificationPattern(ref BinCutInstanceSheet binCutInstanceSheet)
        {
            int lastInitPatternCnt = 0;
            bool existInit = false;

            foreach (BinCutInstanceRow row in binCutInstanceSheet.Rows)
            {
                List<string> patListOrg = row.PatternList;

                BinCutInstanceSheetReaderHelpers.FindTheLastInitPattern(ref lastInitPatternCnt, ref existInit, patListOrg);

                BinCutInstanceSheetReaderHelpers.FindTheFirstPayload(patListOrg, out int firstPayloadCnt, out bool existPayload);

                row.PatternList = [];
                int patternIdx = 0;
                patternIdx = BinCutInstanceSheetReaderHelpers.ClassifyPattern(Type, _basic.PatternStartColNumber, binCutInstanceSheet, lastInitPatternCnt, existInit, row, patListOrg, firstPayloadCnt, existPayload, patternIdx);
            }
        }

        protected BinCutInstanceSheet ReadSheet(string sheetName)
        {
            var sheet = new BinCutInstanceSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                if (_basic.SkipColNumber != -1 && ExcelWorksheet.GetCellValue(i, _basic.SkipColNumber).Trim().EqualsIgnoreCase("Yes"))
                {
                    continue;
                }

                string flowName = _basic.FlowNameColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.FlowNameColNumber).Trim();
                string subFlow = _basic.SubFlowColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.SubFlowColNumber).Trim();
                if (flowName?.Length == 0 && subFlow?.Length == 0)
                {
                    continue;
                }
                if (_basic.PatternStartColNumber != 0 || sheetName == "Instance_EVS_PowerReset")
                {
                    BinCutInstanceRow row = SetRow(sheetName, i);

                    if (row.PatternList.Count != 0 ||
                        !string.IsNullOrEmpty(row.Evs.EvsCategory) ||
                        !string.IsNullOrEmpty(row.Evs.EvsParallelSetting) ||
                        row.FlowName.EqualsIgnoreCase("Flow_EVS") ||
                        row.SubFlow.EqualsIgnoreCase("Flow_EVS") ||
                        row.SubFlow.EqualsIgnoreCase("Flow_EVS_PowerReset") ||
                        !string.IsNullOrEmpty(row.SpecialSetting) || Type.Equals(EnumInstanceSheetType.Efuse)
                        || sheetName.EqualsIgnoreCase("TurboModeCheck"))
                    {
                        sheet.Rows.Add(row);
                    }
                }
            }

            MapIndicesToSheet(sheet);

            return sheet;
        }

        private BinCutInstanceRow SetRow(string sheetName, int i)
        {
            var row = new BinCutInstanceRow
            {
                RowNum = i,
                SheetName = sheetName,
                FlowName = _basic.FlowNameColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.FlowNameColNumber).Trim(),
                SubFlow = _basic.SubFlowColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.SubFlowColNumber).Trim(),
                Instance = ExcelWorksheet.GetCellValue(i, _basic.InstanceColNumber).Trim(),
                Opcode = ExcelWorksheet.GetCellValue(i, _basic.OpcodeColNumber).Trim(),
                EnableAndDevice = _basic.EnableColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.EnableColNumber).Trim(),
                BintableEnableWd = _basic.BinTableEnableColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.BinTableEnableColNumber).Trim(),
                PatSetNameOrange = _basic.PatSetNameOrangeColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.PatSetNameOrangeColNumber).Trim(),
                EnableFlow = _basic.EnableNewColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.EnableNewColNumber).Trim(),
                JobTestStage = _basic.JobTestStageColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.JobTestStageColNumber).Trim(),
                TimeSet = _basic.TimeSetColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.TimeSetColNumber).Trim(),
                ShiftSpeed = _basic.ShiftSpeedColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.ShiftSpeedColNumber).Trim(),
                SiteVar = _basic.SiteVarColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.SiteVarColNumber).Trim(),
                FailFlag = _basic.FailFlagColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.FailFlagColNumber).Trim(),
                PassFlag = _basic.PassFlagColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.PassFlagColNumber).Trim(),
                BinOutStage = _basic.BinOutStageColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.BinOutStageColNumber).Trim(),
                HarvPinGrpEnable = _basic.HarvPinGrpEnableColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.HarvPinGrpEnableColNumber).Trim(),
                HarvestBinningFlag = _basic.HarvestBinningFlagColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.HarvestBinningFlagColNumber).Trim(),
                IsHarvesting = _basic.IsHarvestingColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.IsHarvestingColNumber).Trim(),
                DCcategory = _basic.VoltageColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.VoltageColNumber).Trim(),
                Levels = _basic.LevelsColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.LevelsColNumber).Trim(),
                TempMon = _basic.TempMonColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _basic.TempMonColNumber).Trim(),
            };

            SetRowEvs(i, row);

            SerRowOthers(i, row);

            return row;
        }

        private void SetRowEvs(int i, BinCutInstanceRow binCutInstanceRow)
        {
            binCutInstanceRow.Evs.EvsCategory = _evs.EvsVoltageColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _evs.EvsVoltageColNumber).Trim();
            binCutInstanceRow.Evs.EvsRampingCount = _evs.EvsRampingColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _evs.EvsRampingColNumber).Trim();
            binCutInstanceRow.Evs.EvsStressTime = _evs.EvsStressTimeColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _evs.EvsStressTimeColNumber).Trim();
            binCutInstanceRow.PatternPinGroup = _misc.PatternPinGroupColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _misc.PatternPinGroupColNumber).Trim();
            binCutInstanceRow.PinGroupBinoutFlag = _misc.PinGroupBinoutColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _misc.PinGroupBinoutColNumber).Trim();
            binCutInstanceRow.MultiFstpEnable = _misc.MultiFstpEnableColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _misc.MultiFstpEnableColNumber).Trim();
            binCutInstanceRow.UserFunction = _misc.UserFunctionColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _misc.UserFunctionColNumber).Trim();
            binCutInstanceRow.Evs.EvsCoolingTime = _evs.EvsCoolingTimeColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _evs.EvsCoolingTimeColNumber).Trim();
            binCutInstanceRow.Evs.EvsPwrPin1 = _evs.EvsPwrPin1ColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _evs.EvsPwrPin1ColNumber).Trim();
            binCutInstanceRow.Evs.EvsPwrPin2 = _evs.EvsPwrPin2ColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _evs.EvsPwrPin2ColNumber).Trim();
            binCutInstanceRow.Evs.EvsParallelSetting = _evs.EvsParallelSettingColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _evs.EvsParallelSettingColNumber).Trim();
            binCutInstanceRow.Evs.EvsTotalPwrLimit = _evs.EvsTotalPwrLimitColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _evs.EvsTotalPwrLimitColNumber).Trim();
            binCutInstanceRow.SpecialSetting = _misc.SpecialSettingColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _misc.SpecialSettingColNumber).Trim();
            binCutInstanceRow.Overlay = _misc.OverlayColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _misc.OverlayColNumber).Trim();
            binCutInstanceRow.Char = _misc.CharColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _misc.CharColNumber).Trim();
            binCutInstanceRow.FunctionName = _misc.FunctionNameColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _misc.FunctionNameColNumber).Trim();
            binCutInstanceRow.Burst = _misc.BurstColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _misc.BurstColNumber).Trim();
            binCutInstanceRow.Evs.EvsType = _evs.TypeColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _evs.TypeColNumber).Trim();
            binCutInstanceRow.Evs.DvsCoolingPatIndex = _evs.DvsCoolingPatIndexColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _evs.DvsCoolingPatIndexColNumber).Trim();
            binCutInstanceRow.Evs.DvsCoolingPatSetLoopNumber = _evs.DvsCoolingPatSetLoopNumberColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _evs.DvsCoolingPatSetLoopNumberColNumber).Trim();
        }

        private void SerRowOthers(int i, BinCutInstanceRow binCutInstanceRow)
        {
            if (!string.IsNullOrEmpty(binCutInstanceRow.Evs.EvsPwrPin1))
            {
                binCutInstanceRow.Evs.EvsPwrPin1 = MyRegex8().Replace(binCutInstanceRow.Evs.EvsPwrPin1, "");
            }

            if (!string.IsNullOrEmpty(binCutInstanceRow.Evs.EvsPwrPin2))
            {
                binCutInstanceRow.Evs.EvsPwrPin2 = MyRegex8().Replace(binCutInstanceRow.Evs.EvsPwrPin2, "");
            }

            binCutInstanceRow.PerfMode = _misc.PerfModeColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _misc.PerfModeColNumber).Trim();
            if (_misc.PmOrderColNumber != -1)
            {
                if (float.TryParse(ExcelWorksheet.GetCellValue(i, _misc.PmOrderColNumber).Trim(), out float value))
                {
                    binCutInstanceRow.PmOrder = value;
                }
            }

            binCutInstanceRow.CallFlow = _misc.CallFlowColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _misc.CallFlowColNumber).Trim();

            if (!string.IsNullOrEmpty(binCutInstanceRow.PatternPinGroup))
            {
                binCutInstanceRow.PatternPinGroup = MyRegex8().Replace(binCutInstanceRow.PatternPinGroup, "");
            }

            if (string.IsNullOrEmpty(binCutInstanceRow.FlowName))
            {
                binCutInstanceRow.FlowName = binCutInstanceRow.SubFlow;
            }

            foreach (EvsConditionColIdx conditionIdxs in _evs.ConditionColIdxDic)
            {
                binCutInstanceRow.Evs.EvsConditions.Add(BuildEvsCondition(i, conditionIdxs));
            }
            int lastPatIdx = 0;

            // get final pattern index
            for (int col = 0; col < _basic.PatternCounlumCnt; col++)
            {
                string pattern = ExcelWorksheet.GetCellValue(i, _basic.PatternStartColNumber + col).Trim();
                if (string.IsNullOrEmpty(pattern) || pattern == "NA")
                {
                    continue;
                }

                lastPatIdx = col + 1;
            }

            for (int col = 0; col < lastPatIdx; col++)
            {
                string pattern = ExcelWorksheet.GetCellValue(i, _basic.PatternStartColNumber + col).Trim();

                if (pattern == "NA")
                {
                    continue;
                }

                binCutInstanceRow.PatternList.Add(pattern);
            }
        }

        private EvsCondition BuildEvsCondition(int i, EvsConditionColIdx evsConditionColIdx)
        {
            var condition = new EvsCondition(evsConditionColIdx.JobStage, evsConditionColIdx.EvsNum)
            {
                NumConds = evsConditionColIdx.NumCondsIdx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.NumCondsIdx).Trim(),
                Pulses = evsConditionColIdx.PulsesIdx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.PulsesIdx).Trim(),
                Time = evsConditionColIdx.TimeIdx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.TimeIdx).Trim(),
                Voltage1 = evsConditionColIdx.Voltage1Idx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.Voltage1Idx).Trim(),
                Voltage2 = evsConditionColIdx.Voltage2Idx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.Voltage2Idx).Trim(),
                Cooling = evsConditionColIdx.CoolingIdx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.CoolingIdx).Trim(),
                CoolingAfter = evsConditionColIdx.CoolingAfterIdx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.CoolingAfterIdx).Trim(),
                TotalPwr = evsConditionColIdx.TotalPwrIdx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.TotalPwrIdx).Trim(),
                AlarmFlag = evsConditionColIdx.AlarmFlagIdx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.AlarmFlagIdx).Trim(),
                RisingDelayTimeSec = evsConditionColIdx.RisingDelayTimeSecIdx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.RisingDelayTimeSecIdx).Trim()
            };

            condition.UserFunction["Ramp"] = evsConditionColIdx.RampIdx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.RampIdx).Trim();
            condition.UserFunction["Vtrig"] = evsConditionColIdx.VtrigIdx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.VtrigIdx).Trim();
            condition.UserFunction["IV"] = evsConditionColIdx.IVdx == -1 ? "" : ExcelWorksheet.GetCellValue(i, evsConditionColIdx.IVdx).Trim();

            return condition;
        }

        private void MapIndicesToSheet(BinCutInstanceSheet binCutInstanceSheet)
        {
            binCutInstanceSheet.FlowNameColNumber = _basic.FlowNameColNumber;
            binCutInstanceSheet.InstanceColNumber = _basic.InstanceColNumber;
            binCutInstanceSheet.OpcodeColNumber = _basic.OpcodeColNumber;
            binCutInstanceSheet.EnableColNumber = _basic.EnableColNumber;
            binCutInstanceSheet.PatternStartColNumber = _basic.PatternStartColNumber;
            binCutInstanceSheet.CommnetColNumber = _basic.CommnetColNumber;
            binCutInstanceSheet.PatSetNameOrangeColNumber = _basic.PatSetNameOrangeColNumber;
            binCutInstanceSheet.SubFlowColNumber = _basic.SubFlowColNumber;
            binCutInstanceSheet.EnableNewColNumber = _basic.EnableNewColNumber;
            binCutInstanceSheet.JobTestStageColNumber = _basic.JobTestStageColNumber;
            binCutInstanceSheet.SiteVarColNumber = _basic.SiteVarColNumber;
            binCutInstanceSheet.FailFlagColNumber = _basic.FailFlagColNumber;
            binCutInstanceSheet.PassFlagColNumber = _basic.PassFlagColNumber;
            binCutInstanceSheet.BinOutStageColNumber = _basic.BinOutStageColNumber;
            binCutInstanceSheet.VoltageColNumber = _basic.VoltageColNumber;
            binCutInstanceSheet.LevelsColNumber = _basic.LevelsColNumber;
            binCutInstanceSheet.EvsConditionColIdxDic = _evs.ConditionColIdxDic;
            binCutInstanceSheet.EvsVoltageColNumber = _evs.EvsVoltageColNumber;
            binCutInstanceSheet.EvsRampingColNumber = _evs.EvsRampingColNumber;
            binCutInstanceSheet.EvsStressTimeColNumber = _evs.EvsStressTimeColNumber;
            binCutInstanceSheet.EvsCoolingTimeColNumber = _evs.EvsCoolingTimeColNumber;
            binCutInstanceSheet.EvsPwrPin1ColNumber = _evs.EvsPwrPin1ColNumber;
            binCutInstanceSheet.EvsPwrPin2ColNumber = _evs.EvsPwrPin2ColNumber;
            binCutInstanceSheet.EvsParallelSetting = _evs.EvsParallelSettingColNumber;
            binCutInstanceSheet.TimeSetColNumber = _basic.TimeSetColNumber;
            binCutInstanceSheet.ShiftSpeedColNumber = _basic.ShiftSpeedColNumber;
            binCutInstanceSheet.PatternPinGroupColNumber = _misc.PatternPinGroupColNumber;
            binCutInstanceSheet.PinGroupBinoutColNumber = _misc.PinGroupBinoutColNumber;
            binCutInstanceSheet.MultiFstpEnableColNumber = _misc.MultiFstpEnableColNumber;
            binCutInstanceSheet.UserFunctionColNumber = _misc.UserFunctionColNumber;
            binCutInstanceSheet.PerfModeColNumber = _misc.PerfModeColNumber;
            binCutInstanceSheet.PmOrderColNumber = _misc.PmOrderColNumber;
            binCutInstanceSheet.CallFlowColNumber = _misc.CallFlowColNumber;
            binCutInstanceSheet.SpecialSettingNumber = _misc.SpecialSettingColNumber;
            binCutInstanceSheet.OverlayNumber = _misc.OverlayColNumber;
            binCutInstanceSheet.CharNumber = _misc.CharColNumber;
            binCutInstanceSheet.FunctionNameColNumber = _misc.FunctionNameColNumber;
            binCutInstanceSheet.BurstColNumber = _misc.BurstColNumber;
            binCutInstanceSheet.TypeColNumber = _evs.TypeColNumber;
            binCutInstanceSheet.DvsCoolingPatIndexColNumber = _evs.DvsCoolingPatIndexColNumber;
            binCutInstanceSheet.DvsCoolingPatSetLoopNumberColNumber = _evs.DvsCoolingPatSetLoopNumberColNumber;
            binCutInstanceSheet.TempMonColNumber = _basic.TempMonColNumber;
        }

        protected void GetHeaders()
        {
            for (int i = 1; i <= EndCol; i++)
            {
                string text = ExcelWorksheet.GetCellValue(StartRow, i).ToUpper().Trim();

                GetIndexBasic(i, text);

                if (text.EqualsIgnoreCase(ConHeaderFailFlag))
                {
                    _basic.FailFlagColNumber = i;
                }
                else if (text.EqualsIgnoreCase(ConHeaderPassFlag))
                {
                    _basic.PassFlagColNumber = i;
                }
                else if (text.EqualsIgnoreCase(ConHeaderBinOutStage))
                {
                    _basic.BinOutStageColNumber = i;
                }
                else if (text.EqualsIgnoreCase(ConHeaderHarvPinGrpEnable))
                {
                    _basic.HarvPinGrpEnableColNumber = i;
                }
                else if (text.EqualsIgnoreCase(ConHeaderHarvestBinningFlag))
                {
                    _basic.HarvestBinningFlagColNumber = i;
                }
                else if (text.EqualsIgnoreCase(ConHeaderIsHarvesting))
                {
                    _basic.IsHarvestingColNumber = i;
                }
                else if (text.EqualsIgnoreCase(ConHeaderLevels))
                {
                    _basic.LevelsColNumber = i;
                }
                else if (text.EqualsIgnoreCase(ConHeaderTempMon))
                {
                    _basic.TempMonColNumber = i;
                }

                GetIndexEvs(i, text);

                if (MyRegex26().IsMatch(text))
                {
                    _basic.VoltageColNumber = i;
                }
                else if (MyRegex27().IsMatch(text) || MyRegex28().IsMatch(text))
                {
                    _misc.PatternPinGroupColNumber = i;
                }
                else if (MyRegex29().IsMatch(text))
                {
                    _misc.PinGroupBinoutColNumber = i;
                }
                else if (MyRegex30().IsMatch(text))
                {
                    _misc.MultiFstpEnableColNumber = i;
                }
                else if (MyRegex31().IsMatch(text))
                {
                    _misc.UserFunctionColNumber = i;
                }
                else if (MyRegex32().IsMatch(text))
                {
                    _misc.PerfModeColNumber = i;
                }
                else if (MyRegex33().IsMatch(text))
                {
                    _misc.PmOrderColNumber = i;
                }
                else if (MyRegex34().IsMatch(text))
                {
                    _misc.CallFlowColNumber = i;
                }
                else if (MyRegex35().IsMatch(text))
                {
                    _misc.SpecialSettingColNumber = i;
                }
                else if (MyRegex36().IsMatch(text))
                {
                    _misc.OverlayColNumber = i;
                }
                else if (MyRegex37().IsMatch(text))
                {
                    _misc.CharColNumber = i;
                }
                #region Efuse
                else if (MyRegex38().IsMatch(text))
                {
                    _misc.FunctionNameColNumber = i;
                }
                #endregion
                #region Burst
                else if (MyRegex39().IsMatch(text))
                {
                    _misc.BurstColNumber = i;
                }
                #endregion
            }
        }

        private void GetIndexEvs(int i, string text)
        {
            #region EVS
            if (MyRegex18().IsMatch(text))
            {
                _evs.EvsVoltageColNumber = i;
            }
            else if (MyRegex19().IsMatch(text))
            {
                _evs.EvsRampingColNumber = i;
            }
            else if (MyRegex20().IsMatch(text))
            {
                _evs.EvsStressTimeColNumber = i;
            }
            else if (MyRegex21().IsMatch(text))
            {
                _evs.EvsCoolingTimeColNumber = i;
            }
            else if (MyRegex22().IsMatch(text))
            {
                _evs.EvsPwrPin1ColNumber = i;
            }
            else if (MyRegex23().IsMatch(text))
            {
                _evs.EvsPwrPin2ColNumber = i;
            }
            else if (MyRegex24().IsMatch(text))
            {
                _evs.EvsParallelSettingColNumber = i;
            }
            else if (MyRegex25().IsMatch(text))
            {
                _evs.EvsTotalPwrLimitColNumber = i;
            }
            else if (text.EqualsIgnoreCase(ConHeaderEvsPatternVoltage))
            {
                _basic.VoltageColNumber = i;
            }
            if (MyRegex41().IsMatch(text))
            {
                _evs.TypeColNumber = i;
            }
            if (MyRegex42().IsMatch(text))
            {
                _evs.DvsCoolingPatIndexColNumber = i;
            }
            if (MyRegex43().IsMatch(text))
            {
                _evs.DvsCoolingPatSetLoopNumberColNumber = i;
            }
            if (MyRegex40().IsMatch(text))
            {
                string jobStage = text.Split('_')[0];
                string evsNum = text.Split('_')[1];
                if (!_evs.ConditionColIdxDic.Any(x => x.JobStage == jobStage && x.EvsNum == evsNum))
                {
                    _evs.ConditionColIdxDic.Add(new EvsConditionColIdx(jobStage, evsNum));
                }
                BinCutInstanceSheetReaderHelpers.SetEvsConditionColIdx(i, text, jobStage, evsNum, _evs.ConditionColIdxDic);
            }
            #endregion
        }

        private void GetIndexBasic(int i, string text)
        {
            if (text.EqualsIgnoreCase(ConHeaderFlowName))
            {
                _basic.FlowNameColNumber = i;
            }
            else if (text.EqualsIgnoreCase(ConHeaderSkip))
            {
                _basic.SkipColNumber = i;
            }

            else if (text.EqualsIgnoreCase(ConHeaderInstance))
            {
                _basic.InstanceColNumber = i;
            }
            else if (text.EqualsIgnoreCase(ConHeaderOpcode))
            {
                _basic.OpcodeColNumber = i;
            }
            else if (MyRegex11().IsMatch(text))
            {
                if (_basic.PatternStartColNumber == 0)
                {
                    _basic.PatternStartColNumber = i;
                }
                _basic.PatternCounlumCnt++;
            }
            else if (text.EqualsIgnoreCase(ConHeaderEnbable))
            {
                _basic.EnableColNumber = i;
            }
            else if (MyRegex10().IsMatch(text))
            {
                _basic.BinTableEnableColNumber = i;
            }
            else if (text.EqualsIgnoreCase(ConHeaderComment))
            {
                _basic.CommnetColNumber = i;
            }
            else if (text.EqualsIgnoreCase(ConHeaderPatSetOrange))
            {
                _basic.PatSetNameOrangeColNumber = i;
            }
            else if (MyRegex12().IsMatch(text))
            {
                _basic.SubFlowColNumber = i;
            }
            else if (MyRegex13().IsMatch(text))
            {
                _basic.EnableNewColNumber = i;
            }
            else if (text.EqualsIgnoreCase(ConHeaderTestingStage) || MyRegex14().IsMatch(text) || MyRegex15().IsMatch(text))
            {
                _basic.JobTestStageColNumber = i;
            }
            else if (text.EqualsIgnoreCase(ConHeaderTimeSet))
            {
                _basic.TimeSetColNumber = i;
            }
            else if (text.EqualsIgnoreCase(ConHeaderShiftSpeed))
            {
                _basic.ShiftSpeedColNumber = i;
            }
            else if (MyRegex16().IsMatch(text) || MyRegex17().IsMatch(text))
            {
                _basic.SiteVarColNumber = i;
            }
        }
    }
}
