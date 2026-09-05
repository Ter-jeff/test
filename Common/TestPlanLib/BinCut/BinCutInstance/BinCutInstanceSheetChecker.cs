using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinCut.FlowNew;
using TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc;
using TestPlanLib.PatternListCsvFile;
using TestPlanLib.Utility;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public partial class BinCutInstanceSheetChecker
    {
        [GeneratedRegex(@"_\d+$", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"\s{2,}", RegexOptions.Compiled)]
        private static partial Regex MyRegex1();
        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();

        public EnumInstanceSheetType Type;

        public void WorkFlow(BinCutInstanceSheet binCutInstanceSheet, BinCutFlowSheets? binCutFlowSheets = null, NewBinCutFlowTables? newBinCutFlowTables = null, Dictionary<string, PatternData>? patternDatas = null, Dictionary<string, int>? timeSetDic = null, List<UfDigSrcSheet>? ufDigSrcSheets = null)
        {
            SetType();
            if (binCutFlowSheets != null)
            {
                if (newBinCutFlowTables?.Count > 0)
                {
                    List<string> flowNames = newBinCutFlowTables.GetFlowNames();
                    CheckFlowNameByModeSequence(binCutInstanceSheet, newBinCutFlowTables, flowNames);
                    CheckTestingStageByModeSequence(binCutInstanceSheet, newBinCutFlowTables);
                }
                else
                {
                    List<string> flowNames = binCutFlowSheets.GetFlowNames();
                    foreach (BinCutFlowTables binCutFlowSheet in binCutFlowSheets)
                    {
                        CheckFlowName(binCutInstanceSheet, binCutFlowSheet, flowNames);
                        CheckTestingStageByOldFlow(binCutInstanceSheet, binCutFlowSheet);
                    }
                }
            }

            if (timeSetDic!.Count != 0)
            {
                CheckTimeSet(binCutInstanceSheet, patternDatas!, timeSetDic);
            }
            //CheckPatterns(sheet);

            CheckDuplicatedInstance(binCutInstanceSheet);

            CheckDomain(binCutInstanceSheet);

            CheckPatternTimeSet(binCutInstanceSheet, patternDatas!);

            //CheckBistInitPattern(sheet);

            CheckUserDefinePatternSetName(binCutInstanceSheet);

            //CheckDssc(sheet);

            CheckUserFunction(binCutInstanceSheet, ufDigSrcSheets);
        }

        protected virtual void SetType()
        {
            Type = EnumInstanceSheetType.Bincut;
        }

        protected void CheckUserFunction(BinCutInstanceSheet binCutInstanceSheet, List<UfDigSrcSheet>? ufDigSrcSheets)
        {
            var ufDigSrcHeader = new List<string>();
            if (ufDigSrcSheets == null || ufDigSrcSheets.Count == 0)
            {
                return;
            }

            if (ufDigSrcSheets != null && ufDigSrcSheets.Count != 0)
            {
                foreach (UfDigSrcSheet ufDigSrcSheet in ufDigSrcSheets)
                {
                    ufDigSrcHeader.AddRange(ufDigSrcSheet.Rows[0].DsscDictionary.Select(key => key.Key));
                }
                ufDigSrcHeader = [.. ufDigSrcHeader.Distinct()];
            }

            foreach (BinCutInstanceRow row in binCutInstanceSheet.Rows)
            {
                if (!string.IsNullOrEmpty(row.UserFunction))
                {
                    string[] functions = row.UserFunction.Split([",", "DigSrc:", " "], StringSplitOptions.RemoveEmptyEntries);
                    foreach (string function in functions)
                    {
                        if (ufDigSrcHeader.Contains(function))
                        {
                            continue;
                        }

                        if (Type.Equals(EnumInstanceSheetType.Bincut))
                        {
                            binCutInstanceSheet.AddError(BinCutErrorType.E_UserFunction_01, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.UserFunctionColNumber, $"DigSrc {function} in UserFunction is not defined in UF_DigSrc sheet", [function]);
                        }
                        else
                        {
                            binCutInstanceSheet.AddError(ScanErrorType.E_UserFunction_01, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.UserFunctionColNumber, $"DigSrc {function} in UserFunction is not defined in UF_DigSrc sheet", [function]);
                        }
                    }
                }
                List<string> patternList = [.. ufDigSrcSheets!.SelectMany(x => x.Rows).Where(y => !string.IsNullOrEmpty(y.PatternName)).Select(z => z.PatternName)];
                foreach (string pattern in patternList)
                {
                    if (row.PatternList.Any(x => x.EqualsIgnoreCase(pattern)))
                    {
                        if (string.IsNullOrEmpty(row.UserFunction))
                        {
                            if (Type.Equals(EnumInstanceSheetType.Bincut))
                            {
                                binCutInstanceSheet.AddError(BinCutErrorType.E_UserFunction_02, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.UserFunctionColNumber, $"Need to assign digital source group for digital source pattern {pattern}", [pattern]);
                            }
                            else
                            {
                                binCutInstanceSheet.AddError(ScanErrorType.E_UserFunction_02, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.UserFunctionColNumber, $"Need to assign digital source group for digital source pattern {pattern}", [pattern]);
                            }
                        }
                        else
                        {
                            if (row.UserFunction.Split(':').Length != 2)
                            {

                                if (Type.Equals(EnumInstanceSheetType.Bincut))
                                {
                                    binCutInstanceSheet.AddError(BinCutErrorType.E_UserFunction_03, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.UserFunctionColNumber, "UserFunction format error, should be \"DigSrc:[DigSrcGroup]\"");
                                }
                                else
                                {
                                    binCutInstanceSheet.AddError(ScanErrorType.E_UserFunction_03, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.UserFunctionColNumber, "UserFunction format error, should be \"DigSrc:[DigSrcGroup]\"");
                                }
                            }
                        }
                    }
                }
            }
        }

        protected void CheckTimeSet(BinCutInstanceSheet binCutInstanceSheet, Dictionary<string, PatternData> patternDatas, Dictionary<string, int> timeSetSheets)
        {
            foreach (BinCutInstanceRow row in binCutInstanceSheet.Rows)
            {
                if (!string.IsNullOrEmpty(row.TimeSet))
                {
                    string timeSet = row.TimeSet.Contains(":") ? row.TimeSet.Split(':').First() : row.TimeSet;
                    if (!CheckTimeSetIsInKfolder(timeSetSheets, timeSet))
                    {
                        if (Type.Equals(EnumInstanceSheetType.Bincut))
                        {
                            binCutInstanceSheet.AddError(BinCutErrorType.E_TimeSet_01, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.TimeSetColNumber, $"TimeSet {timeSet} does not exist is K folder", [timeSet]);
                        }
                        else
                        {
                            binCutInstanceSheet.AddError(ScanErrorType.E_TimeSet_01, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.TimeSetColNumber, $"TimeSet {timeSet} does not exist is K folder", [timeSet]);
                        }
                    }
                }

                if (!string.IsNullOrEmpty(row.ShiftSpeed))
                {
                    bool hasScanTset = false;
                    foreach (string pattern in row.PayloadList)
                    {
                        if (!patternDatas.ContainsKey(pattern.ToLower()))
                        {
                            continue;
                        }

                        string tset = patternDatas[pattern.ToLower()].ScanTset;
                        if (!string.IsNullOrEmpty(tset))
                        {
                            hasScanTset = true;
                        }
                    }
                    if (!hasScanTset)
                    {
                        if (Type.Equals(EnumInstanceSheetType.Bincut))
                        {
                            binCutInstanceSheet.AddError(BinCutErrorType.E_TimeSet_02, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.PatternStartColNumber + row.InitList.Count, "Can't get any SCAN Tset in all payload");
                        }
                        else
                        {
                            binCutInstanceSheet.AddError(ScanErrorType.E_TimeSet_02, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.PatternStartColNumber + row.InitList.Count, "Can't get any SCAN Tset in all payload");
                        }
                    }
                }
            }
        }

        private static bool CheckTimeSetIsInKfolder(Dictionary<string, int> timeSetSheets, string timeset)
        {
            timeset = _regex.Replace(timeset, "");
            return timeSetSheets.Any(x => x.Key.EqualsIgnoreCase(timeset));
        }

        //For other tools
        public static void WorkFlow(BinCutInstanceSheet binCutInstanceSheet, BinCutFlowSheets? binCutFlowSheets = null, Dictionary<string, OriPatListItem>? patternDatas = null)
        {
            if (binCutFlowSheets != null)
            {
                List<string> flowNames = binCutFlowSheets.GetFlowNames();

                foreach (BinCutFlowTables binCutFlowSheet in binCutFlowSheets)
                {
                    CheckFlowName(binCutInstanceSheet, binCutFlowSheet, flowNames);
                    CheckTestingStageByOldFlow(binCutInstanceSheet, binCutFlowSheet);
                }
            }

            //CheckPatterns(sheet);

            CheckDuplicatedInstance(binCutInstanceSheet);

            CheckDomain(binCutInstanceSheet);

            CheckPatternTimeSet(binCutInstanceSheet, patternDatas!);

            //CheckPattern(ref sheet);

            //CheckBistInitPattern(sheet);

            CheckUserDefinePatternSetName(binCutInstanceSheet);
        }

        protected static void CheckUserDefinePatternSetName(BinCutInstanceSheet binCutInstanceSheet)
        {
            return;
        }

        protected static void CheckBistInitPattern(BinCutInstanceSheet binCutInstanceSheet)
        {
            for (int i = 0; i < binCutInstanceSheet.Rows.Count; i++)
            {
                for (int j = i + 1; j < binCutInstanceSheet.Rows.Count; j++)
                {
                    BinCutInstanceRow row1 = binCutInstanceSheet.Rows[i];
                    BinCutInstanceRow row2 = binCutInstanceSheet.Rows[j];

                    if (BinCutInstanceRowUtility.IsBist(row1.FlowName) &&
                        row1.FlowName.EqualsIgnoreCase(row2.FlowName) &&
                        row1.Instance.EqualsIgnoreCase(row2.Instance))
                    {
                        if (IsSameInit(row1, row2))
                        {
                            binCutInstanceSheet.AddError(BinCutErrorType.E_BistInitPattern_01, binCutInstanceSheet.SheetName, row2.RowNum, 0, "Row " + row2.RowNum + " and row " + row1.RowNum + " of the init patterns are the same!!!", [row2.RowNum.ToString(), row1.RowNum.ToString()]);
                        }
                    }
                }
            }
        }

        protected void CheckPatternTimeSet(BinCutInstanceSheet binCutInstanceSheet, Dictionary<string, PatternData> patternDatas)
        {
            if (patternDatas.Count != 0)
            {
                return;
            }

            var checkedPatterns = new List<string>();
            for (int i = 0; i < binCutInstanceSheet.Rows.Count; i++)
            {
                BinCutInstanceRow row = binCutInstanceSheet.Rows[i];
                if (row.Type == BincutInstanceType.Rtos)
                {
                    continue;
                }

                var timeSets = new List<string>();
                List<string> patList = GetPatList(row);
                for (int index = 0; index < patList.Count; index++)
                {
                    string pattern = patList[index];
                    if (string.IsNullOrEmpty(pattern))
                    {
                        continue;
                    }

                    if (patternDatas.ContainsKey(pattern.ToLower()))
                    {
                        timeSets.Add(patternDatas[pattern.ToLower()].TimeSetVersion);
                    }
                    if (checkedPatterns.Exists(x => x == pattern))
                    {
                        continue;
                    }

                    if (!patternDatas.ContainsKey(pattern.ToLower()))
                    {
                        string errorMessage = $"{pattern.ToLower()} can't be found the pattern in CSV !!!";
                        if (Type.Equals(EnumInstanceSheetType.Bincut))
                        {
                            binCutInstanceSheet.AddError(BinCutErrorType.W_Pattern_01, binCutInstanceSheet.SheetName, row.RowNum, 0, $"{pattern.ToLower()} can't be found the pattern in CSV !!!", [pattern.ToLower()], new ErrorInfo() { Comments = [pattern.ToLower()] });
                        }
                        else
                        {
                            binCutInstanceSheet.AddError(ScanErrorType.W_Pattern_01, binCutInstanceSheet.SheetName, row.RowNum, 0, $"{pattern.ToLower()} can't be found the pattern in CSV !!!", [pattern.ToLower()], new ErrorInfo() { Comments = [pattern.ToLower()] });
                        }
                    }
                    else
                    {
                        if (patternDatas[pattern.ToLower()].Use.EqualsIgnoreCase("dont_use"))
                        {
                            string errorMessage = $"This pattern {pattern.ToLower()} is \"Dont_useInCsv\" !!!";
                            if (Type.Equals(EnumInstanceSheetType.Bincut))
                            {
                                binCutInstanceSheet.AddError(BinCutErrorType.W_Pattern_02, binCutInstanceSheet.SheetName, row.RowNum, 0, $"This pattern {pattern.ToLower()} is \"Dont_useInCsv\" !!!", [pattern.ToLower()], new ErrorInfo() { Comments = [pattern.ToLower()] });
                            }
                            else
                            {
                                binCutInstanceSheet.AddError(ScanErrorType.W_Pattern_02, binCutInstanceSheet.SheetName, row.RowNum, 0, $"This pattern {pattern.ToLower()} is \"Dont_useInCsv\" !!!", [pattern.ToLower()], new ErrorInfo() { Comments = [pattern.ToLower()] });
                            }
                        }
                        else
                        {
                            if (patternDatas[pattern.ToLower()].FileVersion.EqualsIgnoreCase("n/a"))
                            {
                                string errorMessage = $"There are no FileVersion for {pattern.ToLower()} !!!";
                                if (Type.Equals(EnumInstanceSheetType.Bincut))
                                {
                                    binCutInstanceSheet.AddError(BinCutErrorType.W_Pattern_03, binCutInstanceSheet.SheetName, row.RowNum, 0, $"There are no FileVersion for {pattern.ToLower()} !!!", [pattern.ToLower()], new ErrorInfo() { Comments = [pattern.ToLower()] });
                                }
                                else
                                {
                                    binCutInstanceSheet.AddError(ScanErrorType.W_Pattern_03, binCutInstanceSheet.SheetName, row.RowNum, 0, $"There are no FileVersion for {pattern.ToLower()} !!!", [pattern.ToLower()], new ErrorInfo() { Comments = [pattern.ToLower()] });
                                }
                            }
                        }
                    }
                    checkedPatterns.Add(pattern);
                }

                CheckTimeSet(binCutInstanceSheet, row, timeSets, patList);
            }
        }

        private void CheckTimeSet(BinCutInstanceSheet binCutInstanceSheet, BinCutInstanceRow binCutInstanceRow, List<string> timeSets, List<string> patList)
        {
            timeSets = [.. timeSets.Where(x => !string.IsNullOrEmpty(x) && !x.EqualsIgnoreCase("NA") && !x.EqualsIgnoreCase("N/A")).Distinct()];
            if (timeSets.Count == 0)
            {
                if (Type.Equals(EnumInstanceSheetType.Bincut))
                {
                    binCutInstanceSheet.AddError(BinCutErrorType.E_PatternTimeSet_01, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, 0, "This test can not find any time set by pattern dashboard !!!");
                }
                else if (patList.Count != 0)
                {
                    binCutInstanceSheet.AddError(ScanErrorType.E_PatternTimeSet_01, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, 0, "This test can not find any time set by pattern dashboard !!!");
                }
            }
            else if (timeSets.Count != 1)
            {
                if (binCutInstanceRow.TimeSet.Length != 0)
                {
                    if (Type.Equals(EnumInstanceSheetType.Bincut))
                    {
                        binCutInstanceSheet.AddError(BinCutErrorType.W_PatternTimeSet_01, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, 0, $"There are multi time set {string.Join(",", timeSets)} by patterns overwrite to {binCutInstanceRow.TimeSet}!!!", [string.Join(",", timeSets), binCutInstanceRow.TimeSet]);
                    }
                    else
                    {
                        binCutInstanceSheet.AddError(ScanErrorType.W_PatternTimeSet_01, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, 0, $"There are multi time set {string.Join(",", timeSets)} by patterns overwrite to {binCutInstanceRow.TimeSet}!!!", [string.Join(",", timeSets), binCutInstanceRow.TimeSet]);
                    }
                }
                else
                {
                    if (Type.Equals(EnumInstanceSheetType.Bincut))
                    {
                        binCutInstanceSheet.AddError(BinCutErrorType.E_PatternTimeSet_02, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, 0, $"There are multi time set {string.Join(",", timeSets)} by patterns !!!", [string.Join(",", timeSets)]);
                    }
                    else
                    {
                        binCutInstanceSheet.AddError(ScanErrorType.E_PatternTimeSet_02, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, 0, $"There are multi time set {string.Join(",", timeSets)} by patterns !!!", [string.Join(",", timeSets)]);
                    }
                }
            }
        }

        private static List<string> GetPatList(BinCutInstanceRow binCutInstanceRow)
        {
            List<string> patlist = binCutInstanceRow.PatternList;
            if (patlist.Exists(x => x.Contains('+')))
            {
                patlist = [];
                foreach (string pat in binCutInstanceRow.PatternList)
                {
                    if (pat.Contains('+'))
                    {
                        patlist.AddRange([.. pat.Split(['+', ' '])]);
                    }
                    else
                    {
                        patlist.Add(pat);
                    }
                }
            }

            return patlist;
        }

        protected static void CheckPatternTimeSet(BinCutInstanceSheet binCutInstanceSheet, Dictionary<string, OriPatListItem> patternDatas)
        {
            if (patternDatas.Count != 0)
            {
                return;
            }

            for (int i = 0; i < binCutInstanceSheet.Rows.Count; i++)
            {
                BinCutInstanceRow row = binCutInstanceSheet.Rows[i];
                if (row.Type == BincutInstanceType.Rtos)
                {
                    continue;
                }

                var timeSets = new List<string>();
                List<string> patlist = row.PatternList;
                if (patlist.Exists(x => x.Contains('+')))
                {
                    patlist = [];
                    foreach (string pat in row.PatternList)
                    {
                        if (pat.Contains('+'))
                        {
                            patlist.AddRange([.. pat.Split(['+', ' '])]);
                        }
                        else
                        {
                            patlist.Add(pat);
                        }
                    }
                }

                for (int index = 0; index < patlist.Count; index++)
                {
                    string pattern = patlist[index];
                    if (string.IsNullOrEmpty(pattern))
                    {
                        continue;
                    }

                    if (!patternDatas.TryGetValue(pattern, out OriPatListItem? value))
                    {
                        binCutInstanceSheet.AddError(BinCutErrorType.W_Pattern_04, binCutInstanceSheet.SheetName, row.RowNum, 0, pattern.ToUpper() + " can't be found the pattern in CSV !!!", [pattern.ToUpper()]);
                    }
                    else
                    {
                        if (patternDatas[pattern].UseNoUse.EqualsIgnoreCase("dont_use"))
                        {
                            string errorMessage = $"This pattern {pattern.ToLower()} is \"Dont_useInCsv\" !!!";
                            if (!binCutInstanceSheet.GetErrors().Exists(x => x.Message.EqualsIgnoreCase(errorMessage)))
                            {
                                binCutInstanceSheet.AddError(BinCutErrorType.W_Pattern_05, binCutInstanceSheet.SheetName, row.RowNum, 0, $"This pattern {pattern.ToLower()} is \"Dont_useInCsv\" !!!", [pattern.ToLower()], new ErrorInfo() { Comments = [pattern.ToLower()] });
                            }
                        }

                        timeSets.Add(value.TimeSetVersion.Replace(".TXT", ""));
                    }
                }
                timeSets = [.. timeSets.Where(x => !string.IsNullOrEmpty(x) && !x.EqualsIgnoreCase("NA") &&
                    !x.EqualsIgnoreCase("N/A")).Distinct()];
                if (timeSets.Count == 0)
                {
                    binCutInstanceSheet.AddError(BinCutErrorType.W_PatternTimeSet_02, binCutInstanceSheet.SheetName, row.RowNum, 0, "This test can not find any time set by pattern dashboard !!!");
                }
                else if (timeSets.Count != 1)
                {
                    if (row.TimeSet.Length != 0)
                    {
                        binCutInstanceSheet.AddError(BinCutErrorType.W_PatternTimeSet_03, binCutInstanceSheet.SheetName, row.RowNum, 0, $"There are multi time set {string.Join(",", timeSets)} by patterns overwrite to {row.TimeSet}!!!", [string.Join(",", timeSets), row.TimeSet]);
                    }
                    else
                    {
                        binCutInstanceSheet.AddError(BinCutErrorType.E_PatternTimeSet_03, binCutInstanceSheet.SheetName, row.RowNum, 0, $"There are multi time set {string.Join(",", timeSets)} by patterns !!!", [string.Join(",", timeSets)]);
                    }
                }
            }
        }

        protected static void CheckDomain(BinCutInstanceSheet binCutInstanceSheet)
        {
            for (int i = 0; i < binCutInstanceSheet.Rows.Count; i++)
            {
                BinCutInstanceRow row = binCutInstanceSheet.Rows[i];
                if (row.Type == BincutInstanceType.Rtos)
                {
                    continue;
                }

                Domain domain = BinCutInstanceRowUtility.GetDomainByFlowName(row.FlowName);
                Domain patdomain = BinCutInstanceRowUtility.GetDomainByPattern(row.PatternList);
                int column = binCutInstanceSheet.FlowNameColNumber == -1 ? binCutInstanceSheet.SubFlowColNumber : binCutInstanceSheet.FlowNameColNumber;
                if (string.IsNullOrEmpty(domain.Name) && string.IsNullOrEmpty(patdomain.Name))
                {
                    binCutInstanceSheet.AddError(BinCutErrorType.W_Domain_01, binCutInstanceSheet.SheetName, row.RowNum, 0, "Unable to get domain from flow name and pattern !!!");
                }
                else if (string.IsNullOrEmpty(domain.Name) && !string.IsNullOrEmpty(patdomain.Name))
                {
                    binCutInstanceSheet.AddError(BinCutErrorType.W_Domain_02, binCutInstanceSheet.SheetName, row.RowNum, column, "Unable to get domain from flow name !!!");
                }
                else if (!string.IsNullOrEmpty(domain.Name) && string.IsNullOrEmpty(patdomain.Name))
                {
                    binCutInstanceSheet.AddError(BinCutErrorType.W_Domain_03, binCutInstanceSheet.SheetName, row.RowNum, 0, "Unable to get domain from pattern !!!");
                }
                else
                {
                    if (!string.IsNullOrEmpty(patdomain.Name) && !patdomain.Equals(domain))
                    {
                        //string error = $"This domain of flowName {row.FlowName} is different from the pattern {patdomain.Name}!!!";
                        binCutInstanceSheet.AddError(BinCutErrorType.W_Domain_04, binCutInstanceSheet.SheetName, row.RowNum, column, $"This domain of flowName {row.FlowName} is different from the pattern {patdomain.Name}!!!", [row.FlowName, patdomain.Name]);
                    }
                }
                patdomain.Name = "";
                var tmp = new Domain { Name = "" };
                int col = 0;
                foreach (string pattern in row.PatternList)
                {
                    if (string.IsNullOrEmpty(pattern))
                    {
                        continue;
                    }

                    col++;
                    List<string> arr = [.. pattern.Split('_')];
                    if (arr.Count > 2)
                    {
                        if (arr[2].EqualsIgnoreCase("A"))
                        {
                            patdomain.Name = "HardIP";
                        }
                        else if (arr[2].EqualsIgnoreCase("P"))
                        {
                            patdomain.Name = "HardIP";
                        }
                        else if (arr[2].EqualsIgnoreCase("V"))
                        {
                            patdomain.Name = "HardIP";
                        }
                        else if (arr[2].EqualsIgnoreCase("H"))
                        {
                            patdomain.Name = "HardIP";
                        }
                        else if (arr[2].EqualsIgnoreCase("C"))
                        {
                            patdomain.Name = "CPU";
                        }
                        else if (arr[2].EqualsIgnoreCase("L"))
                        {
                            patdomain.Name = "GPU";
                        }
                        else if (arr[2].EqualsIgnoreCase("S"))
                        {
                            patdomain.Name = "SOC";
                        }

                        if (!string.IsNullOrEmpty(tmp.Name) && !patdomain.Equals(tmp))
                        {
                            if (string.IsNullOrEmpty(domain.Name))
                            {
                                binCutInstanceSheet.AddError(BinCutErrorType.W_Domain_05, binCutInstanceSheet.SheetName, row.RowNum, col + binCutInstanceSheet.PatternStartColNumber - 1, "This domain is different from other patterns!!!");
                            }
                        }
                        tmp.Name = domain.Name;
                    }
                }
            }
        }

        private static void CheckFlowNameByModeSequence(BinCutInstanceSheet binCutInstanceSheet, NewBinCutFlowTables newBinCutFlowTables, List<string> flowNames)
        {
            if (newBinCutFlowTables == null)
            {
                return;
            }

            CheckJobsByModeSequence(binCutInstanceSheet, newBinCutFlowTables);

            CheckFlowNameInternal(binCutInstanceSheet, flowNames, row => FlowNameMatchJob(newBinCutFlowTables, row), isRegexMatch: true);
        }

        private static bool FlowNameMatchJob(NewBinCutFlowTables newBinCutFlowTables, BinCutInstanceRow binCutInstanceRow)
        {
            bool flag = false;
            foreach (NewBinCutFlowTable table in newBinCutFlowTables)
            {
                string jobTestStage = binCutInstanceRow.JobTestStage;
                if (IsMatch(binCutInstanceRow.FlowName, jobTestStage, null, table))
                {
                    flag = true;
                    break;
                }
            }

            return flag;
        }

        private static void CheckFlowName(BinCutInstanceSheet binCutInstanceSheet, BinCutInstanceRow binCutInstanceRow)
        {
            if (_regex2.IsMatch(binCutInstanceRow.FlowNameOri))
            {
                int col = binCutInstanceSheet.FlowNameColNumber == -1 ? binCutInstanceSheet.SubFlowColNumber : binCutInstanceSheet.FlowNameColNumber;
                binCutInstanceSheet.AddError(BinCutErrorType.E_FlowName_03, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, col, "The flowName \"" + binCutInstanceRow.FlowNameOri + "\" with extra spaces !!!", [binCutInstanceRow.FlowNameOri]);
            }
        }

        private static void CheckFlowName(BinCutInstanceSheet binCutInstanceSheet, List<BinCutFlowTable> binCutFlowTables, List<string> flowNames)
        {
            if (binCutFlowTables == null)
            {
                return;
            }

            CheckJobs(binCutInstanceSheet, binCutFlowTables);

            CheckFlowNameInternal(binCutInstanceSheet, flowNames, row => FlowNameMatchJob(binCutFlowTables, row), isRegexMatch: false);
        }

        private static void CheckFlowNameInternal(BinCutInstanceSheet binCutInstanceSheet, List<string> flowNames, Func<BinCutInstanceRow, bool> matchJobFunc, bool isRegexMatch)
        {
            foreach (BinCutInstanceRow row in binCutInstanceSheet.Rows)
            {
                CheckFlowName(binCutInstanceSheet, row);

                bool nameExists = isRegexMatch ? flowNames.Exists(x => Regex.IsMatch(row.FlowName, x, RegexOptions.IgnoreCase))
                    : flowNames.Exists(x => x.EqualsIgnoreCase(row.FlowName));

                if (!nameExists)
                {
                    int col = binCutInstanceSheet.FlowNameColNumber == -1 ? binCutInstanceSheet.SubFlowColNumber : binCutInstanceSheet.FlowNameColNumber;
                    binCutInstanceSheet.AddError(BinCutErrorType.W_FlowName_02, binCutInstanceSheet.SheetName, row.RowNum, col, $"The flowName \"{row.FlowName}\" not found in the flow sheet !!!", [row.FlowName]);
                    continue;
                }

                if (!matchJobFunc(row))
                {
                    binCutInstanceSheet.AddError(BinCutErrorType.E_Missing_01, binCutInstanceSheet.SheetName, row.RowNum, 0, $"The instance row {row.RowNum} not be used in the flow sheet ({row.JobTestStage})!!!", [row.RowNum.ToString(), row.JobTestStage]);
                }
            }
        }

        private static bool FlowNameMatchJob(List<BinCutFlowTable> binCutFlowTables, BinCutInstanceRow binCutInstanceRow)
        {
            bool flag = false;
            foreach (BinCutFlowTable table in binCutFlowTables)
            {
                string jobTestStage = binCutInstanceRow.JobTestStage;
                if (IsMatch(binCutInstanceRow.FlowName, jobTestStage, table))
                {
                    flag = true;
                    break;
                }
            }

            return flag;
        }

        private static void CheckJobsByModeSequence(BinCutInstanceSheet binCutInstanceSheet, NewBinCutFlowTables newBinCutFlowTables)
        {
            var exceptJob = new List<string> { "QA" };
            var totalJobs = newBinCutFlowTables.Select(x => x.FinalJob).ToList();
            totalJobs = [.. totalJobs.Except(exceptJob)];
            var groups = newBinCutFlowTables.SelectMany(x => x.Rows).GroupBy(x => x.RowNum).ToList();
            foreach (IGrouping<int, NewBinCutFlowSheetRow> group in groups)
            {
                string mode = group.First().PerformanceMode;
                var flowNamesByRows1 = new List<string>();
                flowNamesByRows1.AddRange(group.SelectMany(x => x.SubFlows.ToList()));
                flowNamesByRows1 = [.. flowNamesByRows1.Where(x => x != "0" && !string.IsNullOrEmpty(x)).Distinct()];
                foreach (string flowNamesByRow in flowNamesByRows1)
                {
                    IEnumerable<string> jobs1 = group.Where(x => x.SubFlows.ToList().Contains(flowNamesByRow)).Select(y => y.Job.Trim());
                    jobs1 = jobs1.Except(exceptJob);
                    CheckJobs(binCutInstanceSheet, flowNamesByRow, totalJobs, jobs1, mode);
                }
            }
        }

        private static void CheckJobs(BinCutInstanceSheet binCutInstanceSheet, List<BinCutFlowTable> binCutFlowTables)
        {
            var exceptJob = new List<string> { "QA" };
            var totalJobs = binCutFlowTables.SelectMany(x => x.FinalJob).Select(y => y.Trim()).ToList();
            totalJobs = [.. totalJobs.Except(exceptJob)];
            var groups = binCutFlowTables.SelectMany(x => x.Rows).GroupBy(x => x.RowNum).ToList();
            foreach (IGrouping<int, BinCutFlowSheetRow> group in groups)
            {
                string mode = group.First().PerformanceMode;
                var flowNamesByRows1 = new List<string>();
                flowNamesByRows1.AddRange(group.SelectMany(x => x.Atpg.Split(';').ToList()));
                flowNamesByRows1 =
                    [.. flowNamesByRows1.Where(x => x != "0" && !string.IsNullOrEmpty(x)).Distinct()];
                foreach (string flowNamesByRow in flowNamesByRows1)
                {
                    IEnumerable<string> jobs1 = group.Where(x => x.Atpg.Split(';').ToList().Contains(flowNamesByRow)).SelectMany(x => x.Job).Select(y => y.Trim());
                    jobs1 = jobs1.Except(exceptJob);
                    CheckJobs(binCutInstanceSheet, flowNamesByRow, totalJobs, jobs1, mode);
                }

                var flowNamesByRows2 = new List<string>();
                flowNamesByRows2.AddRange(group.SelectMany(x => x.Mbist.Split(';').ToList()));
                flowNamesByRows2 =
                    [.. flowNamesByRows2.Where(x => x != "0" && !string.IsNullOrEmpty(x)).Distinct()];
                foreach (string flowNamesByRow in flowNamesByRows2)
                {
                    IEnumerable<string> jobs2 = group.Where(x => x.Mbist.Split(';').ToList().Contains(flowNamesByRow)).SelectMany(x => x.Job).Select(y => y.Trim());
                    jobs2 = jobs2.Except(exceptJob);
                    CheckJobs(binCutInstanceSheet, flowNamesByRow, totalJobs, jobs2, mode);
                }

                var flowNamesByRows3 = new List<string>();
                flowNamesByRows3.AddRange(group.SelectMany(x => x.SpiRtos.Split(';').ToList()));
                flowNamesByRows3 =
                    [.. flowNamesByRows3.Where(x => x != "0" && !string.IsNullOrEmpty(x)).Distinct()];
                foreach (string flowNamesByRow in flowNamesByRows3)
                {
                    IEnumerable<string> jobs3 = group.Where(x => x.SpiRtos.Split(';').ToList().Contains(flowNamesByRow)).SelectMany(x => x.Job).Select(y => y.Trim());
                    jobs3 = jobs3.Except(exceptJob);
                    CheckJobs(binCutInstanceSheet, flowNamesByRow, totalJobs, jobs3, mode);
                }
            }
        }

        private static void CheckJobs(BinCutInstanceSheet binCutInstanceSheet, string flowNamesByRow, List<string> totalJobs, IEnumerable<string> jobs, string mode)
        {
            IEnumerable<BinCutInstanceRow> rows = binCutInstanceSheet.Rows.Where(x => x.FlowName.EqualsIgnoreCase(flowNamesByRow));
            if (flowNamesByRow.Contains('#'))
            {
                string newFlowName = flowNamesByRow.Replace("#", "[ |_]");
                rows = binCutInstanceSheet.Rows.Where(x => Regex.IsMatch(x.FlowName, newFlowName, RegexOptions.IgnoreCase));
            }
            foreach (BinCutInstanceRow row in rows)
            {
                List<string> targetJobs = BinCutInstanceRowUtility.GetEnableJobs(row, totalJobs);
                IEnumerable<string> remove = targetJobs.Except(jobs, StringExtensions.IgnoreCase);
                if (remove.Any())
                {
                    int col = binCutInstanceSheet.EnableColNumber == -1 ? binCutInstanceSheet.EnableNewColNumber : binCutInstanceSheet.EnableColNumber;
                    string errMsg = $"Jobs in the flow sheet (Performance Mode / FlowName : {mode} / {row.FlowName}) is {string.Join(",", jobs)} , but jobs in BinCut_instance is {row.JobTestStage} !!!";
                    binCutInstanceSheet.AddError(BinCutErrorType.E_JobMisMatch_01, row.SheetName, row.RowNum, col, $"Jobs in the flow sheet (Performance Mode / FlowName : {mode} / {row.FlowName}) is {string.Join(",", jobs)} , but jobs in BinCut_instance is {row.JobTestStage} !!!", [mode, row.FlowName, string.Join(",", jobs), row.JobTestStage]);
                }
            }
        }

        private static bool IsMatch(string flowName, string testingStage, BinCutFlowTable? binCutFlowTable = null, NewBinCutFlowTable? newBinCutFlowTable = null)
        {
            bool foundJob = FoundJob(testingStage, binCutFlowTable!, newBinCutFlowTable);

            if (foundJob)
            {
                bool flag = false;
                if (newBinCutFlowTable != null)
                {
                    foreach (NewBinCutFlowSheetRow row in newBinCutFlowTable.Rows)
                    {
                        List<string> conditions = row.SubFlows.ConvertAll(x => x.Trim().Replace("#", "[ |_]").Split([':'], StringSplitOptions.RemoveEmptyEntries).Last().Trim());
                        if (conditions.Exists(x => Regex.IsMatch(flowName, x, RegexOptions.IgnoreCase)))
                        {
                            flag = true;
                            break;
                        }
                    }
                }
                else
                {
                    foreach (BinCutFlowSheetRow row in binCutFlowTable!.Rows)
                    {
                        var atpgs = row.Atpg.Split(';').Select(x => x.Trim()).ToList();
                        if (atpgs.Exists(x => x.EqualsIgnoreCase(flowName)))
                        {
                            flag = true;
                            break;
                        }
                        var mbists = row.Mbist.Split(';').Select(x => x.Trim()).ToList();
                        if (mbists.Exists(x => x.EqualsIgnoreCase(flowName)))
                        {
                            flag = true;
                            break;
                        }
                        var spiRtoss = row.SpiRtos.Split(';').Select(x => x.Trim()).ToList();
                        if (spiRtoss.Exists(x => x.EqualsIgnoreCase(flowName)))
                        {
                            flag = true;
                            break;
                        }
                    }
                }
                return flag;
            }
            return false;
        }

        private static bool FoundJob(string testingStage, BinCutFlowTable binCutFlowTable, NewBinCutFlowTable? newBinCutFlowTable = null)
        {
            if (string.IsNullOrEmpty(testingStage) ||
                testingStage.EqualsIgnoreCase("All"))
            {
                return true;
            }

            var testingStageArr = testingStage.Replace(" ", "").Split(',').Select(x => x.Trim()).ToList();
            if (newBinCutFlowTable != null)
            {
                bool isExist = testingStageArr.Exists(x => x.EqualsIgnoreCase(newBinCutFlowTable.FinalJob));
                if (isExist)
                {
                    return true;
                }
            }
            else
            {
                foreach (string finaljob in binCutFlowTable.FinalJob)
                {
                    bool isExist = testingStageArr.Exists(x => x.EqualsIgnoreCase(finaljob));
                    if (isExist)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        protected static void CheckDuplicatedInstance(BinCutInstanceSheet binCutInstanceSheet)
        {
            return;
        }

        private static void CheckTestingStage(BinCutInstanceSheet binCutInstanceSheet, List<string> jobs, List<string> jobnames)
        {
            foreach (BinCutInstanceRow row in binCutInstanceSheet.Rows)
            {
                if (string.IsNullOrEmpty(row.JobTestStage))
                {
                    continue;
                }

                IEnumerable<string> arr = row.JobTestStage.Split(',').Select(x => x.Trim());
                foreach (string job in arr)
                {
                    if (job.EqualsIgnoreCase("All") || string.IsNullOrEmpty(job))
                    {
                        continue;
                    }

                    string jobName = job.TrimStart('!');
                    if (!jobs.Exists(x => x.EqualsIgnoreCase(jobName)) &&
                        !jobnames.Exists(x => x.EqualsIgnoreCase(jobName)))
                    {
                        int col = binCutInstanceSheet.TestingStageColNumber == -1 ? binCutInstanceSheet.JobTestStageColNumber : binCutInstanceSheet.TestingStageColNumber;
                        binCutInstanceSheet.AddError(BinCutErrorType.E_TestingStage_01, binCutInstanceSheet.SheetName, row.RowNum, col, "The job name of Testing Stage \"" + job + "\" is unknown !!!", [job]);
                    }
                }
            }
        }

        protected static void CheckTestingStageByOldFlow(BinCutInstanceSheet binCutInstanceSheet, List<BinCutFlowTable> binCutFlowTables)
        {
            if (binCutFlowTables == null)
            {
                return;
            }

            var jobs = binCutFlowTables.SelectMany(x => x.FinalJob).Select(y => y.Trim()).ToList();
            List<string> jobnames = binCutFlowTables.ConvertAll(x => x.JobName);
            CheckTestingStage(binCutInstanceSheet, jobs, jobnames);
        }

        protected static void CheckTestingStageByModeSequence(BinCutInstanceSheet binCutInstanceSheet, NewBinCutFlowTables newBinCutFlowTables)
        {
            if (newBinCutFlowTables == null)
            {
                return;
            }

            var jobs = newBinCutFlowTables.Select(x => x.FinalJob.Trim()).ToList();
            var jobnames = newBinCutFlowTables.Select(x => x.JobName).ToList();
            CheckTestingStage(binCutInstanceSheet, jobs, jobnames);
        }

        private static bool IsSameInit(BinCutInstanceRow row1, BinCutInstanceRow row2)
        {
            if (row1.InitList.Count != row2.InitList.Count)
            {
                return false;
            }

            for (int index = 0; index < row1.InitList.Count; index++)
            {
                if (!string.IsNullOrEmpty(row1.InitList[index]) && !string.IsNullOrEmpty(row2.InitList[index]) &&
                    !row1.InitList[index].EqualsIgnoreCase(row2.InitList[index]))
                {
                    return false;
                }
            }
            return true;
        }
    }
}
