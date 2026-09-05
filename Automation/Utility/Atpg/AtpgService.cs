using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Basic.Business.GenPatSet.Business;
using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Singleton;
using TestPlanLib.VbtLib;

namespace Automation.Utility.Atpg
{
    public static class AtpgService
    {
        private static readonly Regex _regex5 = new Regex(@"\w+:(?<str>\w+)\(\w+\)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        internal delegate List<T> BinOutFactory<T>(List<string> flagList, string voltage, string binOp, BinCutFinalInstanceRow dataRow);

        public static void SetDigSrc(
            BinCutFinalInstanceRow row,
            List<string> ufDigSrcPats,
            UserFunctionTableRow ufFuncSetting,
            HardIpInfos hardIpInfos,
            string selsramSetting,
            List<string> patSets,
            ref Function vbtFunctionBase
        )
        {
            var insDigSrcPats = new List<string>();
            var digSrcEquation = new List<string>();
            var digSrcAssignment = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string mfstpDigSrc = ufFuncSetting?.MultiFstpSetting ?? "";
            if (string.IsNullOrEmpty(mfstpDigSrc))
            {
                string[] ufString = row.BinCutInstanceRow?.UserFunction?.Split(':');
                mfstpDigSrc = (ufString?.Length > 1) ? ufString[1] : mfstpDigSrc;
            }
            int mfstpIndex = 0;
            string[] mfstpSet = mfstpDigSrc.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string pat in patSets ?? Enumerable.Empty<string>())
            {
                if (ufDigSrcPats.Exists(x => x.Equals(pat, StringComparison.OrdinalIgnoreCase)))
                {
                    string assignName = "D".ToExcelColName(mfstpIndex);
                    insDigSrcPats.Add(pat);
                    digSrcEquation.Add(assignName);
                    if (mfstpSet.Length > mfstpIndex)
                    {
                        digSrcAssignment.Add($"{assignName}=DSSC({mfstpSet[mfstpIndex]})");
                        mfstpIndex++;
                    }
                }
                else if (pat.ContainsIgnoreCase("DSSC") && !pat.ContainsIgnoreCase("OCCM"))
                {
                    insDigSrcPats.Add(pat);
                    digSrcEquation.Add("C");
                    digSrcAssignment.Add($"C=Selsram({selsramSetting})");
                }
                else
                {
                    digSrcEquation.Add("");
                }
            }

            if (insDigSrcPats.Any())
            {
                string sendPinName = "JTAG_TDI";
                if (hardIpInfos != null)
                {
                    HardIpInfo target = hardIpInfos.GetHardIpInfo(insDigSrcPats.FirstOrDefault());
                    if (target != null && !string.IsNullOrEmpty(target.SendPinName))
                    {
                        sendPinName = target.SendPinName;
                    }
                }
                vbtFunctionBase.SetParamValue("digSrcPin", sendPinName);
                vbtFunctionBase.SetParamValue("digSrcEquation", string.Join("|", digSrcEquation));
                vbtFunctionBase.SetParamValue("digSrcAssignment", string.Join(";", digSrcAssignment));
            }
        }

        public static bool IsSamePatternList(List<string> row1, List<string> row2)
        {
            if (row1.Count != row2.Count)
            {
                return false;
            }

            for (int index = 0; index < row1.Count; index++)
            {
                if (!string.IsNullOrEmpty(row1[index]) && !string.IsNullOrEmpty(row2[index]) &&
                    !row1[index].Equals(row2[index], StringComparison.CurrentCultureIgnoreCase))
                {
                    return false;
                }
            }
            return true;
        }

        public static void GenerateCSharpInstanceRow(BinCutFinalInstanceRow row, InstanceRow instanceRow, Function vbtFunctionBase)
        {
            if (row.GetDcCategory().Contains("_EQN") && LocalSpecs.EquationVoltagesFileName != "N/A")
            {
                vbtFunctionBase.SetParamValue("ateTestCondition", row.BinCutInstanceRow.DCcategory);
                List<string> items = row.PatSetName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                items.Add("EQN");
                row.PatSetName = string.Join("_", items.Where(x => !string.IsNullOrEmpty(x)));
                instanceRow.TestName = row.GetParameter();
            }
            vbtFunctionBase.SetParamValue("Patterns", row.PatternList.Count == 1 ? row.PatternList[0] : row.PatSetName);
            if (row.IsBurstNonBinCutInstance())
            {
                vbtFunctionBase.SetParamValue("ResultMode", "1");
            }
            else
            {
                vbtFunctionBase.SetParamValue("ResultMode", "0");
            }

            vbtFunctionBase.SetParamValue("RelayMode", "1");

            vbtFunctionBase.SetParamValue("harvestOtherFailFlag", string.Join(",", row.GetPinGroupFailFlags()));
            vbtFunctionBase.SetParamValue("isHarvesting", row.BinCutInstanceRow?.IsHarvesting);

            UserFunctionTableRow ufFuncSetting = null;
            if (!string.IsNullOrEmpty(row.BinCutInstanceRow.UserFunction) && TestPlanStatic.UserFunctionSheet != null)
            {
                ufFuncSetting = TestPlanStatic.UserFunctionSheet.Rows
                .FirstOrDefault(x => x.UserFunction.Equals(row.BinCutInstanceRow.UserFunction, StringComparison.OrdinalIgnoreCase));
            }
            if (ufFuncSetting != null)
            {
                TestPlanStatic.UserFunctionSheet.ArgumentSetting(row.BinCutInstanceRow.UserFunction, vbtFunctionBase);
            }
            List<string> ufDigSrcPats = TestPlanStatic.UfDigSrcSheets
                .SelectMany(x => x.Rows)
                .Where(y => !string.IsNullOrEmpty(y.PatternName)).Select(z => z.PatternName).ToList();
            SetDigSrc(row, ufDigSrcPats, ufFuncSetting, LocalSpecs.HardIpInfos, "", row.PatternList, ref vbtFunctionBase);
            instanceRow.VbtType = ".NET";
            instanceRow.VbtName = vbtFunctionBase.FullFunctionName;
            instanceRow.ArgList = vbtFunctionBase.Parameters;
            instanceRow.Args = vbtFunctionBase.ArgList;
        }

        public static string GenerateGetAcCategory(BinCutInstanceRow instanceDataRow, string timeSet)
        {
            string timeSetVersion = AcTSetCategoryMapSingleton.Instance().GetTimeSetVersion(timeSet);
            timeSetVersion = string.IsNullOrEmpty(timeSetVersion) ? timeSet : timeSetVersion;
            string acCategory = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSetVersion, BlockType.Scan);
            if (acCategory.Equals("TBD", StringComparison.OrdinalIgnoreCase))
            {
                acCategory = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSetVersion);
            }

            if (instanceDataRow != null && !string.IsNullOrEmpty(instanceDataRow.ShiftSpeed))
            {
                acCategory += "_" + instanceDataRow.ShiftSpeed;
            }

            return acCategory;
        }

        private static string GetMode(string pattern)
        {
            string performanceMode = "";
            string[] subStrings = pattern.Split('_');

            if (subStrings.Length > 9 && Regex.IsMatch(subStrings[9], @"^M[a-zA-Z]+\d+$"))
            {
                if (!Regex.IsMatch(subStrings[9], "999|010"))
                {
                    performanceMode = subStrings[9];
                }
            }

            return performanceMode;
        }

        public static string GetPerformanceMode(List<string> initList, List<string> modeList)
        {
            string performanceMode = "";
            foreach (string init in initList)
            {
                string mode = GetMode(init);
                performanceMode = string.IsNullOrEmpty(mode) ? performanceMode : mode;
            }

            if (performanceMode == "")
            {
                foreach (string init in initList)
                {
                    string[] subStrings = init.Split('_');
                    if (subStrings.Length > 9)
                    {
                        string module = ModuleSingleton.GetModuleByPerformanceMode(subStrings[9]);
                        if (module == ModuleSingleton.Instance().ModuleCpu || module == ModuleSingleton.Instance().ModuleGfx || module == ModuleSingleton.Instance().ModuleSoc)
                        {
                            if (!Regex.IsMatch(subStrings[9], "999|010"))
                            {
                                performanceMode = subStrings[9];
                            }
                        }
                    }
                }
            }

            return performanceMode;
        }

        public static void CheckPatternPins(
            List<string> pins,
            List<string> patternList,
            PinMapSheet pinMap,
            ExcelPackage missingPinsReport,
            List<string> harvestCheckedPattern
        )
        {
            var patternPathListInK = AcTSetCategoryMapSingleton.Instance()
                .PatternPathListInK
                .ToList();
            foreach (string pattern in patternList)
            {
                //Skip pattern which was checked
                if (harvestCheckedPattern.Contains(pattern))
                {
                    continue;
                }

                List<string> patternPins = GetPatternPins(pattern, patternPathListInK);
                if (patternPins == null)
                {
                    continue;
                }
                List<string> ioPins = pinMap.GetIoPins(patternPins);
                var except = pins.Except(ioPins, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
                var exceptPattern = ioPins.Except(pins, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
                if (exceptPattern.Count != 0)
                {
                    WriteMissingPin(
                        missingPinsReport.Workbook.Worksheets["Extra harvest pins"],
                        pattern,
                        exceptPattern,
                        harvestCheckedPattern);
                    string msg = pattern
                        + " have more pins than harvest pin Groups : "
                        + string.Join(",", exceptPattern)
                        + "!";
                    Response.Report(msg, EnumMessageLevel.Error, 10);
                }

                if (except.Count != 0)
                {
                    WriteMissingPin(
                        missingPinsReport.Workbook.Worksheets["Miss harvest pins"],
                        pattern,
                        except,
                        harvestCheckedPattern);
                    string msg = pattern
                        + " have not these pins : "
                        + string.Join(",", except)
                        + "!";
                    Response.Report(msg, EnumMessageLevel.Error, 10);
                }
            }
        }

        public static void CheckHarvestGroupPins(
            BinCutInstanceSheet sheet,
            ExcelPackage missingPinsReport,
            List<string> harvestCheckedPattern
        )
        {
            foreach (BinCutInstanceRow row in sheet.Rows)
            {
                if (string.IsNullOrEmpty(row.PatternPinGroup))
                {
                    continue;
                }
                var pins = new List<string>();
                List<string> pinGroups = GetGroupPins(row.PatternPinGroup);
                PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
                if (pinMap != null && pinGroups.Any())
                {
                    foreach (string pinGroup in pinGroups)
                    {
                        pins.AddRange(pinMap.GetPinsFromGroup(pinGroup).Select(x => x.PinName));
                    }

                    CheckPatternPins(
                        pins,
                        row.PayloadList,
                        pinMap,
                        missingPinsReport,
                        harvestCheckedPattern);
                }
            }
        }

        private static List<string> GetGroupPins(string patternPinGroup)
        {
            if (!string.IsNullOrEmpty(patternPinGroup))
            {
                string[] arr = patternPinGroup.Split(';');
                var pins = new List<string>();
                foreach (string item in arr)
                {
                    string text = item.Trim();
                    if (_regex5.IsMatch(text))
                    {
                        string pin = _regex5.Match(text).Groups["str"].ToString();
                        pins.Add(pin);
                    }
                }
                return pins;
            }
            return new List<string>();
        }

        private static List<string> GetPatternPins(string pattern, List<string> patternPathListInK)
        {
            Dictionary<string, PatternData> dic = AcTSetCategoryMapSingleton.Instance().PatternList;
            string key = pattern.ToLower();
            if (dic.ContainsKey(key))
            {
                string patternVersion = dic[key].PatternVersion + ".pat.gz";
                var fullPaths = patternPathListInK
                    .Where(s => s.EndsWith(patternVersion, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                if (!fullPaths.Any())
                {
                    patternVersion = dic[key].PatternVersion + ".patx.gz";
                    fullPaths = patternPathListInK
                        .Where(s => s.EndsWith(patternVersion, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }

                if (!fullPaths.Any())
                {
                    return null;
                }

                string fullPath = fullPaths.First();
                string pinListContent = "";
                var patInfoCmd = new PatInfoCmd();
                if (patInfoCmd.ConvertByArgs(fullPath, ref pinListContent, "-config"))
                {
                    return patInfoCmd.GetPinList(pinListContent.Split('\n').ToList());
                }
            }

            return null;
        }

        private static void WriteMissingPin(
            ExcelWorksheet sheet,
            string pattern,
            List<string> pins,
            List<string> harvestCheckedPattern
        )
        {
            if (pins.Count == 0)
            {
                return;
            }
            var rowTmp = new List<string> { pattern };
            rowTmp.AddRange(pins);
            sheet.Cells[sheet.Dimension.Rows + 1, 1].PrintExcelRow(rowTmp.ToArray());
            harvestCheckedPattern.Add(pattern);
        }

        internal static string GenGetSubName(string name, string rule)
        {
            var resultList = new List<string>();
            string[] words = name.Split('_');
            string[] numbersStrings = rule.Split(',');
            foreach (string numbers in numbersStrings)
            {
                if (rule.ToLower() == "full")
                {
                    return name;
                }

                if (rule == "")
                {
                    //no operation
                }
                else
                {
                    int getNumber = int.Parse(numbers);
                    if (words.Length > getNumber && getNumber >= 0)
                    {
                        //resultName = resultName + "_" + words[getNumber];
                        resultList.Add(words[getNumber]);

                    }
                    //Error message
                }
            }
            string resultName = string.Join("_", resultList);

            return resultName;
        }
    }
}
