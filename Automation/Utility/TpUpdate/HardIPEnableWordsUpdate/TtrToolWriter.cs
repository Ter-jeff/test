using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace Automation.Utility.TpUpdate.HardIPEnableWordsUpdate
{
    public class TtrToolWriter
    {

        private readonly HashSet<string> _lvBinCutVoltageSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "BV",
            "BV_chk",
            "LV_PBCut"
        };
        private readonly HashSet<string> _hvBinCutVoltageSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "HBV",
            "HBV_chk",
            "HV_PBCut"
        };

        public SubFlowSheet UpdateOverlayFlowSheetEnableWords(SubFlowSheet sheet, EnableWordTable table, List<string> enableList, Dictionary<string, string> subflowStatus, ref List<string> nonUsedItems, List<string> tpJobs = null)
        {
            subflowStatus[sheet.Name] = "SubFlow no change";
            var flowSheet = new SubFlowSheet(sheet.Name, sheet.SourceInfo.Name);
            var forFlows = new List<FlowRow>();
            var relayFlows = new List<FlowRow>();

            foreach (FlowRow row in sheet.Rows)
            {
                bool oriEnable = !string.IsNullOrEmpty(row.Enable);
                if (row.Opcode.Equals("Test", StringComparison.OrdinalIgnoreCase) || row.Opcode.Equals("characterize", StringComparison.OrdinalIgnoreCase))
                {
                    string originEnable = row.Enable;
                    row.Enable = ClearCurrentEnableWords(row.Enable, enableList);
                    UpdateEnableWordsForOverlayFlows(row, table, oriEnable, tpJobs);

                    if (!string.Equals(originEnable, row.Enable, StringComparison.Ordinal))
                    {
                        subflowStatus[sheet.Name] = "SubFlow Edited";
                    }

                    if (forFlows.Count > 0)
                    {
                        foreach (FlowRow forFlow in forFlows)
                        {
                            forFlow.Enable = row.Enable;
                        }
                        forFlows.Clear();
                    }
                }
                if (row.Opcode.Equals("if", StringComparison.OrdinalIgnoreCase))
                {
                    row.Enable = "";
                }

                if (row.Opcode.Equals("for", StringComparison.OrdinalIgnoreCase))
                {
                    forFlows.Add(row);
                }

                if (row.Opcode.Equals("next", StringComparison.OrdinalIgnoreCase))
                {
                    forFlows.Clear();
                }

                RelayControlEnableFlow(row, relayFlows, out bool isEmptyRelaySet);
                if (isEmptyRelaySet)
                {
                    flowSheet.Rows.RemoveAt(flowSheet.Rows.Count - 1);
                }
                else
                {
                    flowSheet.AddRow(row);
                }
            }

            return flowSheet;
        }

        public (SubFlowSheet flowSheet, List<FailControlData> failControls) UpdateFlowSheetEnableWords(SubFlowSheet sheet, EnableWordTable table, List<string> enableList, Dictionary<string, string> subflowStatus, ref List<string> nonUsedItems, List<string> tpJobs = null)
        {
            subflowStatus[sheet.Name] = "SubFlow no change";
            var flowSheet = new SubFlowSheet(sheet.Name, sheet.SourceInfo.Name);
            var forFlows = new List<FlowRow>();
            var relayFlows = new List<FlowRow>();

            var parameterToItemCache = new Dictionary<string, EnableWordTableRow>(StringComparer.OrdinalIgnoreCase);
            var usedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var failControls = new List<FailControlData>();

            foreach (FlowRow row in sheet.Rows)
            {
                string opcode = row.Opcode;
                bool isTestOrChar = string.Equals(opcode, "Test", StringComparison.OrdinalIgnoreCase) || string.Equals(opcode, "characterize", StringComparison.OrdinalIgnoreCase);

                if (isTestOrChar)
                {
                    if (!parameterToItemCache.TryGetValue(row.Parameter, out EnableWordTableRow tableItem))
                    {
                        tableItem = table.Rows.FirstOrDefault(p => row.Parameter.ContainsIgnoreCase($"_{p.SubBlock}_{p.Name}")) ??
                            table.Rows.Where(p => p.IsOther).ToList().FirstOrDefault(p => row.Parameter.ContainsIgnoreCase(p.Name));

                        parameterToItemCache[row.Parameter] = tableItem;
                    }

                    if (tableItem != null)
                    {
                        string originalEnable = row.Enable;
                        row.Enable = ClearCurrentEnableWords(row.Enable, enableList);

                        failControls.AddRange(JudgeFailControl(sheet.Rows, row, tableItem));
                        UpdateEnableWords(row, tableItem, table, !string.IsNullOrEmpty(originalEnable), tpJobs);

                        if (!string.Equals(originalEnable, row.Enable, StringComparison.Ordinal))
                        {
                            subflowStatus[sheet.Name] = "SubFlow Edited";
                        }

                        if (forFlows.Count > 0)
                        {
                            foreach (FlowRow forFlow in forFlows)
                            {
                                forFlow.Enable = row.Enable;
                            }

                            forFlows.Clear();
                        }
                        usedKeys.Add($"{tableItem.SubBlock}_{tableItem.Name}");
                    }
                }
                else if (string.Equals(opcode, "if", StringComparison.OrdinalIgnoreCase))
                {
                    row.Enable = string.Empty;
                }
                else if (string.Equals(opcode, "for", StringComparison.OrdinalIgnoreCase))
                {
                    forFlows.Add(row);
                }
                else if (string.Equals(opcode, "next", StringComparison.OrdinalIgnoreCase))
                {
                    forFlows.Clear();
                }

                RelayControlEnableFlow(row, relayFlows, out bool isEmptyRelaySet);
                if (isEmptyRelaySet)
                {
                    flowSheet.Rows.RemoveAt(flowSheet.Rows.Count - 1);
                }
                else
                {
                    flowSheet.AddRow(row);
                }
            }

            nonUsedItems = table.Rows.Select(r => $"{r.SubBlock}_{r.Name}").Where(k => !usedKeys.Contains(k)).ToList();

            return (flowSheet, failControls);
        }

        public InstanceSheet UpdateInsSheetOverlays(InstanceSheet sheet, EnableWordTable table, Dictionary<string, string> binCutOverlayName)
        {
            foreach (InstanceRow instance in sheet.Rows)
            {
                if (!instance.TestName.EndsWith("_LV", StringComparison.OrdinalIgnoreCase) && !instance.TestName.EndsWith("_HV", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string voltageType = instance.TestName.Split('_').Last().ToUpper();

                EnableWordTableRow ttrItem = table.Rows.FirstOrDefault(p => instance.TestName.ContainsIgnoreCase($"_{p.SubBlock}_{p.Name}")) ??
                            table.Rows.Where(p => p.IsOther).ToList().FirstOrDefault(p => instance.TestName.ContainsIgnoreCase(p.Name));

                if (ttrItem != null)
                {
                    if (!string.IsNullOrEmpty(ttrItem.EnableWords.Keys.FirstOrDefault(x => x.Equals(voltageType.Equals("LV") ? "LV_PBCut" : "HV_PBCut", StringComparison.OrdinalIgnoreCase))))
                    {
                        if (ttrItem.Name.Split('_').Length > 9)
                        {
                            string mode = ttrItem.Name.Split('_')[9];
                            KeyValuePair<string, string> targetOverlay = binCutOverlayName.FirstOrDefault(x => x.Key.ToUpper() == (voltageType.Equals("LV") ? mode.ToUpper().Replace("X", "") : "HBV_" + mode.ToUpper().Replace("X", "")));
                            if (!string.IsNullOrEmpty(targetOverlay.Key))
                            {
                                instance.Overlay = targetOverlay.Value;
                            }
                        }
                    }
                }
            }
            return sheet;
        }

        private void RelayControlEnableFlow(FlowRow row, List<FlowRow> relayFlows, out bool isEmptyRelaySet)
        {
            isEmptyRelaySet = false;

            var firstRelayOn = new List<string>();
            var firstRelayOff = new List<string>();
            var lastRelayOn = new List<string>();
            var lastRelayOff = new List<string>();
            bool isRelayItem = row.Parameter.StartsWith(HardIpConstData.PrefixAtgRelay) &&
                (row.Parameter.ContainsIgnoreCase("relayon") || row.Parameter.ContainsIgnoreCase("relayoff"));

            //if opcode is return, clear setting and return
            if (row.Parameter.Equals("return"))
            {
                relayFlows.Clear();
                return;
            }
            //start to record relay when relay on
            if (relayFlows.Count == 0 && isRelayItem)
            {
                relayFlows.Add(row);
                return;
            }
            //if exist relay events....
            if (relayFlows.Count > 0 && row.Opcode.Equals("test", StringComparison.OrdinalIgnoreCase))
            {
                //if not trigger relay event => just record
                if (!isRelayItem)
                {
                    relayFlows.Add(row);
                }
                // Trigger summary relay event, 
                else
                {
                    relayFlows.Add(row);
                    JudgeRelaySetting(relayFlows.First().Parameter, firstRelayOn, firstRelayOff);
                    JudgeRelaySetting(relayFlows.Last().Parameter, lastRelayOn, lastRelayOff);
                    //if switch on/off between test items are fully same => 

                    if (firstRelayOn.SequenceEqual(lastRelayOff) && firstRelayOff.SequenceEqual(lastRelayOn))
                    {
                        if (relayFlows.Count == 2)
                        {
                            isEmptyRelaySet = true;
                        }
                        else if (relayFlows.Count == 3)
                        {
                            relayFlows.First().Enable = relayFlows[1].Enable;
                            relayFlows.Last().Enable = relayFlows[1].Enable;
                        }
                        else if (relayFlows.Count > 3)
                        {
                            string finalEnableWd = BuildMergedRelayEnableWord(relayFlows);
                            relayFlows.First().Enable = finalEnableWd;
                            relayFlows.Last().Enable = finalEnableWd;
                        }
                    }
                    else
                    {
                        relayFlows.First().Enable = "";
                        relayFlows.Last().Enable = "";
                    }
                    relayFlows.Clear();
                }
            }
        }

        private string BuildMergedRelayEnableWord(List<FlowRow> relayFlows)
        {
            HashSet<string> enableWdWithOrOr = new HashSet<string>();
            HashSet<string> enableWdWithAndAnd = new HashSet<string>();

            for (int i = 1; i < relayFlows.Count - 2; i++)
            {
                string relayFlowEnable = relayFlows[i].Enable;
                int idxAndAnd = relayFlowEnable.IndexOf("&&", StringComparison.Ordinal);

                if (idxAndAnd >= 0)
                {
                    string beforeAndAnd = relayFlowEnable.Substring(0, idxAndAnd);
                    string afterAndAnd = relayFlowEnable.Substring(idxAndAnd + 2);

                    enableWdWithOrOr.AddRange(beforeAndAnd.Split(new[] { "||", "(", ")" },
                        StringSplitOptions.RemoveEmptyEntries));
                    enableWdWithAndAnd.AddRange(afterAndAnd.Split(new[] { "&&" },
                        StringSplitOptions.RemoveEmptyEntries));
                }
                else
                {
                    enableWdWithOrOr.AddRange(relayFlowEnable.Split(new[] { "||", "(", ")" },
                        StringSplitOptions.RemoveEmptyEntries));
                }
            }

            string finalEnableWd = string.Empty;
            if (enableWdWithOrOr.Any())
            {
                finalEnableWd = $"({string.Join("||", enableWdWithOrOr)})";
            }

            if (enableWdWithAndAnd.Any())
            {
                finalEnableWd = string.IsNullOrEmpty(finalEnableWd)
                    ? string.Join("&&", enableWdWithAndAnd)
                    : $"{finalEnableWd}&&{string.Join("&&", enableWdWithAndAnd)}";
            }

            if (finalEnableWd.StartsWith("(") &&
                finalEnableWd.EndsWith(")") &&
                !finalEnableWd.Contains("||") &&
                !finalEnableWd.Contains("&&"))
            {
                finalEnableWd = finalEnableWd[1..^1];
            }

            return finalEnableWd;
        }
        private void UpdateEnableWordsForOverlayFlows(FlowRow row, EnableWordTable table, bool oriEnable, List<string> tpJobs)
        {
            string item;
            string regPmode = @"_(?<pmode>M([a-z]|[A-Z]+)\d+([a-z]|[A-Z]|[0-9]+))";
            item = Regex.Match(row.Parameter, regPmode, RegexOptions.IgnoreCase).Groups["pmode"].ToString();

            if (item != null)
            {
                var enables = new List<string>();
                if (row.Enable != "")
                {
                    row.Enable = ClearCurrentEnableWords(row.Enable, table.EnableWords);
                }

                var testEnables = new List<string>();
                if (item != "" && table.OverlayInfoDic.ContainsKey(item.ToUpper()))
                {
                    if (table.OverlayInfoDic.TryGetValue(item.ToUpper(), out List<string> overlayEnables))
                    {
                        testEnables.AddRange(overlayEnables);
                    }
                }
                List<string> list = new List<string>();
                if (tpJobs != null)
                {
                    var updatedJobs = new List<string>();
                    List<string> allJobEnables = new List<string>(testEnables);
                    allJobEnables.AddRange(list);
                    foreach (string enable in allJobEnables)
                    {
                        IEnumerable<string> ttrJobs = tpJobs.Where(x => Regex.IsMatch(enable, x, RegexOptions.IgnoreCase)).Select(x => x.ToUpper());
                        updatedJobs.AddRange(ttrJobs);
                    }
                    updatedJobs = updatedJobs.Distinct().ToList();
                    updatedJobs.Sort();

                    if (!updatedJobs.Any())
                    {
                        row.Job = string.Join(",", tpJobs.Select(x => $"!{x}"));
                    }
                    else
                    {
                        row.Job = string.Join(",", updatedJobs);
                    }
                }
                if (testEnables.Any())
                {
                    if (oriEnable)
                    {
                        enables.AddRange(MergeEnableWord(row.Enable, testEnables, table.EnableWords));
                    }
                    row.Enable = RemoveRedundantProp(string.Join("||", enables));
                }
            }
            if (Regex.IsMatch(row.Enable, @"^\(\w+\)$", RegexOptions.IgnoreCase))
            {
                row.Enable = row.Enable.Trim('(').Trim(')');
            }
        }

        private void UpdateEnableWords(FlowRow row, EnableWordTableRow info, EnableWordTable table, bool oriEnable, List<string> tpJobs)
        {
            string item;
            if (!info.IsOther)
            {
                string regVol = "_(?<vol>[HLN]V$)";
                item = Regex.Match(row.Parameter, regVol, RegexOptions.IgnoreCase).Groups["vol"].ToString();
            }
            else
            {
                string regVol = "_*(?<vol>[HLN]V)";
                item = Regex.Match(row.Enable, regVol, RegexOptions.IgnoreCase).Groups["vol"].ToString();
            }
            if (item != "" || (item == "" && info.IsOther))
            {
                var enables = new List<string>();
                if (row.Enable != "")
                {
                    row.Enable = ClearCurrentEnableWords(row.Enable, info.TotalEnableWords);
                }
                var testEnables = new List<string>();

                if (item != "" && (info.EnableWords.ContainsKey(item) || info.EnableWords.ContainsKey($"{item}_PBCut"))) //Select from NV/LV/HV
                {
                    if (info.EnableWords.TryGetValue(item, out List<string> voltageEnables))
                    {
                        testEnables.AddRange(voltageEnables);
                    }
                    if (info.EnableWords.TryGetValue($"{item}_PBCut", out List<string> pbCutEnables))
                    {
                        testEnables.AddRange(pbCutEnables);
                    }
                }
                else if (item == "" && info.IsOther) //Combine all enable words for other type instance without voltage
                {
                    testEnables = info.EnableWords.SelectMany(x => x.Value).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                }
                List<string> list;
                (testEnables, list) = AppendBvUsingEnableWord(item, info, testEnables);

                if (tpJobs != null)
                {
                    var updatedJobs = new List<string>();
                    List<string> allJobEnables = new List<string>(testEnables);
                    allJobEnables.AddRange(list);
                    foreach (string enable in allJobEnables)
                    {
                        IEnumerable<string> ttrJobs = tpJobs.Where(x => Regex.IsMatch(enable, x, RegexOptions.IgnoreCase)).Select(x => x.ToUpper());
                        updatedJobs.AddRange(ttrJobs);
                    }
                    updatedJobs = updatedJobs.Distinct().ToList();
                    updatedJobs.Sort();

                    if (testEnables.Any(x => x.Equals("MON", StringComparison.OrdinalIgnoreCase)))
                    {
                        string[] supportedJobs = { "CP1", "CP2" };

                        row.Job = string.Join(",",
                            supportedJobs.Where(job =>
                                tpJobs.Contains(job, StringComparer.OrdinalIgnoreCase)));
                    }
                    else if (!updatedJobs.Any())
                    {
                        row.Job = string.Join(",", tpJobs.Select(x => $"!{x}"));
                    }
                    else
                    {
                        row.Job = string.Join(",", updatedJobs);
                    }


                }

                if (testEnables.Any())
                {
                    if (oriEnable)
                    {
                        enables.AddRange(MergeEnableWord(row.Enable, testEnables, table.EnableWords));
                    }
                    row.Enable = RemoveRedundantProp(string.Join("||", enables));
                }
            }
            if (Regex.IsMatch(row.Enable, @"^\(\w+\)$", RegexOptions.IgnoreCase))
            {
                row.Enable = row.Enable.Trim('(').Trim(')');
            }
        }

        private (List<string>, List<string>) AppendBvUsingEnableWord(string voltage, EnableWordTableRow info, List<string> testEnables)
        {
            bool bvUsing = false;
            var bvJobEnableWord = new List<string>();
            if (voltage.ToUpper() == "LV")
            {
                foreach (string binCutVoltage in _lvBinCutVoltageSet)
                {
                    if (info.EnableWords.TryGetValue(binCutVoltage, out List<string> word))
                    {
                        bvJobEnableWord.AddRange(word);
                        if (!bvUsing)
                        {
                            bvUsing = true;
                        }
                    }
                }
                bvJobEnableWord = bvJobEnableWord.Distinct().ToList();
            }
            if (voltage.ToUpper() == "HV")
            {
                foreach (string binCutVoltage in _hvBinCutVoltageSet)
                {
                    if (info.EnableWords.TryGetValue(binCutVoltage, out List<string> word))
                    {
                        bvJobEnableWord.AddRange(word);
                        if (!bvUsing)
                        {
                            bvUsing = true;
                        }
                    }
                }
                bvJobEnableWord = bvJobEnableWord.Distinct().ToList();
            }
            if (bvUsing)
            {
                testEnables.Add("BV_Validation");
            }

            return (testEnables, bvJobEnableWord);
        }

        private List<FailControlData> JudgeFailControl(List<FlowRow> flowList, FlowRow flowRow, EnableWordTableRow ttrInfo)
        {
            var failControls = new List<FailControlData>();
            string ttrVoltage;
            if (!ttrInfo.IsOther)
            {
                string regVol = "_(?<vol>[HLN]V$)";
                ttrVoltage = Regex.Match(flowRow.Parameter, regVol, RegexOptions.IgnoreCase).Groups["vol"].ToString();
            }
            else
            {
                string regVol = "_*(?<vol>[HLN]V)";
                ttrVoltage = Regex.Match(flowRow.Enable, regVol, RegexOptions.IgnoreCase).Groups["vol"].ToString();
            }

            if (!string.IsNullOrEmpty(ttrVoltage) && !string.IsNullOrEmpty(ttrInfo.FailControl) && ttrInfo.EnableWords.ContainsKey(ttrVoltage))
            {
                bool isPerTD = false;
                string failControl = ttrInfo.FailControl;

                if (ttrInfo.FailControl.Contains(','))
                {
                    string[] parts = ttrInfo.FailControl.Split(',', 2);

                    failControl = parts[0].Trim();
                    isPerTD = string.Equals(
                        parts[1].Trim(),
                        "PerTD",
                        StringComparison.OrdinalIgnoreCase);
                }

                FlowRow previousTest = flowList.FirstOrDefault(x => x.Opcode.Equals("Test", StringComparison.OrdinalIgnoreCase)
                    && x.Parameter.ContainsIgnoreCase("_" + failControl.ToUpper())
                    && x.Parameter.ContainsIgnoreCase("_" + ttrVoltage.ToUpper()));

                if (previousTest == null)
                {
                    return failControls;
                }

                ttrInfo.EnableWords[ttrVoltage].Add("HIP_FailRun");
                var failControlData = new FailControlData
                {
                    InstanceName = string.Format($"{ttrInfo.Category.Replace("_", "")}_{ttrInfo.SubBlock}_{ttrInfo.Name}_{ttrVoltage}"),
                    EnableAllSite = isPerTD ? "TRUE" : "FALSE"
                };
                if (string.IsNullOrEmpty(previousTest.FailAction))
                {
                    string[] splitComment = previousTest.Comment1.Split('\t');
                    if (splitComment.Length > 1)
                    {
                        failControlData.FailFlagName = splitComment[1];
                    }
                }
                else
                {
                    failControlData.FailFlagName = previousTest.FailAction;
                }
                failControls.Add(failControlData);
            }

            return failControls;
        }

        public string WriteFailControlTable(string outFolder, List<FailControlData> failControls)
        {
            if (failControls.Any())
            {
                var excelPck = new ExcelPackage(new FileInfo(Path.Combine(outFolder, "HardIP_FailRun.xlsx")));
                ExcelWorksheet workSheet = excelPck.Workbook.AddSheet("HardIP_FailRun");
                var headers = new List<string> { "InstanceName", "FailFlagName", "EnableAllSite" };
                int col = 1;
                foreach (string header in headers)
                {
                    workSheet.Cells[1, col].Value = header;
                    col++;
                }

                int row = 2;
                var sorted = failControls.OrderBy(x => x.InstanceName).ToList();
                foreach (FailControlData failControlDate in sorted)
                {
                    workSheet.Cells[row, 1].Value = failControlDate.InstanceName;
                    workSheet.Cells[row, 2].Value = failControlDate.FailFlagName;
                    workSheet.Cells[row, 3].Value = failControlDate.EnableAllSite;
                    row++;
                }
                excelPck.ExportWorkBook2Txt(outFolder);
                return "HardIP_FailRun";
            }
            return "";
        }

        private string ClearCurrentEnableWords(string flowRowEnables, List<string> tableEnables)
        {
            string regEnable = @"(?<prefix>\!)?(?<Enable>\w+)";
            string regCombine = @"(?<interp>(\|\|)|(\&\&))?(?<prefix>\!)?(?<Enable>\w+)?";
            MatchCollection allItems = Regex.Matches(flowRowEnables, regCombine);
            char[] modified = flowRowEnables.ToArray();
            foreach (Match item in allItems)
            {
                string target = item.ToString().Replace("||", "").Replace("&&", "");
                string enable = Regex.Match(target, regEnable).Groups["Enable"].ToString();
                if (tableEnables.Contains(enable))
                {
                    for (int i = item.Index; i < item.Index + item.Length; i++)
                    {
                        modified[i] = ' ';
                    }
                }
            }
            string modifiedFlowRowEnables = new string(modified).Replace(" ", "");
            return RemoveRedundantProp(modifiedFlowRowEnables);
        }

        private string RemoveRedundantProp(string enable)
        {
            string regRemove = @"(^\&\&)?(\&\&$)?(^\|\|)?(\|\|$)?";
            return Regex.Replace(enable, regRemove, "", RegexOptions.IgnoreCase);
        }

        private string MergeEnableWordsWithVoltage(string rowEnable, List<string> items)
        {
            string regFlagHln = @"\|*\&*\s*(?<HardIP>HardIP_[HLN]V\w*)\s*\|*\&*";
            string str;
            if (Regex.IsMatch(rowEnable, regFlagHln, RegexOptions.IgnoreCase))
            {
                string matchItem = Regex.Match(rowEnable, regFlagHln, RegexOptions.IgnoreCase).Groups["HardIP"].ToString();
                string replaceItem;
                if (Regex.IsMatch(rowEnable, $"\\({matchItem}\\)"))
                {
                    replaceItem = $"{matchItem}||{string.Join("||", items)}";
                }
                else
                {
                    replaceItem = $"({matchItem}||{string.Join("||", items)})";
                }

                str = Regex.Replace(rowEnable, matchItem, replaceItem);
            }
            else
            {
                if (rowEnable != "")
                {
                    str = $"({rowEnable}||{string.Join("||", items)})";
                }
                else if (items.Count == 1)
                {
                    str = items[0];
                }
                else
                {
                    str = $"({string.Join("||", items)})";
                }
            }

            return str;
        }

        private List<string> MergeEnableWord(string rowEnable, List<string> testEnables, List<string> allEnableWords)
        {
            string regStr = @"(?<Enable>[a-zA-Z]+_*)\s*(?<Job>\w*(cp)|(ft)|(wlft)\d*)";
            List<string> itemsTotal = allEnableWords.FindAll(p => Regex.IsMatch(p, regStr, RegexOptions.IgnoreCase));
            //Use Job to Search all enable words
            var enableList = new List<string>();
            foreach (string item in itemsTotal)
            {
                string enableItem = Regex.Match(item, regStr, RegexOptions.IgnoreCase).Groups["Enable"].ToString();
                if (!enableList.Contains(enableItem))
                {
                    enableList.Add(enableItem);
                }
            }
            //Use Enable Words to search all jobs with specified Enable Word
            foreach (string enable in enableList)
            {
                regStr = $@"{enable}_*\s*(?<Job>\w*(cp)|(ft)|(wlft)\d*)";
                itemsTotal = allEnableWords.FindAll(p => Regex.IsMatch(p, regStr, RegexOptions.IgnoreCase));
                bool isNeedToMerge = true;
                foreach (string itemToCheck in itemsTotal)
                {
                    if (testEnables.Contains(itemToCheck))
                    {
                        continue;
                    }

                    isNeedToMerge = false;
                    break;
                }
                if (isNeedToMerge)
                {
                    foreach (string itemSel in itemsTotal)
                    {
                        testEnables.Remove(itemSel);
                    }
                    testEnables.Add(enable + "All");
                }
            }

            return new List<string> { MergeEnableWordsWithVoltage(rowEnable, testEnables) };
        }

        private void JudgeRelaySetting(string relaySetStr, List<string> relayOn, List<string> relayOff)
        {
            //relayon_K1_K2_K3_relayoff_K5_K6_K15
            //relayon_K1_K2_K3
            //relayoff_K5_K6_K15
            string regOff = @"relayoff_(?<set>\w+)";
            string regOn = @"relayon_(?<set>\w+)";
            if (Regex.IsMatch(relaySetStr, regOff, RegexOptions.IgnoreCase))
            {
                string set = Regex.Match(relaySetStr, regOff, RegexOptions.IgnoreCase).Groups["set"].Value.Trim('_');
                relayOff.AddRange(set.Split('_'));
                relaySetStr = Regex.Replace(relaySetStr, regOff, "", RegexOptions.IgnoreCase);
            }
            if (Regex.IsMatch(relaySetStr, regOn, RegexOptions.IgnoreCase))
            {
                string set = Regex.Match(relaySetStr, regOn, RegexOptions.IgnoreCase).Groups["set"].Value.Trim('_');
                relayOn.AddRange(set.Split('_'));
            }
        }
    }
}
