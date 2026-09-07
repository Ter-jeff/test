using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc;

namespace BinCutScriptLib.Comparer
{
    internal class HarvestSourceCodeManager(CheckManager checkManager)
    {
        private readonly CheckManager _checkManager = checkManager;

        public void Check(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, OneStep oneStep, string curInstanceName)
        {
            HarvPmodeTableGroupSheet? harvestPmodeGroup = BinCutData.HarvPmodeTableGroupSheet;
            string curInstanceGroup = "";
            if (curInstanceName.Contains("GROUP", StringComparison.OrdinalIgnoreCase))
            {
                curInstanceGroup = curInstanceName.Split('_').First(x => x.Contains("GROUP", StringComparison.OrdinalIgnoreCase));
            }
            if (oneStep.OneHarvestSourceCodetRows.Count != 0)
            {
                foreach (HarvestSourceCodetRow row in oneStep.OneHarvestSourceCodetRows)
                {
                    CheckRow(streamWriter, siteInfoArray, curInstanceName, harvestPmodeGroup, curInstanceGroup, row);
                }
            }
        }

        private void CheckRow(StreamWriter streamWriter, SiteInfo[] siteInfoArray, string curInstanceName, HarvPmodeTableGroupSheet? harvPmodeTableGroupSheet, string curInstanceGroup, HarvestSourceCodetRow harvestSourceCodetRow)
        {
            Dictionary<string, string> harvesFlags = siteInfoArray[harvestSourceCodetRow.Site].HarvesFlags;
            if (harvesFlags == null || harvesFlags.Count == 0)
            {
                BinCutPrint.PrintHarvFlagNotFind(streamWriter, harvestSourceCodetRow.Line);
            }

            string groupCondition = "";
            string overWriteSequence = "";
            string flagName = !string.IsNullOrEmpty(harvestSourceCodetRow.MainFlag) ? harvestSourceCodetRow.MainFlag.Split('=').First().Trim() : "";
            string deviceCondition = !string.IsNullOrEmpty(harvestSourceCodetRow.MainFlag) ? harvestSourceCodetRow.MainFlag.Split('=').Last().Trim() : "";

            if (harvPmodeTableGroupSheet != null && harvPmodeTableGroupSheet.Rows.Count > 0)
            {
                for (int i = 0; i < harvPmodeTableGroupSheet.Rows.Count; i++)
                {
                    foreach (KeyValuePair<string, string> group in harvPmodeTableGroupSheet.Rows[i].Groups)
                    {
                        string key = harvPmodeTableGroupSheet.Rows[i].ByMode + "_" + group.Key;
                        if (curInstanceName.Contains(key, StringComparison.OrdinalIgnoreCase) && group.Key.EqualsIgnoreCase(curInstanceGroup))
                        {
                            for (int j = 0; j < harvPmodeTableGroupSheet.Rows[j].FailflagName.Count; j++)
                            {
                                if (
                                    harvPmodeTableGroupSheet.Rows[i].FailflagName[j].EqualsIgnoreCase(flagName) &&
                                    harvPmodeTableGroupSheet.Rows[i].DeviceCondition[j].EqualsIgnoreCase(deviceCondition))
                                {
                                    GetGroupCondition(harvPmodeTableGroupSheet.Rows[i].ByMode, group.Key, harvesFlags!, ref groupCondition, ref overWriteSequence);
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(groupCondition))
            {
                CheckByGroupCondition(streamWriter, harvestSourceCodetRow, harvesFlags!, groupCondition, overWriteSequence);
            }
            else if (BinCutConfig.MultiFstpEnable)
            {
                List<bool> totalBits = BinCutData.DigSrcInstructionSheet!.GetBits(harvestSourceCodetRow.PatternName, harvestSourceCodetRow.PatternName, harvesFlags!, streamWriter, harvestSourceCodetRow.Line);
                var bitString = totalBits.Select(totalBit => totalBit ? "1" : "0").ToList();
                string finalBits = string.Join("", bitString);
                string datalog = harvestSourceCodetRow.HarvestSourceCode.Split(':').Last();
                if (datalog != finalBits)
                {
                    BinCutPrint.PrintHarvestSourceCodeError(streamWriter, harvestSourceCodetRow.Line, finalBits);
                }

                _checkManager.CheckHarvestSourceCodeLine(harvestSourceCodetRow.Line);
            }
            else
            {
                BinCutPrint.PrintHarvestSourceCodeError(streamWriter, harvestSourceCodetRow.Line);
                _checkManager.CheckHarvestSourceCodeLine(harvestSourceCodetRow.Line);
            }
        }

        private void CheckByGroupCondition(StreamWriter streamWriter, HarvestSourceCodetRow harvestSourceCodetRow, Dictionary<string, string> harvesFlags, string groupCondition, string overWriteSequence)
        {
            List<string> fstp = BinCutData.HarvPmodeTableFstpSheet!.GetFstp(groupCondition);
            Dictionary<string, string> flags = BinCutData.HarvPmodeTableHarvCheckSheet!.GetFlags(groupCondition);
            var totalBits = new List<bool>();

            for (int index = 0; index < fstp.Count; index++)
            {
                string item = fstp[index];
                KeyValuePair<string, string> flag = flags.ElementAt(index);
                foreach (KeyValuePair<string, string> result in harvesFlags)
                {
                    if (result.Key.EqualsIgnoreCase(flag.Key))
                    {
                        if (result.Value.EqualsIgnoreCase(flag.Value[..1]))
                        {
                            continue;
                        }

                        List<bool> bits = BinCutData.HarvMappingTableSheet!.GetBits(item);
                        totalBits = totalBits.Count == 0 ? bits : ListOrOperator(bits, totalBits);
                    }
                }
            }

            var bitString = totalBits.Select(totalBit => totalBit ? "1" : "0").ToList();
            Dictionary<int, int> sequenceList = BinCutData.HarvPmodeTableGroupSheet!.SequenceList;
            for (int i = 0; i < sequenceList.Count; i++)
            {
                if (sequenceList.ElementAt(i).Key > sequenceList.ElementAt(i).Value)
                {
                    throw new Exception("Autogen : Start and End position for Sequence in groupCondition is Invalid !!!");
                }

                string[] spilt = overWriteSequence.Split(':');
                spilt[i] = Reg.RegexB.Replace(spilt[i].TrimStart(' '), "");
                if (!string.IsNullOrEmpty(spilt[i]))
                {
                    for (int j = 0; j < sequenceList.ElementAt(i).Value - sequenceList.ElementAt(i).Key + 1; j++)
                    {
                        bitString[j + sequenceList.ElementAt(i).Key] = spilt[i].Substring(j, 1);
                    }
                }
            }
            string finalBits = string.Join("", bitString);
            string datalog = harvestSourceCodetRow.HarvestSourceCode.Split(':').Last();
            if (datalog != finalBits)
            {
                BinCutPrint.PrintHarvestSourceCodeError(streamWriter, harvestSourceCodetRow.Line, finalBits);
            }

            _checkManager.CheckHarvestSourceCodeLine(harvestSourceCodetRow.Line);
        }

        public bool CheckHarvestSourceCodeCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<HarvestSourceCodetRow> harvestSourceCodetRows, bool isSelSram = false)
        {
            bool flag = false;
            if (harvestSourceCodetRows.Count != 0)
            {
                foreach (HarvestSourceCodetRow row in harvestSourceCodetRows)
                {
                    Dictionary<string, string> harvesFlags = siteInfoArray[row.Site].HarvesFlags;
                    if (harvesFlags == null || harvesFlags.Count == 0)
                    {
                        BinCutPrint.PrintHarvFlagNotFind(streamWriter, row.Line);
                    }
                    else
                    {
                        if (BinCutData.DigSrcInstructionSheet != null)
                        {
                            List<bool> totalBits = BinCutData.DigSrcInstructionSheet.GetBits(row.PatternName, row.Assignment, harvesFlags, streamWriter, row.Line);
                            var bitString = totalBits.Select(totalBit => totalBit ? "1" : "0").ToList();
                            string finalBits = string.Join("", bitString);
                            string datalog = row.HarvestSourceCode.Split(':').Last();
                            _checkManager.CheckHarvestSourceCodeLine(row.Line);
                            if (datalog != finalBits)
                            {
                                if (!isSelSram || datalog.Length == finalBits.Length)
                                {
                                    int secondBracketIndex = row.Line.Line.IndexOf(']', row.Line.Line.IndexOf(']') + 1);

                                    if (secondBracketIndex != -1)
                                    {
                                        string expectedLine = string.Concat(row.Line.Line.AsSpan(0, secondBracketIndex + 1), " ", finalBits);
                                        BinCutPrint.PrintHarvestSourceCodeErrorCs(streamWriter, row.Line, expectedLine);
                                        flag = false;
                                    }
                                }
                                else
                                {
                                    flag = true;
                                }
                            }
                        }
                        continue;
                    }
                }
            }
            return flag;
        }

        private static List<bool> ListOrOperator(List<bool> a, List<bool> b)
        {
            if (a.Count == b.Count)
            {
                for (int i = 0; i < a.Count; i++)
                {
                    a[i] |= b[i];
                }
            }

            return a;
        }

        private static void GetGroupCondition(string byMode, string group, Dictionary<string, string> harvesFlags, ref string groupCondition, ref string overWriteSequence)
        {
            groupCondition = GetGroupCondition(byMode, harvesFlags, group);
            HarvPmodeTableGroupRow? harvPmodeTableGroupRow = GetHarvPmodeTableGroupRow(byMode, harvesFlags);
            if (harvPmodeTableGroupRow != null)
            {
                overWriteSequence = harvPmodeTableGroupRow.Sequence;
            }
        }

        private static HarvPmodeTableGroupRow? GetHarvPmodeTableGroupRow(string byMode, Dictionary<string, string> harvesFlags)
        {
            List<string> mainFlags = BinCutData.HarvPmodeTableGroupSheet!.GetMainFlag(byMode) ?? [];
            var deviceConditions = new List<string>();
            int getFlags = 0;
            foreach (KeyValuePair<string, string> item in harvesFlags)
            {
                foreach (string mainFlag in mainFlags)
                {
                    if (item.Key.EqualsIgnoreCase(mainFlag))
                    {
                        string deviceCondition = "TRUE";
                        if (item.Value == "F")
                        {
                            deviceCondition = "FALSE";
                        }

                        getFlags++;
                        deviceConditions.Add(deviceCondition);
                    }
                }
                if (getFlags == mainFlags.Count)
                {
                    return BinCutData.HarvPmodeTableGroupSheet.GetRow(byMode, deviceConditions);
                }
            }

            return null;
        }

        private static string GetGroupCondition(string byMode, Dictionary<string, string> harvesFlags, string group)
        {
            List<string> mainFlags = BinCutData.HarvPmodeTableGroupSheet!.GetMainFlag(byMode) ?? [];
            var flags = new List<string>();
            int getFlags = 0;
            foreach (KeyValuePair<string, string> item in harvesFlags)
            {
                foreach (string mainFlag in mainFlags)
                {
                    if (item.Key.EqualsIgnoreCase(mainFlag))
                    {
                        string flag = "TRUE";
                        if (item.Value == "F")
                        {
                            flag = "FALSE";
                        }

                        getFlags++;
                        flags.Add(flag);
                    }
                }
                if (getFlags == mainFlags.Count)
                {
                    return BinCutData.HarvPmodeTableGroupSheet.GetCondition(byMode, flags, group);
                }
            }
            return "";
        }
    }
}
