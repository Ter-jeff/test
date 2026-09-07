using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using CommonLib.Extension;

using IgxlLib.Enums;

using TagDiff.Core.Common;
using TagDiff.Core.Input;
using TagDiff.Core.Output;
using TagDiff.Core.Utility;

namespace TagDiff.Core.Comparer
{
    internal class FlowSheetComparer : IgxlSheetComparerBase
    {
        private readonly string _job;

        public FlowSheetComparer(string sheetName, string baseFile, string compFile, string job) : base(sheetName, baseFile, compFile, "", [])
        {
            _job = job;
            SheetType = EnumSheetType.DTFlowtableSheet;
        }

        public override List<CompareReport> Compare()
        {
            var result1 = new CompareReport();
            if (CompSheet.Count == 0)
            {
                result1.SheetName = SheetName;
                result1.SheetType = SheetType;
                result1.TotalCount = BaseSheet.Count;
                result1.ManualCount = BaseSheet.Count;
                result1.SemiAutoCount = 0;
                return [result1];
            }

            List<List<FlowRowCompare>> groupOri = FlowReader.ReadTxt(BaseFile, "base_sheet", _job);
            List<List<FlowRowCompare>> groupCompare = FlowReader.ReadTxt(CompFile, "compare", _job);
            (List<List<FlowRowCompare>> result, double accuracy, CompareReport countRow) result = FindDifferences(groupOri, groupCompare);
            var resultDict = new ConcurrentDictionary<string, (List<List<FlowRowCompare>> result, double accuracy, CompareReport countRow)> { [SheetName] = result };
            WriteToTxt(resultDict.ToDictionary(kvp => kvp.Key, kvp => kvp.Value), BaseFile);
            return [result.countRow];
        }

        public (List<List<FlowRowCompare>> result, double accuracy, CompareReport countRow) FindDifferences(List<List<FlowRowCompare>> groupedA, List<List<FlowRowCompare>> groupedB)
        {
            var result = new List<List<FlowRowCompare>>();
            Dictionary<string, List<List<FlowRowCompare>>> baseGroups = GroupFlowRows(groupedA);
            Dictionary<string, List<List<FlowRowCompare>>> comparedGroups = GroupFlowRows(groupedB);
            int passCount = 0;
            int diffCount = 0;
            int missingComparedCount = 0;
            foreach (KeyValuePair<string, List<List<FlowRowCompare>>> baseGroupPair in baseGroups)
            {
                if (!comparedGroups.TryGetValue(baseGroupPair.Key, out List<List<FlowRowCompare>>? comparedGroup))
                {
                    foreach (FlowRowCompare item in baseGroupPair.Value.SelectMany(x => x))
                    {
                        item.IsChecked = true;
                        item.Status += "MissingCompared";
                        missingComparedCount++;
                    }
                    result.AddRange(baseGroupPair.Value);
                }
                else
                {
                    List<List<FlowRowCompare>> baseGroup = baseGroupPair.Value;
                    if (baseGroup.Count > comparedGroup.Count)
                    {
                        comparedGroup.AddRange(Enumerable.Repeat(new List<FlowRowCompare>(), baseGroup.Count - comparedGroup.Count));
                    }
                    else if (baseGroup.Count < comparedGroup.Count)
                    {
                        baseGroup.AddRange(Enumerable.Repeat(new List<FlowRowCompare>(), comparedGroup.Count - baseGroup.Count));
                    }

                    for (int i = 0; i < baseGroup.Count; i++)
                    {
                        List<FlowRowCompare> baseRows = baseGroup[i];
                        List<FlowRowCompare> comparedRows = comparedGroup[i];
                        var resultRows = new List<FlowRowCompare>();
                        if (baseRows.Count > comparedRows.Count)
                        {
                            comparedRows.AddRange(Enumerable.Repeat(new FlowRowCompare(), baseRows.Count - comparedRows.Count));
                        }
                        else if (baseRows.Count < comparedRows.Count)
                        {
                            baseRows.AddRange(Enumerable.Repeat(new FlowRowCompare(), comparedRows.Count - baseRows.Count));
                        }

                        foreach ((FlowRowCompare baseRow, FlowRowCompare comparedRow) in baseRows.Zip(comparedRows, (a, b) => (a, b)))
                        {
                            var resultRow = new FlowRowCompare();
                            bool allPass = true;
                            resultRow.Status = "P";
                            baseRow.IsChecked = true;
                            comparedRow.IsChecked = true;

                            resultRow.BaseSheet = baseRow.BaseSheet;
                            if (string.IsNullOrEmpty(baseRow.BaseSheet))
                            {
                                resultRow.Status = "MissingBase";
                            }

                            resultRow.BaseRow = baseRow.BaseRow;
                            resultRow.ComparedSheet = comparedRow.ComparedSheet;
                            if (string.IsNullOrEmpty(comparedRow.ComparedSheet))
                            {
                                resultRow.Status = "MissingCompared";
                                missingComparedCount++;
                            }

                            resultRow.ComparedRow = comparedRow.ComparedRow;

                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.Label = v, baseRow.Label, comparedRow.Label);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.Enable = v, baseRow.Enable, comparedRow.Enable);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.Job = v, baseRow.Job.Replace(" ", ""), comparedRow.Job.Replace(" ", ""));
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.Part = v, baseRow.Part, comparedRow.Part);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.Env = v, baseRow.Env, comparedRow.Env);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.Opcode = v, baseRow.Opcode, comparedRow.Opcode);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.Parameter = v, baseRow.Parameter, comparedRow.Parameter);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.TName = v, baseRow.TName, comparedRow.TName);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.TNum = v, baseRow.TNum, comparedRow.TNum, true);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.LoLim = v, baseRow.LoLim, comparedRow.LoLim);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.HiLim = v, baseRow.HiLim, comparedRow.HiLim);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.Scale = v, baseRow.Scale, comparedRow.Scale);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.Units = v, baseRow.Units, comparedRow.Units);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.Format = v, baseRow.Format, comparedRow.Format);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.BinPass = v, baseRow.BinPass, comparedRow.BinPass);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.BinFail = v, baseRow.BinFail, comparedRow.BinFail);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.SortPass = v, baseRow.SortPass, comparedRow.SortPass);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.SortFail = v, baseRow.SortFail, comparedRow.SortFail);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.Result = v, baseRow.Result, comparedRow.Result);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.PassAction = v, baseRow.PassAction.Replace(" ", ""), comparedRow.PassAction.Replace(" ", ""));
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.FailAction = v, baseRow.FailAction.Replace(" ", ""), comparedRow.FailAction.Replace(" ", ""));
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.State = v, baseRow.State, comparedRow.State);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.GroupSpecifier = v, baseRow.GroupSpecifier, comparedRow.GroupSpecifier);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.GroupSense = v, baseRow.GroupSense, comparedRow.GroupSense);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.GroupCondition = v, baseRow.GroupCondition, comparedRow.GroupCondition);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.GroupName = v, baseRow.GroupName, comparedRow.GroupName);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.DeviceSense = v, baseRow.DeviceSense, comparedRow.DeviceSense);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.DeviceCondition = v, baseRow.DeviceCondition, comparedRow.DeviceCondition);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.DeviceName = v, baseRow.DeviceName, comparedRow.DeviceName);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.DebugAssume = v, baseRow.DebugAssume, comparedRow.DebugAssume);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.DebugSites = v, baseRow.DebugSites, comparedRow.DebugSites);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.CtProfileDataElapsedTime = v, baseRow.CtProfileDataElapsedTime, comparedRow.CtProfileDataElapsedTime);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.CtProfileDataBackgroundType = v, baseRow.CtProfileDataBackgroundType, comparedRow.CtProfileDataBackgroundType);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.CtProfileDataSerialize = v, baseRow.CtProfileDataSerialize, comparedRow.CtProfileDataSerialize);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.CtProfileDataResourceLock = v, baseRow.CtProfileDataResourceLock, comparedRow.CtProfileDataResourceLock);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.CtProfileDataFlowStepLocked = v, baseRow.CtProfileDataFlowStepLocked, comparedRow.CtProfileDataFlowStepLocked);
                            CompareGetProperty(ref allPass, resultRow, (r, v) => r.Comment = v, baseRow.Comment, comparedRow.Comment, true);

                            if (!allPass)
                            {
                                resultRow.Status = "M";
                                diffCount++;
                            }
                            else
                            {
                                passCount++;
                            }

                            resultRows.Add(resultRow);
                        }

                        result.Add(resultRows);
                    }
                }
            }
            var dummyInCompared = comparedGroups.Values.SelectMany(x => x).SelectMany(x => x).Where(x => !x.IsChecked).ToList();
            dummyInCompared.ForEach(x => x.Status = "MissingBase");
            result.Add(dummyInCompared);
            int totalCount = passCount + diffCount + missingComparedCount;
            double accuracy = totalCount > 0 ? (double)passCount / totalCount : 0;
            var countRow = new CompareReport
            {
                SheetName = SheetName,
                SheetType = EnumSheetType.DTFlowtableSheet,
                TotalCount = totalCount,
                PassCount = passCount,
                ManualCount = totalCount - passCount
            };
            return (result, accuracy, countRow);
        }

        private static void CompareGetProperty(ref bool allPass, FlowRowCompare flowRowCompare, Action<FlowRowCompare, string> setter, string baseVal, string comparedVal, bool byPass = false)
        {
            if (baseVal.EqualsIgnoreCase(comparedVal))
            {
                setter(flowRowCompare, baseVal);
            }
            else
            {
                setter(flowRowCompare, byPass ? baseVal : $"{baseVal} => {comparedVal}");
                if (!byPass)
                {
                    allPass = false;
                }
            }
        }

        private static Dictionary<string, List<List<FlowRowCompare>>> GroupFlowRows(List<List<FlowRowCompare>> groups)
        {
            var flowRowGroups = new Dictionary<string, List<List<FlowRowCompare>>>(StringExtensions.IgnoreCase);
            foreach (List<FlowRowCompare> group in groups)
            {
                string key = GenerateKey(group);
                if (string.IsNullOrEmpty(key))
                {
                    continue;
                }

                if (!flowRowGroups.TryGetValue(key, out List<List<FlowRowCompare>>? value))
                {
                    value = [];
                    flowRowGroups[key] = value;
                }

                value.Add(group);
            }

            return flowRowGroups;
        }

        private static string GenerateKey(List<FlowRowCompare> flowRowCompares)
        {
            if (string.IsNullOrEmpty(flowRowCompares.First().Opcode) && string.IsNullOrEmpty(flowRowCompares.First().Parameter) && string.IsNullOrEmpty(flowRowCompares.First().TName))
            {
                return "";
            }

            return (flowRowCompares.First().Opcode + "|" + flowRowCompares.First().Parameter + "|" + flowRowCompares.First().TName).ToLower();
        }

        public static List<CompareReport> WriteToTxt(Dictionary<string, (List<List<FlowRowCompare>> result, double accuracy, CompareReport countRow)> resultDict, string txtFile)
        {
            var sortedByValueReDic = resultDict.OrderByDescending(x => x.Value.accuracy).ToList();
            var countRows = new List<CompareReport>();

            using (var writer = new StreamWriter(txtFile, false, Encoding.UTF8))
            {
                foreach (KeyValuePair<string, (List<List<FlowRowCompare>> result, double accuracy, CompareReport countRow)> entry in sortedByValueReDic)
                {
                    string sheetName = Path.GetFileNameWithoutExtension(entry.Key);
                    List<List<FlowRowCompare>> result = entry.Value.result;
                    double accuracy = entry.Value.accuracy;

                    var compareReport = new CompareReport()
                    {
                        SheetName = sheetName,
                        PassCount = entry.Value.countRow.PassCount,
                        TotalCount = entry.Value.countRow.TotalCount,
                        ManualCount = entry.Value.countRow.TotalCount - entry.Value.countRow.PassCount,
                    };
                    ClassifyCompareReport.SetCategoryForFlow(sheetName, ref compareReport);
                    countRows.Add(compareReport);

                    // --- Sheet summary header ---
                    writer.WriteLine($"Pass Rate\t{accuracy * 100:0.00}%{string.Concat(Enumerable.Repeat("\t", 20))}");
                    writer.WriteLine($"Total Rows\t{entry.Value.countRow.TotalCount}\tPass Rows\t{entry.Value.countRow.PassCount}{string.Concat(Enumerable.Repeat("\t", 22))}");
                    writer.WriteLine(string.Concat(Enumerable.Repeat("\t", 24)));

                    // --- Column headers ---
                    string[] headers =
                    [
                        "Status",
                        "BaseSheet",
                        "ComparedSheet",
                        "BaseRow",
                        "ComparedRow",
                        "Label",
                        "Enable",
                        "Job",
                        "Part",
                        "Env",
                        "Opcode",
                        "Parameter",
                        "TName",
                        "TNum",
                        "LoLim",
                        "HiLim",
                        "Scale",
                        "Units",
                        "Format",
                        "BinPass",
                        "BinFail",
                        "SortPass",
                        "SortFail",
                        "Result",
                        "PassAction",
                        "FailAction",
                        "State",
                        "GroupSpecifier",
                        "GroupSense",
                        "GroupCondition",
                        "GroupName",
                        "DeviceSense",
                        "DeviceCondition",
                        "DeviceName",
                        "DebugAssume",
                        "DebugSites",
                        "CtProfileDataElapsedTime",
                        "CtProfileDataBackgroundType",
                        "CtProfileDataSerialize",
                        "CtProfileDataResourceLock",
                        "CtProfileDataFlowStepLocked",
                        "Comment"
                    ];
                    writer.WriteLine(string.Join("\t", headers));

                    // --- Write data ---
                    foreach (List<FlowRowCompare> group in result)
                    {
                        foreach (FlowRowCompare rowData in group)
                        {
                            object[] rowValues =
                            [
                                rowData.Status,
                                rowData.BaseSheet,
                                rowData.ComparedSheet,
                                rowData.BaseRow,
                                rowData.ComparedRow,
                                rowData.Label,
                                rowData.Enable,
                                rowData.Job,
                                rowData.Part,
                                rowData.Env,
                                rowData.Opcode,
                                rowData.Parameter,
                                rowData.TName,
                                rowData.TNum,
                                rowData.LoLim,
                                rowData.HiLim,
                                rowData.Scale,
                                rowData.Units,
                                rowData.Format,
                                rowData.BinPass,
                                rowData.BinFail,
                                rowData.SortPass,
                                rowData.SortFail,
                                rowData.Result,
                                rowData.PassAction,
                                rowData.FailAction,
                                rowData.State,
                                rowData.GroupSpecifier,
                                rowData.GroupSense,
                                rowData.GroupCondition,
                                rowData.GroupName,
                                rowData.DeviceSense,
                                rowData.DeviceCondition,
                                rowData.DeviceName,
                                rowData.DebugAssume,
                                rowData.DebugSites,
                                rowData.CtProfileDataElapsedTime,
                                rowData.CtProfileDataBackgroundType,
                                rowData.CtProfileDataSerialize,
                                rowData.CtProfileDataResourceLock,
                                rowData.CtProfileDataFlowStepLocked,
                                rowData.Comment
                            ];

                            writer.WriteLine(string.Join("\t", rowValues.Select(v => v?.ToString() ?? "")));
                        }
                    }

                    writer.WriteLine();
                }
            }

            return countRows;
        }
    }
}
