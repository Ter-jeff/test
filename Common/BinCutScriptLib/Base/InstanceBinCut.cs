using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;

using TestPlanLib.BinCut;

namespace BinCutScriptLib.Base
{
    public class InstanceBinCut
    {
        //instanse type
        private readonly List<string> _tdType = ["SOCTD_", "GFXTD_", "CPUTD_"];
        private readonly List<string> _mbistType = ["SOCMBIST_", "GFXMBIST_", "CPUMBIST_"];

        public BinCutLineBase CurInstanceLine;
        public string CurInstanceName = "";
        public string CurInstType;
        public InstanceRow InstanceRow;
        public string CurDcSpecCondtion;
        public bool HasPayload;
        public List<SelsrmMappingTableRow> SelsrmMappingTalbeRows = [];
        public bool IsBvPass = true;
        public bool IsSearch;

        public List<string> DigSrcAssignmentList { get; } = [];

        public InstanceBinCut(EnumJob enumJob, OneGradeSearch oneGradeSearch, InstanceRow instanceRow, bool isSelSram)
        {
            CurInstanceName = oneGradeSearch.InstanceLine.Line != null ? oneGradeSearch.InstanceLine.Line.Trim() : "";
            CurInstanceLine = oneGradeSearch.InstanceLine;
            CurInstType = GetInstanceType(CurInstanceName);
            InstanceRow = instanceRow;
            CurDcSpecCondtion = GetDcSpecsName(InstanceRow);
            if (isSelSram)
            {
                (SelsrmMappingTalbeRows, DigSrcAssignmentList) = GetSelsrmMappingTalbeRows(enumJob, oneGradeSearch);
            }
        }

        private static (List<SelsrmMappingTableRow>, List<string>) GetSelsrmMappingTalbeRows(EnumJob enumJob, OneGradeSearch oneGradeSearch)
        {
            string digSrcAssignment = GetDigSrcAssigment(oneGradeSearch);
            return GetSelsrmMappingTalbeRowsNew(oneGradeSearch, enumJob, digSrcAssignment);
        }

        private static (List<SelsrmMappingTableRow>, List<string>) GetSelsrmMappingTalbeRowsNew(OneGradeSearch oneGradeSearch, EnumJob enumJob, string digSrcAssignment)
        {
            var patternNames = oneGradeSearch.Steps.SelectMany(x => x.OneStepPatternRows).Select(x => x.PatternName).ToList();
            var selsramList = new List<SelsrmMappingTableRow>();
            List<SelsrmMappingTableRow> rows = BinCutData.SelsrmMappingSheet.Rows;
            var digSrcAssignmentList = new List<string>();
            if (digSrcAssignment.Length != 0)
            {
                string[] elements = digSrcAssignment.Split('-');
                foreach (string element in elements)
                {
                    digSrcAssignmentList.Add(element);
                }
            }
            else
            {
            }

            var rowsDic = rows.GroupBy(x => x.Stage + "|" + x.Block + "|" + x.Pattern).ToDictionary(p => p.Key, p => p.ToList());
            //new method : Compare only stage & pattern column in SELSRM_Mapping_Table sheet.

            if (digSrcAssignmentList.Count > 0)
            {
                foreach (KeyValuePair<string, List<SelsrmMappingTableRow>> row in rowsDic)
                {
                    var diglist = row.Value.Select(x => x.DigSrcAssignment.ToLower()).ToList();

                    if (!digSrcAssignmentList.Except(diglist).Any() &&
                        !diglist.Except(digSrcAssignmentList).Any())
                    {
                        selsramList.AddRange(row.Value);
                    }

                    if (selsramList.Count > 0)
                    {
                        return (selsramList, digSrcAssignmentList);
                    }
                }
            }
            else
            {
                for (int i = 0; i < rowsDic.Count; i++)
                {
                    if (rowsDic.ElementAt(i).Key.Contains("END", StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }

                    string[] key = rowsDic.ElementAt(i).Key.Split('|');
                    if (enumJob != EnumJob.None)
                    {
                        string patText = "^" + key[2].Trim().Replace("*", ".*");
                        if (!patternNames.Exists(x => Regex.IsMatch(x, patText)))
                        {
                            continue;
                        }

                        foreach (string text in key[0].Split(','))
                        {
                            string regexText = "^" + text.Trim().Replace("*", ".*") + "$";
                            if (Regex.IsMatch(enumJob.ToString(), regexText, RegexOptions.IgnoreCase))
                            {
                                selsramList.AddRange(rowsDic.ElementAt(i).Value);
                                if (selsramList.Count > 0)
                                {
                                    return (selsramList, digSrcAssignmentList);
                                }
                            }
                        }
                    }
                }
            }

            if (selsramList.Count <= 0)
            {
                for (int i = 0; i < rowsDic.Count; i++)
                {
                    if (rowsDic.ElementAt(i).Key.Contains("END", StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }

                    string[] key = rowsDic.ElementAt(i).Key.Split('|');
                    if (enumJob != EnumJob.None)
                    {
                        string patText = "^" + key[2].Trim().Replace("*", ".*");
                        if (!patternNames.Exists(x => Regex.IsMatch(x, patText)))
                        {
                            continue;
                        }

                        foreach (string text in key[0].Split(','))
                        {
                            string regexText = "^" + text.Trim().Replace("*", ".*") + "$";
                            if (Regex.IsMatch(enumJob.ToString(), regexText, RegexOptions.IgnoreCase))
                            {
                                selsramList.AddRange(rowsDic.ElementAt(i).Value);
                                if (selsramList.Count > 0)
                                {
                                    return (selsramList, digSrcAssignmentList);
                                }
                            }
                        }
                    }
                }
            }
            return (selsramList, digSrcAssignmentList);
        }

        private static string GetDigSrcAssigment(OneGradeSearch oneGradeSearch)
        {
            foreach (OneStep step in oneGradeSearch.Steps)
            {
                foreach (BinCutLineBase eachDsscLine in step.OneStepDssc)
                {
                    string[] arr = eachDsscLine.Line.Split(',');
                    if (arr.Length > 3)
                    {
                        MatchCollection matches = Reg.RegexDigSrcAssignment.Matches(eachDsscLine.Line);

                        List<string> elements = [];

                        foreach (Match match in matches)
                        {
                            elements.Add(match.Groups[1].Value);
                        }
                        string digScrAssigmentString = string.Join("-", elements);
                        return digScrAssigmentString;
                    }

                }
            }
            return "";
        }

        private static string GetDcSpecsName(InstanceRow instanceRow)
        {
            string dcSpecs = "BINCUT_X_X_X";
            if (instanceRow != null)
            {
                dcSpecs = (instanceRow.DcCategory + "_" + instanceRow.DcSelector).ToUpper();
            }

            return dcSpecs;
        }

        private string GetInstanceType(string instanceName)
        {
            string type;
            foreach (string item in _tdType)
            {
                if (instanceName.ContainsIgnoreCase(item))
                {
                    type = "TD";
                    return type;
                }
            }

            foreach (string item in _mbistType)
            {
                if (instanceName.ContainsIgnoreCase(item))
                {
                    type = "MBIST";
                    return type;
                }
            }

            type = "FUNC";
            return type;
        }
    }
}
