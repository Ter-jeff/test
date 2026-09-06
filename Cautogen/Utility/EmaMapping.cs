using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using OfficeOpenXml;

namespace Cautogen.Utility
{
    public class EmaMappingReader
    {
        private readonly List<EmaMappingItem> _emaMappingRows = new List<EmaMappingItem>();
        private readonly Dictionary<string, int> _headerOrderDic = new Dictionary<string, int>();
        private readonly Dictionary<string, int> _groupHeaderDic = new Dictionary<string, int>();

        public List<EmaMappingItem> ReadMappingTable(ExcelWorksheet sheet)
        {
            _emaMappingRows.Clear();

            // read header
            int rowindex = 1;
            //var startRow = _SearchTableStartRow(sheet, ref rowindex);
            //var endRow = sheet.Dimension.End.Row;

            // early return
            int patternIndex = -1;
            while (_SearchTableStartRow(sheet, ref rowindex, ref patternIndex))
            {
                EmaMappingItem item = null;
                var datasets = new List<EmaSubset>();
                if (rowindex >= sheet.Dimension.End.Row)
                {
                    return _emaMappingRows;
                }

                for (int i = rowindex; i <= sheet.Dimension.End.Row; i++)
                {
                    EmaSubset set = null;
                    if (_headerOrderDic.TryGetValue("PATTERN", out int value))
                    {
                        string cellValue = sheet.Cells[i, value].Text;
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            item = new EmaMappingItem();
                            _emaMappingRows.Add(item);
                            item.Pattern = cellValue;
                        }
                    }
                    if (_headerOrderDic.TryGetValue("SEGMENTNUMBER", out int value1))
                    {
                        string cellValue = sheet.Cells[i, value1].Text;
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            set = new EmaSubset(cellValue);
                            datasets.Add(set);
                        }
                    }
                    else
                    {
                        continue;
                    }

                    foreach (KeyValuePair<string, int> group in _groupHeaderDic)
                    {
                        string cellValue = sheet.Cells[i, _groupHeaderDic[group.Key]].Text;
                        if (string.IsNullOrEmpty(cellValue))
                        {
                            cellValue = "";
                        }

                        if (!item.CasesList.Contains(group.Key))
                        {
                            item.CasesList.Add(group.Key);
                        }

                        set.Data.Add(group.Key, cellValue);
                    }
                    rowindex = i;
                }
                item.ReferenceSets = datasets.GroupBy(p => p.Sgmt + "#" + p.Register).ToDictionary(p => p.Key, p => p.ToList());
            }
            return _emaMappingRows;
        }

        private void _UpdateTableHeader(ExcelWorksheet sheet, int startrow)
        {
            _headerOrderDic.Clear();
            _groupHeaderDic.Clear();
            for (int i = 1; i <= sheet.Dimension.End.Column; i++)
            {
                string header = "";
                if (sheet.Cells[startrow, i].Value != null)
                {
                    header = sheet.Cells[startrow, i].Value.ToString().Replace(" ", "").ToUpper();
                }

                if (header == "PATTERN" || header == "PATTERNNAME")
                {
                    header = "PATTERN";
                }

                if (header.Trim() != "")
                {
                    if (header == "PATTERN" || header == "SEGMENTNUMBER" ||
                        header == "VECTORNUMBER" || header == "CYCLENUMBER" ||
                        header == "EMA")
                    {
                        _headerOrderDic[header] = i;
                    }
                    else
                    {
                        _groupHeaderDic[header] = i;
                    }
                }
            }
        }

        private bool _SearchTableStartRow(ExcelWorksheet sheet, ref int rowindex, ref int patternIndex)
        {
            /* Return the row number of Column A == "Pattern Name" */
            for (int i = rowindex; i <= sheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= sheet.Dimension.End.Column; j++)
                {
                    if (sheet.Cells[i, j].Value != null &&
                        (sheet.Cells[i, j].Value.ToString().Replace(" ", "").ToUpper() == "PATTERN" ||
                         sheet.Cells[i, j].Value.ToString().Replace(" ", "").ToUpper() == "PATTERNNAME"))
                    {
                        _UpdateTableHeader(sheet, rowindex);
                        patternIndex = j;
                        rowindex = i + 1;
                        return true;
                    }
                }
            }
            return false;
        }

    }

    public class EmaMappingItem
    {
        public string Pattern { get; set; }
        public List<string> CasesList = new List<string>();
        public Dictionary<string, List<EmaSubset>> ReferenceSets = new Dictionary<string, List<EmaSubset>>(StringComparer.OrdinalIgnoreCase);

        public string GetCaseData(string key)
        {
            var result = new List<string>();

            var setGroups = ReferenceSets.SelectMany(p => p.Value).GroupBy(p => p.Sgmt).ToDictionary(p => p.Key, p => p.ToList());
            foreach (KeyValuePair<string, List<EmaSubset>> group in setGroups)
            {
                string tmpValue = "";
                var bitList =
                    group.Value.Select(
                        p => p.BitPos).ToList();
                for (int i = 0; i <= bitList.Max(); i++)
                {
                    string value = "0";
                    EmaSubset target = group.Value.FirstOrDefault(p => p.BitPos == i);
                    if (target != null)
                    {
                        if (target.Data.TryGetValue(key, out string value1))
                        {
                            value = value1;
                        }
                    }
                    tmpValue += value;
                }
                result.Add(group.Key + 'f' + tmpValue);
            }
            return "sgmtdef0" + string.Join("", result);
        }
    }

    public class EmaSubset
    {
        public string Sgmt;
        public string Register;
        public int BitPos;
        public Dictionary<string, string> Data = new Dictionary<string, string>();

        public EmaSubset(string segment)
        {
            segment = segment.Replace(" ", "");
            string regbitPos = @"\[(?<pos>\d+)\]";
            string[] part = Regex.Split(segment, @"(sgmt\d+)").Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();
            Sgmt = part[0];
            Register = part[1];
            try
            {
                BitPos = int.Parse(Regex.Match(Register, regbitPos, RegexOptions.IgnoreCase).Groups["pos"].Value);
            }
            catch (Exception)
            {
            }
        }
    }

}
