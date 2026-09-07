using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public partial class BinCutInstanceNamingRuleReader : MySheetReader<BinCutInstanceNamingSheet>
    {
        public const string ConInit = "Init";
        public const string ConPayload = "payload";
        public const string ConUnknow = "Unknow";
        private const string ConHeaderFlowName = "Flow name";

        [GeneratedRegex(@"Pattern\d*", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        private int _indexFlowName = -1;
        private int _indexPatternStart = -1;
        private int _patternCnt;

        public override BinCutInstanceNamingSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            GetDimensions();

            GetFirstHeaderPosition();

            GetHeaderIndex();

            BinCutInstanceNamingSheet binCutInstanceNamingsheet = ReadSheet(sheetName);

            return binCutInstanceNamingsheet;
        }

        private BinCutInstanceNamingSheet ReadSheet(string sheetName)
        {
            var sheet = new BinCutInstanceNamingSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new BinCutInstanceNamingRow(sheetName)
                {
                    RowNum = i,
                    FlowName = ExcelWorksheet.GetCellValue(i, _indexFlowName)
                };

                if (_indexPatternStart != -1)
                {
                    for (int col = 0; col < _patternCnt; col++)
                    {
                        string payload = ExcelWorksheet.GetCellValue(i, _indexPatternStart + col).Trim();
                        row.PatternList.Add(payload);
                    }
                    sheet.Rows.Add(row);
                }
            }

            sheet.IndexFlowName = _indexFlowName;
            sheet.IndexPatternStart = _indexPatternStart;
            sheet.PatternCnt = _patternCnt;
            return sheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string lStrHeader = ExcelWorksheet.GetCellValue(StartRow, i).ToUpper().Trim();

                if (lStrHeader.EqualsIgnoreCase(ConHeaderFlowName))
                {
                    _indexFlowName = i;
                    continue;
                }

                if (_regex.IsMatch(lStrHeader))
                {
                    if (_indexPatternStart == -1)
                    {
                        _indexPatternStart = i;
                    }
                    _patternCnt++;
                }
            }
        }

        private void GetFirstHeaderPosition()
        {
            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderFlowName))
                    {
                        StartRow = i;
                        break;
                    }
                }
            }
        }
    }

    public partial class BinCutInstanceNamingSheet : MySheet
    {
        [GeneratedRegex(@":\d+", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        public List<BinCutInstanceNamingRow> Rows = [];

        public int IndexFlowName { get; set; }
        public int IndexPatternStart { get; set; }
        public int PatternCnt { get; set; }

        public BinCutInstanceNamingSheet(string sheetName)
        {
            SheetName = sheetName;
        }

        public string GetInitName(BinCutInstanceRow binCutInstanceRow, string mode)
        {
            BinCutInstanceNamingRow? binCutInstanceNamingRow = Rows.Find(x => x.FlowName.EqualsIgnoreCase(binCutInstanceRow.FlowName));
            if (binCutInstanceNamingRow == null)
            {
                List<string> arr = [.. binCutInstanceRow.FlowName.Split(' ')];
                arr.RemoveAt(0);
                string flowName = string.Join(" ", arr);
                binCutInstanceNamingRow = Rows.Find(x => x.FlowName.EqualsIgnoreCase(flowName));
            }
            if (binCutInstanceNamingRow == null)
            {
                return "";
            }

            var initNames = new List<string>();
            if (binCutInstanceNamingRow.PatternList.Exists(x => x.Contains(':')))
            {
                initNames.AddRange(GetInitNameByOrder(binCutInstanceRow, binCutInstanceNamingRow.PatternList));
            }
            else
            {
                for (int i = 0; i < binCutInstanceNamingRow.PatternList.Count; i++)
                {
                    string rule = binCutInstanceNamingRow.PatternList[i];
                    if (string.IsNullOrEmpty(rule))
                    {
                        continue;
                    }

                    if (i < binCutInstanceRow.PatternList.Count)
                    {
                        string pattern = binCutInstanceRow.PatternList[i];
                        if (!pattern.IsInit())
                        {
                            continue;
                        }

                        List<string> arr = [.. rule.Split(',')];
                        List<string> patternArr = [.. pattern.Split('_')];
                        foreach (string item in arr)
                        {
                            if (int.TryParse(item, out int value))
                            {
                                if (value >= 0 && value < patternArr.Count)
                                {
                                    initNames.Add(patternArr[value]);
                                }
                            }
                        }
                    }
                }
            }
            initNames = [.. initNames.Where(x => !x.EqualsIgnoreCase(mode))];
            return string.Join("_", initNames);
        }

        public static List<string> GetInitNameByOrder(BinCutInstanceRow binCutInstanceRow, List<string> patList)
        {
            var initNames = new List<string>();
            var initNamesNormal = new List<string>();
            var nameDic = new Dictionary<int, List<string>>();

            for (int i = 0; i < binCutInstanceRow.InitList.Count; i++)
            {
                string rule = patList[i];
                string pattern = binCutInstanceRow.InitList[i];
                List<string> patternArr = [.. pattern.Split('_')];
                if (!rule.Contains(':'))
                {
                    List<string> arr = [.. rule.Split(',')];
                    foreach (string item in arr)
                    {
                        if (int.TryParse(item, out int value))
                        {
                            if (value >= 0 && value < patternArr.Count)
                            {
                                initNamesNormal.Add(patternArr[value]);
                            }
                        }
                    }
                }
                else
                {
                    int index = Convert.ToInt32(rule.Split(':').Last());
                    rule = _regex.Replace(rule, "");
                    List<string> arr = [.. rule.Split(',')];
                    var arrList = new List<string>();
                    foreach (string item in arr)
                    {
                        if (int.TryParse(item, out int value))
                        {
                            if (value >= 0 && value < patternArr.Count)
                            {
                                arrList.Add(patternArr[value]);
                            }
                        }
                    }
                    while (nameDic.ContainsKey(index))
                    {
                        index++;
                    }

                    nameDic.Add(index, arrList);
                }
            }

            initNames.AddRange(nameDic.OrderBy(x => x.Key).SelectMany(x => x.Value));
            initNames.AddRange(initNamesNormal);
            return initNames;
        }

    }

    public class BinCutInstanceNamingRow : MyRow
    {
        public string FlowName = "";
        public List<string> PatternList;

        #region Constructor
        public BinCutInstanceNamingRow(string sheetName = "")
        {
            SheetName = sheetName;
            PatternList = [];
        }
        #endregion
    }
}
