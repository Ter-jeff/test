using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc
{
    public partial class HarvPmodeTableGroupReader(int harvCheckCnt, int bincutPatCnt) : MySheetReader<HarvPmodeTableGroupSheet>
    {
        private const string ConHeaderByMode = "By_Mode";
        private const string ConHeaderMainCore = "Main_Core";
        private const string ConHeaderSequence = "Sequence";
        private const string ConHeaderFailflagName = "FailFlag Name";
        private const string ConHeaderDeviceCondition = "Device Condition";

        [GeneratedRegex(@"(?<flag>[^\d]+)(?<value>\d+)", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex();

        private readonly int _harvCheckCnt = harvCheckCnt;
        private readonly int _bincutPatCnt = bincutPatCnt;
        private int _indexByMode = -1;
        private int _indexMainCore = -1;
        private int _indexSequence = -1;
        private readonly Dictionary<string, int> _indexFailflagNames = [];
        private readonly Dictionary<string, int> _indexDeviceCondition = [];
        private readonly Dictionary<string, int> _indexGroups = [];
        private Dictionary<int, int> _sequenceList = [];

        public override HarvPmodeTableGroupSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            HarvPmodeTableGroupSheet harvPmodeTableGroupTableSheet = ReadSheet(sheetName, _harvCheckCnt, _bincutPatCnt);

            return harvPmodeTableGroupTableSheet;
        }

        private HarvPmodeTableGroupSheet ReadSheet(string sheetName, int harvCheckCnt, int bincutPatCnt)
        {
            var hArvPmodeTableGroupTableSheet = new HarvPmodeTableGroupSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new HarvPmodeTableGroupRow(sheetName)
                {
                    Groups = [],
                    FailflagName = [],
                    DeviceCondition = [],
                    RowNum = i
                };
                if (_indexByMode != -1)
                {
                    row.ByMode = ExcelWorksheet.GetCellValue(i, _indexByMode).Trim();
                }

                if (_indexMainCore != -1)
                {
                    row.MainCore = ExcelWorksheet.GetCellValue(i, _indexMainCore).Trim();
                }

                if (_indexSequence != -1)
                {
                    row.Sequence = ExcelWorksheet.GetCellValue(i, _indexSequence).Trim();
                }

                if (_indexFailflagNames.Count != 0)
                {
                    foreach (KeyValuePair<string, int> index in _indexFailflagNames)
                    {
                        row.FailflagName.Add(ExcelWorksheet.GetCellValue(i, index.Value).Trim());
                    }
                }
                if (_indexDeviceCondition.Count != 0)
                {
                    foreach (KeyValuePair<string, int> index in _indexDeviceCondition)
                    {
                        row.DeviceCondition.Add(ExcelWorksheet.GetCellValue(i, index.Value).Trim());
                    }
                }
                foreach (KeyValuePair<string, int> index in _indexGroups)
                {
                    row.Groups.Add(index.Key, ExcelWorksheet.GetCellValue(i, index.Value).Trim());
                }
                if (string.IsNullOrEmpty(row.ByMode))
                {
                    break;
                }

                hArvPmodeTableGroupTableSheet.Rows.Add(row);

                for (int j = 1; j < bincutPatCnt; j++)
                {
                    var rowOther = new HarvPmodeTableGroupRow(sheetName)
                    {
                        ByMode = row.ByMode,
                        MainCore = row.MainCore,
                        Sequence = row.Sequence,
                        Groups = row.Groups,
                        DeviceCondition = row.DeviceCondition,
                        FailflagName = []
                    };
                    foreach (string failFlag in row.FailflagName)
                    {
                        string flagNew;
                        _ = int.TryParse(_regex.Match(failFlag).Groups["value"].ToString(), out int index);
                        string flag = _regex2.Match(failFlag).Groups["flag"].ToString();
                        if (index != -1)
                        {
                            int indexNew = index + harvCheckCnt;
                            flagNew = flag + indexNew;
                        }
                        else
                        {
                            flagNew = failFlag;
                        }

                        rowOther.FailflagName.Add(flagNew);
                    }
                    hArvPmodeTableGroupTableSheet.Rows.Add(rowOther);
                }
            }

            hArvPmodeTableGroupTableSheet.IndexByMode = _indexByMode;
            hArvPmodeTableGroupTableSheet.IndexMainCore = _indexMainCore;
            hArvPmodeTableGroupTableSheet.IndexSequence = _indexSequence;
            hArvPmodeTableGroupTableSheet.IndexFailflagName = _indexFailflagNames;
            hArvPmodeTableGroupTableSheet.IndexDeviceCondition = _indexDeviceCondition;
            hArvPmodeTableGroupTableSheet.IndexGroups = _indexGroups;
            hArvPmodeTableGroupTableSheet.SequenceList = _sequenceList;
            return hArvPmodeTableGroupTableSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderByMode))
                {
                    _indexByMode = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderMainCore))
                {
                    _indexMainCore = i;
                    continue;
                }
                if (header.StartsWithIgnoreCase(ConHeaderSequence))
                {
                    _indexSequence = i;
                    var sequenceDic = new Dictionary<int, int>();
                    string[] split = header.Split([' ', '-'], StringSplitOptions.RemoveEmptyEntries);
                    foreach (string s in split)
                    {
                        string[] sequence = s.Contains("SEQUENCE", StringComparison.OrdinalIgnoreCase) ? s.Split('e').Last().Trim('[', ']').Split(':') : s.Trim('[', ']').Split(':');
                        sequenceDic.Add(Convert.ToInt32(sequence[0]), Convert.ToInt32(sequence.Last()));
                    }
                    _sequenceList = sequenceDic;
                    continue;
                }
                if (header.EndsWithIgnoreCase(ConHeaderFailflagName))
                {
                    _indexFailflagNames.Add(header.Split(':').First(), i);
                    continue;
                }
                if (header.EndsWithIgnoreCase(ConHeaderDeviceCondition))
                {
                    _indexDeviceCondition.Add(header.Split(':').First(), i);
                    continue;
                }
                if (header.Contains("GROUP", StringComparison.OrdinalIgnoreCase))
                {
                    _indexGroups.Add(header, i);
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            _ = EndRow > 10 ? 10 : EndRow;

            _ = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= EndRow; i++)
            //for (int j = 1; j <= colNum; j++)
            {
                if (ExcelWorksheet.GetCellValue(i, 1).Trim().EqualsIgnoreCase(ConHeaderByMode))
                //if (ExcelWorksheet.GetCellValue( i, j).Trim().Equals(ConHeaderByMode, StringComparison.OrdinalIgnoreCase))
                {
                    StartRow = i;
                    return true;
                }
            }
            return false;
        }
    }

    public class HarvPmodeTableGroupSheet : MySheet
    {
        public List<HarvPmodeTableGroupRow> Rows { set; get; }

        public int IndexByMode = -1;
        public int IndexMainCore = -1;
        public int IndexSequence = -1;
        public Dictionary<string, int> IndexFailflagName = [];
        public Dictionary<string, int> IndexDeviceCondition = [];
        public Dictionary<string, int> IndexGroups = [];
        public Dictionary<int, int> SequenceList = [];

        #region Constructor
        public HarvPmodeTableGroupSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }
        #endregion

        public List<string>? GetMainFlag(string name)
        {
            return Rows.Exists(x => x.ByMode.EqualsIgnoreCase(name))
                ? Rows.Find(x => x.ByMode.EqualsIgnoreCase(name))!.FailflagName
                : null;
        }

        public string GetCondition(string name, List<string> deviceCondition, string groupName)
        {
            HarvPmodeTableGroupRow? row = GetRow(name, deviceCondition);
            return row != null ? row.Groups.First(x => x.Key.EqualsIgnoreCase(groupName)).Value : "";
        }

        public HarvPmodeTableGroupRow? GetRow(string name, List<string> deviceConditions)
        {
            return Rows.Exists(x => x.ByMode.EqualsIgnoreCase(name) &&
               x.DeviceCondition.SequenceEqual(deviceConditions, StringExtensions.IgnoreCase))
                ? Rows.Find(x => x.ByMode.EqualsIgnoreCase(name) &&
                     x.DeviceCondition.SequenceEqual(deviceConditions, StringExtensions.IgnoreCase))
                : null;
        }
    }

    public class HarvPmodeTableGroupRow : MyRow
    {
        public string ByMode { set; get; } = "";
        public string MainCore { set; get; } = "";
        public string Sequence { set; get; } = "";
        public List<string> FailflagName { set; get; } = [];
        public List<string> DeviceCondition { set; get; } = [];
        public Dictionary<string, string> Groups { set; get; } = [];

        #region Constructor
        public HarvPmodeTableGroupRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion
    }
}
