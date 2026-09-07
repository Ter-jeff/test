using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

using ScghLib.Base;
using ScghLib.Enums;
using ScghLib.Utility;

namespace ScghLib.Reader
{
    public class BistProdFlowReader : MySheetReader<BistProdFlowSheet>
    {
        public const string ConHeaderLabel = "Label";
        public const string ConHeaderAction = "Action";
        public const string ConHeaderPattern = "Pattern";
        public const string ConHeaderTimeSet = "Timeset";
        public const string ConHeaderVoltage = "Voltage";
        public const string ConHeaderNote = "Notes";
        public const string ConHeaderPassBranch = "Pass Branch";
        public const string ConHeaderFailBranch = "Fail Branch";
        public const string ConHeaderTestName = "TestName";

        private const int MaxColumnCount = 10;
        private const int MaxRowCount = 10;

        private readonly BistProdFlowSheet _outProdFlowSheet;

        private readonly Dictionary<string, int> _headColumnDictionary;
        private int _testNameColumnIndex = -1;

        public int LabelColumn = -1;
        public int ActionColumn = -1;
        public int PatternColumn = -1;
        public int TimeSetColumn = -1;
        public int VoltageColumn = -1;
        public int NoteColumn = -1;
        public int PassColumn = -1;
        public int FailColumn = -1;

        private readonly MbistSheet _mbistSheet;

        public BistProdFlowReader(MbistSheet mbistSheet)
        {
            _mbistSheet = mbistSheet;
            _headColumnDictionary = [];
            _outProdFlowSheet = new BistProdFlowSheet();
            StartRow = 0;
            StartCol = 0;
        }

        public override BistProdFlowSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;

            _outProdFlowSheet.MbistSheet = _mbistSheet;

            ReadHeader();

            ReadData();

            _outProdFlowSheet.Rows.ForEach(x => x.BistActionType = BistAction.GetActionType(x));
            _outProdFlowSheet.Rows.ForEach(x => x.Chiplet = _outProdFlowSheet.Chiplet);
            return _outProdFlowSheet;
        }

        private void ReadHeader()
        {
            string header;
            for (int i = 1; i < MaxRowCount; i++)
            {
                for (int j = 1; j < MaxColumnCount; j++)
                {
                    header = EpplusExtensions.GetCellValue(ExcelWorksheet, i, j);
                    if (header.EqualsIgnoreCase(ConHeaderLabel))
                    {
                        StartRow = i;
                        StartCol = j;
                        break;
                    }
                }
            }

            if (StartCol == 0 || StartRow == 0)
            {
                throw new Exception($"Read Header Error when Read Sheet: {ExcelWorksheet.Name}");
            }
            EndCol = ExcelWorksheet.Dimension.End.Column;
            EndRow = ExcelWorksheet.Dimension.End.Row;

            for (int i = StartCol; i <= EndCol; i++)
            {
                header = EpplusExtensions.GetCellValue(ExcelWorksheet, StartRow, i);
                if (header.EqualsIgnoreCase(ConHeaderLabel))
                {
                    _headColumnDictionary.Add(ConHeaderLabel, i);
                }
                else if (header.EqualsIgnoreCase(ConHeaderAction))
                {
                    _headColumnDictionary.Add(ConHeaderAction, i);
                }
                else if (header.EqualsIgnoreCase(ConHeaderPattern))
                {
                    _headColumnDictionary.Add(ConHeaderPattern, i);
                }
                else if (header.EqualsIgnoreCase(ConHeaderTimeSet))
                {
                    _headColumnDictionary.Add(ConHeaderTimeSet, i);
                }
                else if (header.EqualsIgnoreCase(ConHeaderVoltage))
                {
                    _headColumnDictionary.Add(ConHeaderVoltage, i);
                }
                else if (header.EqualsIgnoreCase(ConHeaderTestName))
                {
                    _testNameColumnIndex = i;
                }
                else if (header.EqualsIgnoreCase(ConHeaderNote))
                {
                    _headColumnDictionary.Add(ConHeaderNote, i);
                }
                else if (header.EqualsIgnoreCase(ConHeaderPassBranch))
                {
                    _headColumnDictionary.Add(ConHeaderPassBranch, i);
                }
                else if (header.EqualsIgnoreCase(ConHeaderFailBranch))
                {
                    _headColumnDictionary.Add(ConHeaderFailBranch, i);
                }
            }
        }

        private void ReadData()
        {
            try
            {
                LabelColumn = _headColumnDictionary[ConHeaderLabel];
                ActionColumn = _headColumnDictionary[ConHeaderAction];
                PatternColumn = _headColumnDictionary[ConHeaderPattern];
                TimeSetColumn = _headColumnDictionary[ConHeaderTimeSet];
                VoltageColumn = _headColumnDictionary[ConHeaderVoltage];
                NoteColumn = _headColumnDictionary[ConHeaderNote];
                PassColumn = _headColumnDictionary[ConHeaderPassBranch];
                FailColumn = _headColumnDictionary[ConHeaderFailBranch];
                for (int i = StartRow + 1; i <= EndRow; i++)
                {
                    var row = new BistProdFlowRow
                    {
                        Label = EpplusExtensions.GetCellValue(ExcelWorksheet, i, LabelColumn),
                        Action = EpplusExtensions.GetCellValue(ExcelWorksheet, i, ActionColumn),
                        Pattern = EpplusExtensions.GetCellValue(ExcelWorksheet, i, PatternColumn)
                    };

                    if (row.Pattern.EqualsIgnoreCase("N/A"))
                    {
                        row.Pattern = "";
                    }
                    row.TimeSet = EpplusExtensions.GetCellValue(ExcelWorksheet, i, TimeSetColumn);
                    row.Note = EpplusExtensions.GetCellValue(ExcelWorksheet, i, NoteColumn);
                    row.PassBranch = EpplusExtensions.GetCellValue(ExcelWorksheet, i, PassColumn);
                    row.FailBranch = EpplusExtensions.GetCellValue(ExcelWorksheet, i, FailColumn);
                    row.Voltage = EpplusExtensions.GetCellValue(ExcelWorksheet, i, VoltageColumn);
                    if (_testNameColumnIndex != -1)
                    {
                        row.TestName = EpplusExtensions.GetCellValue(ExcelWorksheet, i, _testNameColumnIndex);
                    }

                    row.SheetName = ExcelWorksheet.Name;
                    row.IsMultipleMbistScghSheet = _outProdFlowSheet.MbistSheet.IsMultipleSheet;
                    row.RowNum = i;
                    if (!string.IsNullOrEmpty(row.Label))
                    {
                        _outProdFlowSheet.Rows.Add(row);
                    }

                }
                _outProdFlowSheet.LabelColumn = LabelColumn;
                _outProdFlowSheet.ActionColumn = ActionColumn;
                _outProdFlowSheet.PatternColumn = PatternColumn;
                _outProdFlowSheet.TimeSetColumn = TimeSetColumn;
                _outProdFlowSheet.VoltageColumn = VoltageColumn;
                _outProdFlowSheet.NoteColumn = NoteColumn;
                _outProdFlowSheet.PassColumn = PassColumn;
                _outProdFlowSheet.FailColumn = FailColumn;
            }
            catch (Exception e)
            {
                throw new Exception($"Read Data Error When Read Sheet: {ExcelWorksheet.Name}" + e.StackTrace);
            }
        }
    }

    public partial class BistProdFlowSheet : MySheet
    {
        private const string RegChiplet = @"\w+_(?<chiplet>[A-z]\d+$)";

        [GeneratedRegex(RegChiplet, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();

        private List<BistProdFlowRow> _dataRows = [];

        public int LabelColumn = -1;
        public int ActionColumn = -1;
        public int PatternColumn = -1;
        public int TimeSetColumn = -1;
        public int VoltageColumn = -1;
        public int NoteColumn = -1;
        public int PassColumn = -1;
        public int FailColumn = -1;

        public Dictionary<string, List<BistProdFlowRow>> LabelDic { get; private set; } = new Dictionary<string, List<BistProdFlowRow>>(StringExtensions.IgnoreCase);
        public List<BistProdFlowRow> Rows
        {
            set { _dataRows = value; }
            get
            {
                return _dataRows ??= [];
            }
        }
        public MbistSheet MbistSheet { get; set; } = new MbistSheet();
        public string Chiplet
        {
            get { return MyRegex().Match(MbistSheet.SheetName).Groups["chiplet"].ToString(); }
        }

        public BistProdFlowSheet()
        {
        }

        public BistProdFlowSheet(BistProdFlowSheet bistProdFlowSheet) : base(bistProdFlowSheet)
        {
            if (bistProdFlowSheet == null)
            {
                return;
            }

            LabelColumn = bistProdFlowSheet.LabelColumn;
            ActionColumn = bistProdFlowSheet.ActionColumn;
            PatternColumn = bistProdFlowSheet.PatternColumn;
            TimeSetColumn = bistProdFlowSheet.TimeSetColumn;
            VoltageColumn = bistProdFlowSheet.VoltageColumn;
            NoteColumn = bistProdFlowSheet.NoteColumn;
            PassColumn = bistProdFlowSheet.PassColumn;
            FailColumn = bistProdFlowSheet.FailColumn;

            MbistSheet = bistProdFlowSheet.MbistSheet?.Copy() ?? new MbistSheet();

            Rows = bistProdFlowSheet.Rows?.Select(row => row.Copy()).ToList() ?? [];

            if (bistProdFlowSheet.LabelDic != null)
            {
                LabelDic = bistProdFlowSheet.LabelDic.ToDictionary(k => k.Key, v => v.Value.Select(row => row.Copy()).ToList());
            }
        }

        public BistProdFlowSheet Copy()
        {
            return new BistProdFlowSheet(this);
        }

        public bool TryGetSheetRow(string label, out BistProdFlowRow? bistProdFlowRow)
        {
            if (LabelDic.TryGetValue(label, out List<BistProdFlowRow>? list) && list.Count > 0)
            {
                bistProdFlowRow = list[0];
                return true;
            }

            bistProdFlowRow = null;
            return false;
        }

        public void CreateLabelDic()
        {
            if (_dataRows != null)
            {
                LabelDic = _dataRows.GroupBy(x => x.Label).ToDictionary(x => x.Key, x => x.ToList(), StringExtensions.IgnoreCase);
            }
        }

    }

    public partial class BistProdFlowRow : MyRow
    {
        private string _action = string.Empty;
        public BistActionType BistActionType;

        public string FailFlag { set; get; } = string.Empty;
        public string Label { get; set; } = string.Empty;
        public string Action
        {
            set { _action = value; }
            get { return _action.StartsWithIgnoreCase("patt_") ? BistProdFlowRowHelpers.MyRegex().Replace(_action, "") : _action; }
        }
        public string Pattern { get; set; } = string.Empty;
        public List<string> BurstPatterns { set; get; } = [];
        public bool IsPatBurst { set; get; }
        public string TimeSet { get; set; } = string.Empty;
        public string Voltage { get; set; } = string.Empty;
        public string Note { get; set; } = string.Empty;
        public string PassBranch { get; set; } = string.Empty;
        public string FailBranch { get; set; } = string.Empty;
        public string PerformanceMode { get; set; } = string.Empty;
        public string VoltageMode { get; set; } = string.Empty;
        public string Chiplet { set; get; } = string.Empty;
        public string NextPayLoadPattern { get; set; } = string.Empty;
        public bool IsDsscRow { set; get; }
        public bool IsMultipleMbistScghSheet { set; get; }
        public string OriVoltage { get; set; } = string.Empty;
        public string OriPerformance { get; set; } = string.Empty;
        public string OtherPModeInfoStr { get; set; } = string.Empty;
        public string TestName { set; get; } = string.Empty;
        public string DcCategory { get; set; } = string.Empty;
        public string EqnVoltage { set; get; } = string.Empty;

        public BistProdFlowRow()
        {
        }

        public BistProdFlowRow(BistProdFlowRow bistProdFlowRow) : base(bistProdFlowRow)
        {
            if (bistProdFlowRow == null)
            {
                return;
            }

            Label = bistProdFlowRow.Label;
            _action = bistProdFlowRow._action;
            Pattern = bistProdFlowRow.Pattern;
            TimeSet = bistProdFlowRow.TimeSet;
            Voltage = bistProdFlowRow.Voltage;
            Note = bistProdFlowRow.Note;
            PassBranch = bistProdFlowRow.PassBranch;
            FailBranch = bistProdFlowRow.FailBranch;
            VoltageMode = bistProdFlowRow.VoltageMode;
            IsPatBurst = bistProdFlowRow.IsPatBurst;
            FailFlag = bistProdFlowRow.FailFlag;
            NextPayLoadPattern = bistProdFlowRow.NextPayLoadPattern;
            IsDsscRow = bistProdFlowRow.IsDsscRow;
            OriVoltage = bistProdFlowRow.OriVoltage;
            OriPerformance = bistProdFlowRow.OriPerformance;
            OtherPModeInfoStr = bistProdFlowRow.OtherPModeInfoStr;
            Chiplet = bistProdFlowRow.Chiplet;
            EqnVoltage = bistProdFlowRow.EqnVoltage;
            IsMultipleMbistScghSheet = bistProdFlowRow.IsMultipleMbistScghSheet;

            // Deep copy the list of strings
            BurstPatterns = bistProdFlowRow.BurstPatterns != null
                ? [.. bistProdFlowRow.BurstPatterns]
                : [];

            // Copy public fields and auto-properties
            BistActionType = bistProdFlowRow.BistActionType;
            PerformanceMode = bistProdFlowRow.PerformanceMode;
            TestName = bistProdFlowRow.TestName;
            DcCategory = bistProdFlowRow.DcCategory;
        }

        public BistProdFlowRow Copy()
        {
            return new BistProdFlowRow(this);
        }

        public bool IsRscr()
        {
            if (string.IsNullOrEmpty(Pattern))
            {
                return false;
            }

            string[] patItems = Pattern.Split('_');
            if (patItems.Length < 7)
            {
                return false;
            }

            if (patItems.Contains("RSCRCLR"))
            {
                return false;
            }

            return patItems[6].ContainsIgnoreCase("RPI") && patItems[4].ContainsIgnoreCase("BI");
        }
    }
}
