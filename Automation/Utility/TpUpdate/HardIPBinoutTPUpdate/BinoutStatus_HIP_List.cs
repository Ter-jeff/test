using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using OfficeOpenXml;

namespace Automation.Utility.TpUpdate.HardIPBinoutTPUpdate
{
    public class BinOutStatusHipListRow
    {
        #region Properity
        public string SourceSheetName { set; get; }
        public int RowNum { get; set; }
        public string Testinstance { set; get; }
        public string Testnum { set; get; }
        public string Testname { set; get; }
        public string Type { set; get; }
        public string PatternPin { set; get; }
        public string Lolimit { set; get; }
        public string Hilimit { set; get; }
        public string Unit { set; get; }
        public string TestResult { set; get; }
        public string Meterirangesetting { set; get; }
        public Dictionary<string, string> StatusDictionary = new Dictionary<string, string>();
        public Dictionary<string, string> EnableDictionary = new Dictionary<string, string>();
        public string TPuselimit { set; get; }
        public string Comment { set; get; }

        #region Update Column
        public Dictionary<string, string> BinOutEnableDictionary = new Dictionary<string, string>();
        public Dictionary<string, string> UpdatedLoLimitDic = new Dictionary<string, string>();
        public Dictionary<string, string> UpdatedHiLimitDic = new Dictionary<string, string>();
        #endregion

        public Dictionary<string, string> LoLimitDic = new Dictionary<string, string>();
        public Dictionary<string, string> HiLimitDic = new Dictionary<string, string>();
        public HashSet<string> IsUpdated = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public string TpLoLimit { set; get; }
        public string TpHiLimit { set; get; }
        public List<string> JobList { get; set; }
        public string ExecutionProfileTestTime { set; get; }
        public string Notes { set; get; }
        public string SoftBin { set; get; }
        public string ErrorMessage { set; get; }
        public string Job;
        public bool Use;
        public string ByStage;
        #endregion

        public BinOutStatusHipListRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Testinstance = "";
            Testnum = "";
            Testname = "";
            Type = "";
            PatternPin = "";
            Lolimit = "";
            Hilimit = "";
            Unit = "";
            TestResult = "";
            Meterirangesetting = "";
            TPuselimit = "";
            Notes = "";
            TpLoLimit = "";
            TpHiLimit = "";
            SoftBin = "";
            ExecutionProfileTestTime = "";
        }
    }

    public class BinOutStatusHipListSheet
    {
        public string Job;

        public string SheetName { get; set; }
        public List<BinOutStatusHipListRow> Rows { get; }
        public Dictionary<string, int> HeaderIndex { get; } = new Dictionary<string, int>();

        public BinOutStatusHipListSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<BinOutStatusHipListRow>();
        }
    }

    public class BinOutStatusHipListReader
    {
        private ExcelWorksheet _excelWorksheet;
        private string _sheetName;
        private BinOutStatusHipListSheet _binOutStatusHipListSheet;

        private const string ConHeaderTestinstance = "TestInstance";
        private const string ConHeaderTestnum = "TestNum";
        private const string ConHeaderTestname = "TestName";
        private const string ConHeaderType = "Type";
        private const string ConHeaderPatternPin = "Pattern/Pin";
        private const string ConHeaderLolimit = "LoLimit";
        private const string ConHeaderHilimit = "HiLimit";
        private const string ConHeaderUnit = "T/P Unit";
        private const string ConHeaderTestResult = "TestResult";
        private const string ConHeaderMeterirangesetting = "MeterIRangeSetting";
        private const string ConHeaderBinOutStatus = "BinOut$";
        private const string ConHeaderEnable = "Enable";
        private const string ConHeaderComment = "Comment";
        private const string ConHeaderTPuselimit = "T/P UseLimit";
        private const string ConHeaderTpLolimit = "T/P LoLimit";
        private const string ConHeaderTpHilimit = "T/P HiLimit";
        private const string ConHeaderBinoutEnable = "BinOut Update";
        private const string ConHeaderUpdatedLoLimit = "LSL Update";
        private const string ConHeaderUpdatedHiLimit = "USL Update";
        private const string ConHeaderNotes = "Notes";
        private const string ConHeaderSwBin = "SwBin";
        private const string ConHeaderExecutionProfileTestTime = "ExecutionProfileTestTime";
        private const string ConHeaderByStage = "By Stage";

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _testinstanceIndex = -1;
        private int _testnumIndex = -1;
        private int _testnameIndex = -1;
        private int _typeIndex = -1;
        private int _patternPinIndex = -1;
        private int _lolimitIndex = -1;
        private int _hilimitIndex = -1;
        private readonly Dictionary<string, int> _lolimitIndexDic = new Dictionary<string, int>();
        private readonly Dictionary<string, int> _hilimitIndexDic = new Dictionary<string, int>();
        private int _unitIndex = -1;
        private int _testresultIndex = -1;
        private int _meterirangesettingIndex = -1;
        private Dictionary<string, int> _statusIndex = new Dictionary<string, int>();
        private readonly Dictionary<string, int> _enableIndex = new Dictionary<string, int>();
        private int _tPuselimitIndex = -1;
        private int _commentIndex = -1;
        private int _tPlolimitIndex = -1;
        private int _tPhilimitIndex = -1;
        private Dictionary<string, int> _binoutEnableIndexDictionary = new Dictionary<string, int>();
        private Dictionary<string, int> _updatedLoLimitDic = new Dictionary<string, int>();
        private Dictionary<string, int> _updatedHiLimitDic = new Dictionary<string, int>();
        private int _swBinIndex = -1;
        private int _notesIndex = -1;
        private int _executionProfileTestTimeIndex = -1;
        private readonly Dictionary<string, int> _itemCountDictionary = new Dictionary<string, int>();
        private int _byStageIndex = -1;

        public HashSet<string> Jobs { get; }

        private readonly Dictionary<string, bool> _headerOptional = new Dictionary<string, bool>
        {
            { "TestInstance", true }, { "TestNum", true }, { "TestName", true }, { "Type", true }, { "Pattern/Pin", true },
            { "LoLimit", true }, { "HiLimit", true }, { "Unit", true }, { "TestResult", true }, { "MeterIRangeSetting", true },
            { "T/P UseLimit", true }, { "Comment", true },
            { "Updated_Lo_Limit", true }, { "Updated_Hi_Limit", true }, { "Notes", true }
        };

        public BinOutStatusHipListReader(HashSet<string> jobs)
        {
            Jobs = jobs;
        }

        public BinOutStatusHipListSheet ReadSheet(ExcelWorksheet worksheet)
        {

            if (worksheet == null)
            {
                return null;
            }

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _binOutStatusHipListSheet = new BinOutStatusHipListSheet(_sheetName);

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            if (!GetHeaderIndex())
            {
                return null;
            }

            _binOutStatusHipListSheet = ReadSheetData();

            return _binOutStatusHipListSheet;
        }

        private BinOutStatusHipListSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                BinOutStatusHipListRow row = new BinOutStatusHipListRow(_sheetName) { RowNum = i };
                PopulateScalarColumns(row, i);
                PopulateDictionaryColumns(row, i);
                PopulateTailColumns(row, i);

                if (!_itemCountDictionary.ContainsKey(row.Testname + "#" + row.PatternPin))
                {
                    _itemCountDictionary.Add(row.Testname + "#" + row.PatternPin, -1);
                }

                _binOutStatusHipListSheet.Rows.Add(row);
            }

            foreach (KeyValuePair<string, int> statusIndexPair in _statusIndex)
            {
                Jobs.Add(Regex.Match(statusIndexPair.Key, @"(?<str>\w+)_" + ConHeaderBinOutStatus).Groups["str"].ToString());
            }
            return _binOutStatusHipListSheet;
        }

        private void PopulateScalarColumns(BinOutStatusHipListRow row, int i)
        {
            if (_testinstanceIndex != -1)
            {
                row.Testinstance = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _testinstanceIndex).Trim();
            }

            if (_testnumIndex != -1)
            {
                row.Testnum = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _testnumIndex).Trim();
            }

            if (_testnameIndex != -1)
            {
                row.Testname = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _testnameIndex).Trim();
            }

            if (_typeIndex != -1)
            {
                row.Type = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _typeIndex).Trim();
            }

            if (_patternPinIndex != -1)
            {
                row.PatternPin = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _patternPinIndex).Trim();
            }

            if (_lolimitIndex != -1)
            {
                row.Lolimit = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _lolimitIndex).Trim();
            }

            if (_hilimitIndex != -1)
            {
                row.Hilimit = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _hilimitIndex).Trim();
            }

            if (_unitIndex != -1)
            {
                row.Unit = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _unitIndex).Trim();
            }

            if (_testresultIndex != -1)
            {
                row.TestResult = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _testresultIndex).Trim();
            }

            if (_meterirangesettingIndex != -1)
            {
                row.Meterirangesetting = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _meterirangesettingIndex).Trim();
            }
        }

        private void PopulateDictionaryColumns(BinOutStatusHipListRow row, int i)
        {
            if (_statusIndex.Count > 0)
            {
                foreach (KeyValuePair<string, int> statusIndexPair in _statusIndex)
                {
                    row.StatusDictionary.Add(statusIndexPair.Key, EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, statusIndexPair.Value).Trim());
                }

            }
            if (_enableIndex.Any())
            {
                foreach (KeyValuePair<string, int> enableIndex in _enableIndex)
                {
                    row.EnableDictionary.Add(enableIndex.Key, EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, enableIndex.Value).Trim());
                }
            }
            if (_tPuselimitIndex != -1)
            {
                row.TPuselimit = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _tPuselimitIndex).Trim();
            }

            if (_commentIndex != -1)
            {
                row.Comment = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _commentIndex).Trim();
            }

            if (_binoutEnableIndexDictionary.Count > 0)
            {
                foreach (KeyValuePair<string, int> binOutIndexPair in _binoutEnableIndexDictionary)
                {
                    row.BinOutEnableDictionary.Add(binOutIndexPair.Key, EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, binOutIndexPair.Value).Trim());
                }
            }
            if (_updatedLoLimitDic.Any())
            {
                foreach (KeyValuePair<string, int> updateLoLimit in _updatedLoLimitDic)
                {
                    string data = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, updateLoLimit.Value).Trim();
                    row.UpdatedLoLimitDic.Add(updateLoLimit.Key, data);
                }
            }

            if (_updatedHiLimitDic.Any())
            {
                foreach (KeyValuePair<string, int> updateHiLimit in _updatedHiLimitDic)
                {
                    string data = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, updateHiLimit.Value).Trim();
                    row.UpdatedHiLimitDic.Add(updateHiLimit.Key, data);
                }
            }
            if (_lolimitIndexDic.Any())
            {
                foreach (KeyValuePair<string, int> loLimit in _lolimitIndexDic)
                {
                    string data = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, loLimit.Value).Trim();
                    row.LoLimitDic.Add(loLimit.Key, data);
                }
            }

            if (_hilimitIndexDic.Any())
            {
                foreach (KeyValuePair<string, int> hiLimit in _hilimitIndexDic)
                {
                    string data = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, hiLimit.Value).Trim();
                    row.HiLimitDic.Add(hiLimit.Key, data);
                }
            }
        }

        private void PopulateTailColumns(BinOutStatusHipListRow row, int i)
        {
            if (_tPlolimitIndex != -1)
            {
                row.TpLoLimit = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _tPlolimitIndex).Trim();
            }

            if (_tPhilimitIndex != -1)
            {
                row.TpHiLimit = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _tPhilimitIndex).Trim();
            }

            if (_notesIndex != -1)
            {
                row.Notes = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _notesIndex).Trim();
            }

            if (_swBinIndex != -1)
            {
                row.SoftBin = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _swBinIndex).Trim();
            }

            if (_executionProfileTestTimeIndex != -1)
            {
                row.ExecutionProfileTestTime = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _executionProfileTestTimeIndex).Trim();
            }

            if (_byStageIndex != -1)
            {
                row.ByStage = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _byStageIndex).Trim();
            }
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = EpplusExtensions.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                MapHeaderColumn(lStrHeader, i);
            }

            foreach (KeyValuePair<string, int> header in _binOutStatusHipListSheet.HeaderIndex)
            {
                if (header.Value == -1 && _headerOptional.ContainsKey(header.Key) && _headerOptional[header.Key])
                {
                    return false;
                }
            }

            return true;
        }

        private void MapHeaderColumn(string lStrHeader, int i)
        {
            if (TryMapPrimaryHeader(lStrHeader, i))
            {
                return;
            }

            TryMapExtendedHeader(lStrHeader, i);
        }

        private bool TryMapPrimaryHeader(string lStrHeader, int i)
        {
            if (lStrHeader.Equals(ConHeaderTestinstance, StringComparison.OrdinalIgnoreCase))
            {
                _testinstanceIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderTestinstance, i);
                return true;
            }
            if (lStrHeader.Equals(ConHeaderTestnum, StringComparison.OrdinalIgnoreCase))
            {
                _testnumIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderTestnum, i);
                return true;
            }
            if (lStrHeader.Equals(ConHeaderTestname, StringComparison.OrdinalIgnoreCase))
            {
                _testnameIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderTestname, i);
                return true;
            }
            if (lStrHeader.Equals(ConHeaderType, StringComparison.OrdinalIgnoreCase))
            {
                _typeIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderType, i);
                return true;
            }
            if (lStrHeader.Equals(ConHeaderPatternPin, StringComparison.OrdinalIgnoreCase))
            {
                _patternPinIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderPatternPin, i);
                return true;
            }
            if (lStrHeader.Equals(ConHeaderLolimit, StringComparison.OrdinalIgnoreCase))
            {
                _lolimitIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderLolimit, i);
                _lolimitIndexDic.Add(Jobs.First(), i);
                return true;
            }
            if (lStrHeader.Equals(ConHeaderHilimit, StringComparison.OrdinalIgnoreCase))
            {
                _hilimitIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderHilimit, i);
                _hilimitIndexDic.Add(Jobs.First(), i);
                return true;
            }
            if (lStrHeader.Equals(ConHeaderUnit, StringComparison.OrdinalIgnoreCase))
            {
                _unitIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderUnit, i);
                return true;
            }
            if (lStrHeader.Equals(ConHeaderTestResult, StringComparison.OrdinalIgnoreCase))
            {
                _testresultIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderTestResult, i);
                return true;
            }
            if (lStrHeader.Equals(ConHeaderMeterirangesetting, StringComparison.OrdinalIgnoreCase))
            {
                _meterirangesettingIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderMeterirangesetting, i);
                return true;
            }
            if (lStrHeader.Equals(ConHeaderExecutionProfileTestTime, StringComparison.OrdinalIgnoreCase))
            {
                _executionProfileTestTimeIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderExecutionProfileTestTime, i);
                return true;
            }

            if (Regex.IsMatch(lStrHeader, "_" + ConHeaderBinOutStatus, RegexOptions.IgnoreCase))
            {
                if (!_statusIndex.ContainsKey(lStrHeader))
                {
                    _statusIndex.Add(lStrHeader, i);
                }
                return true;
            }

            if (Regex.IsMatch(lStrHeader, "_" + ConHeaderEnable, RegexOptions.IgnoreCase))
            {
                string key = !lStrHeader.Equals(ConHeaderEnable, StringComparison.OrdinalIgnoreCase) ?
                            Regex.Replace(lStrHeader, "_" + ConHeaderEnable, "", RegexOptions.IgnoreCase) : Jobs.First();

                if (!_enableIndex.ContainsKey(lStrHeader))
                {
                    _enableIndex.Add(key, i);
                }

                return true;
            }

            if (lStrHeader.Equals(ConHeaderTPuselimit, StringComparison.OrdinalIgnoreCase))
            {
                _tPuselimitIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderTPuselimit, i);
                return true;
            }
            if (lStrHeader.Equals(ConHeaderComment, StringComparison.OrdinalIgnoreCase))
            {
                _commentIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderComment, i);
                return true;
            }

            return false;
        }

        private void TryMapExtendedHeader(string lStrHeader, int i)
        {
            if (Regex.IsMatch(lStrHeader, ConHeaderBinoutEnable, RegexOptions.IgnoreCase))
            {
                string key = Regex.Match(lStrHeader, @"(?<str>\w+)" + "_" + ConHeaderBinoutEnable).Groups["str"].ToString();
                if (!_binoutEnableIndexDictionary.ContainsKey(lStrHeader))
                {
                    _binoutEnableIndexDictionary.Add(key, i);
                }
                return;
            }
            if (Regex.IsMatch(lStrHeader, ConHeaderUpdatedLoLimit, RegexOptions.IgnoreCase))
            {
                if (!_updatedLoLimitDic.ContainsKey(lStrHeader))
                {
                    string key = !lStrHeader.Equals(ConHeaderUpdatedLoLimit, StringComparison.OrdinalIgnoreCase) ?
                                Regex.Replace(lStrHeader, "_" + ConHeaderUpdatedLoLimit, "", RegexOptions.IgnoreCase) : Jobs.First();
                    _updatedLoLimitDic.Add(key, i);
                }
                return;
            }
            if (Regex.IsMatch(lStrHeader, ConHeaderUpdatedHiLimit, RegexOptions.IgnoreCase))
            {
                if (!_updatedHiLimitDic.ContainsKey(lStrHeader))
                {
                    string key = !lStrHeader.Equals(ConHeaderUpdatedHiLimit, StringComparison.OrdinalIgnoreCase) ?
                                Regex.Replace(lStrHeader, "_" + ConHeaderUpdatedHiLimit, "", RegexOptions.IgnoreCase) : Jobs.First();
                    _updatedHiLimitDic.Add(key, i);
                }
                return;
            }
            if (lStrHeader.Equals(ConHeaderNotes, StringComparison.OrdinalIgnoreCase))
            {
                _notesIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderNotes, i);
                return;
            }
            if (lStrHeader.Equals(ConHeaderSwBin, StringComparison.OrdinalIgnoreCase))
            {
                _swBinIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderSwBin, i);
                return;
            }
            if (lStrHeader.Equals(ConHeaderTpLolimit, StringComparison.OrdinalIgnoreCase))
            {
                _tPlolimitIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderTpLolimit, i);
                return;
            }
            if (Regex.IsMatch(lStrHeader, "_" + ConHeaderTpLolimit, RegexOptions.IgnoreCase))
            {
                if (!_lolimitIndexDic.ContainsKey(lStrHeader))
                {
                    string key = !lStrHeader.Equals(ConHeaderTpLolimit, StringComparison.OrdinalIgnoreCase) ?
                                Regex.Replace(lStrHeader, "_" + ConHeaderTpLolimit, "", RegexOptions.IgnoreCase) : Jobs.First();
                    _lolimitIndexDic.Add(key, i);
                }
                return;
            }
            if (lStrHeader.Equals(ConHeaderTpHilimit, StringComparison.OrdinalIgnoreCase))
            {
                _tPhilimitIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderTpHilimit, i);
                return;
            }
            if (Regex.IsMatch(lStrHeader, "_" + ConHeaderTpHilimit, RegexOptions.IgnoreCase))
            {
                if (!_hilimitIndexDic.ContainsKey(lStrHeader))
                {
                    string key = !lStrHeader.Equals(ConHeaderTpHilimit, StringComparison.OrdinalIgnoreCase) ?
                                Regex.Replace(lStrHeader, "_" + ConHeaderTpHilimit, "", RegexOptions.IgnoreCase) : Jobs.First();
                    _hilimitIndexDic.Add(key, i);
                }
                return;
            }
            if (lStrHeader.Equals(ConHeaderByStage, StringComparison.OrdinalIgnoreCase))
            {
                _byStageIndex = i;
                _binOutStatusHipListSheet.HeaderIndex.Add(ConHeaderByStage, i);
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (EpplusExtensions.GetCellValue(_excelWorksheet, i, j).Trim().Equals(ConHeaderTestinstance, StringComparison.OrdinalIgnoreCase))
                    {
                        _startRowNumber = i;
                        return true;
                    }
                }
            }

            return false;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet.Dimension.Start.Row;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }
            return false;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _testinstanceIndex = -1;
            _testnumIndex = -1;
            _testnameIndex = -1;
            _typeIndex = -1;
            _patternPinIndex = -1;
            _lolimitIndex = -1;
            _hilimitIndex = -1;
            _unitIndex = -1;
            _testresultIndex = -1;
            _meterirangesettingIndex = -1;
            _statusIndex = new Dictionary<string, int>();
            _tPuselimitIndex = -1;
            _commentIndex = -1;
            _binoutEnableIndexDictionary = new Dictionary<string, int>();
            _updatedLoLimitDic = new Dictionary<string, int>();
            _updatedHiLimitDic = new Dictionary<string, int>();
            _notesIndex = -1;
            _executionProfileTestTimeIndex = -1;
        }
    }
}
