using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.BinCut
{
    public class PwrbinSeqSheetReader
    {
        private const string ConHeaderSequence = "Sequence";
        private const string ConHeaderBin1 = "Bin1";
        private const string ConHeaderBinx = "BinX";

        private ExcelWorksheet? _excelWorksheet;
        private string _sheetName = "";
        private PwrbinSeqSheet? _pwrbinSeqSheet;
        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _sequenceIndex = -1;
        private int _bin1Index = -1;
        private int _binxIndex = -1;

        public PwrbinSeqSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            _excelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;

            _pwrbinSeqSheet = new PwrbinSeqSheet(_sheetName);

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

            _pwrbinSeqSheet = ReadSheetData();

            _pwrbinSeqSheet!.Bin1List = [.. _pwrbinSeqSheet!.Rows.Select(x => x.Bin1).Where(y => !string.IsNullOrEmpty(y))];
            _pwrbinSeqSheet!.BinXList = [.. _pwrbinSeqSheet!.Rows.Select(x => x.BinX).Where(y => !string.IsNullOrEmpty(y))];
            return _pwrbinSeqSheet!;
        }

        private PwrbinSeqSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new PwrbinSeqRow(_sheetName) { RowNum = i };
                if (_sequenceIndex != -1)
                {
                    row.Sequence = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, _sequenceIndex).Trim();
                }

                if (_bin1Index != -1)
                {
                    row.Bin1 = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, _bin1Index).Trim();
                }

                if (_binxIndex != -1)
                {
                    row.BinX = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, _binxIndex).Trim();
                }

                _pwrbinSeqSheet!.Rows.Add(row);
            }
            return _pwrbinSeqSheet!;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string header = EpplusExtensions.GetCellValue(_excelWorksheet!, _startRowNumber, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderSequence))
                {
                    _sequenceIndex = i;
                    _pwrbinSeqSheet!.HeaderIndex.Add(ConHeaderSequence, i);
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderBin1))
                {
                    _bin1Index = i;
                    _pwrbinSeqSheet!.HeaderIndex.Add(ConHeaderBin1, i);
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderBinx))
                {
                    _binxIndex = i;
                    _pwrbinSeqSheet!.HeaderIndex.Add(ConHeaderBinx, i);
                    continue;
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (EpplusExtensions.GetCellValue(_excelWorksheet!, i, j).Trim().EqualsIgnoreCase(ConHeaderSequence))
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
            if (_excelWorksheet!.Dimension != null)
            {
                _startColNumber = _excelWorksheet!.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet!.Dimension.Start.Row;
                _endColNumber = _excelWorksheet!.Dimension.End.Column;
                _endRowNumber = _excelWorksheet!.Dimension.End.Row;
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
            _sequenceIndex = -1;
            _bin1Index = -1;
            _binxIndex = -1;
        }

        public List<Dictionary<string, string>> GenMappingDictionary()
        {
            var dics = new List<Dictionary<string, string>>();
            foreach (PwrbinSeqRow row in _pwrbinSeqSheet!.Rows)
            {
                var dic = new Dictionary<string, string>
                {
                    { "Sequence", row.Sequence }, { "Bin1", row.Bin1 }, { "BinX", row.BinX }
                };
                dics.Add(dic);
            }
            return dics;
        }
    }

    public class PwrbinSeqSheet(string sheetname)
    {
        #region Properity
        public string SheetName { get; set; } = sheetname;
        public List<PwrbinSeqRow> Rows { get; } = [];
        public Dictionary<string, int> HeaderIndex { get; } = [];
        public List<string> Bin1List = [];
        public List<string> BinXList = [];

        #endregion
        #region Constructor
        #endregion

        public string GetPwrbinSeqSheet(int seq, int binNum)
        {
            if (seq == -1)
            {
                return "";
            }

            if (binNum == 1 && seq < Bin1List.Count)
            {
                return Bin1List[seq];
            }

            if (binNum == 2 && seq < BinXList.Count)
            {
                return BinXList[seq];
            }

            return "";
        }

        public string GetFinalPwrbinSeqSheet(int binNum)
        {
            if (binNum == 1)
            {
                return Bin1List.Last();
            }

            if (binNum == 2)
            {
                return BinXList.Last();
            }

            return "";
        }
    }

    public class PwrbinSeqRow
    {
        #region Properity
        public string SourceSheetName { set; get; } = "";
        public int RowNum { get; set; }
        public string Sequence { set; get; }
        public string Bin1 { set; get; }
        public string BinX { set; get; }
        #endregion

        #region Constructor
        public PwrbinSeqRow()
        {
            Sequence = "";
            Bin1 = "";
            BinX = "";
        }

        public PwrbinSeqRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Sequence = "";
            Bin1 = "";
            BinX = "";
        }
        #endregion
    }
}
