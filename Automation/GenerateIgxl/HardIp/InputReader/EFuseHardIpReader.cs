using System;
using System.Collections.Generic;

using CommonLib.Extension;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.HardIp.InputReader
{
    public class EFuseHardIpRow
    {
        #region Properity
        public string SourceSheetName { set; get; }

        public int RowNum { get; set; }

        public string Block { set; get; }

        public string Mode { set; get; }

        public string Bank { set; get; }

        public string Efusedefinition { set; get; }

        public string Width { set; get; }

        public string Pattern { set; get; }

        public string Referencedblock { set; get; }

        public string Referencedmode { set; get; }

        public string Referencedsendbitname { set; get; }

        public string Job { set; get; }

        public string Comment { set; get; }

        #endregion

        #region Constructor
        public EFuseHardIpRow()
        {
            Block = "";
            Mode = "";
            Bank = "";
            Efusedefinition = "";
            Width = "";
            Pattern = "";
            Referencedblock = "";
            Referencedmode = "";
            Referencedsendbitname = "";
            Job = "";
            Comment = "";
        }

        public EFuseHardIpRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Block = "";
            Mode = "";
            Bank = "";
            Efusedefinition = "";
            Width = "";
            Pattern = "";
            Referencedblock = "";
            Referencedmode = "";
            Referencedsendbitname = "";
            Job = "";
            Comment = "";
        }
        #endregion
    }

    public class EFuseHardIpSheet
    {
        #region Properity
        public string SheetName { get; set; }
        public List<EFuseHardIpRow> Rows { get; }
        public Dictionary<string, int> HeaderIndex { get; } = new Dictionary<string, int>();

        #endregion

        #region Constructor
        public EFuseHardIpSheet(string sheetname)
        {
            SheetName = sheetname;
            Rows = new List<EFuseHardIpRow>();
        }
        #endregion
    }

    public class EFuseHardIpReader
    {
        private ExcelWorksheet _excelWorksheet;
        private string _sheetName;
        private EFuseHardIpSheet _eFuseHardIpTableSheet;

        private const string ConHeaderBlock = "Block";
        private const string ConHeaderMode = "Mode";
        private const string ConHeaderBank = "Bank";
        private const string ConHeaderEfusedefinition = "eFuse Definition";
        private const string ConHeaderEfusedefinitionWidth = "Width";
        private const string ConHeaderPattern = "Pattern";
        private const string ConHeaderReferencedblock = "Referenced Block";
        private const string ConHeaderReferencedmode = "Referenced Mode";
        private const string ConHeaderReferencedsendbitname = "Referenced Send Bit Name";
        private const string ConHeaderInsert = "Insert";
        private const string ConHeaderJob = "Job";
        private const string ConHeaderComment = "Comment";

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _blockIndex = -1;
        private int _modeIndex = -1;
        private int _bankIndex = -1;
        private int _efusedefinitionIndex = -1;
        private int _efusedefinitionWidthIndex = -1;
        private int _patternIndex = -1;
        private int _referencedblockIndex = -1;
        private int _referencedmodeIndex = -1;
        private int _referencedsendbitnameIndex = -1;
        private int _job = -1;
        private int _commentIndex = -1;
        private readonly Dictionary<string, bool> _headerOptional = new Dictionary<string, bool>
        {
            { "Block", true }, { "Mode", true }, { "Bank", true }, { "eFuse Definition", true }, { "Pattern", true }, { "Referenced Block", true }, { "Referenced Mode", true }, { "Referenced Send Bit Name", true }, { "Comment", true }
        };

        public EFuseHardIpSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return null;
            }

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _eFuseHardIpTableSheet = new EFuseHardIpSheet(_sheetName);

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

            _eFuseHardIpTableSheet = ReadSheetData();

            return _eFuseHardIpTableSheet;
        }

        private EFuseHardIpSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                EFuseHardIpRow row = new EFuseHardIpRow(_sheetName) { RowNum = i };
                if (_blockIndex != -1)
                {
                    row.Block = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _blockIndex).Trim();
                }

                if (_modeIndex != -1)
                {
                    row.Mode = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _modeIndex).Trim();
                }

                if (_bankIndex != -1)
                {
                    row.Bank = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _bankIndex).Trim();
                }

                if (_efusedefinitionIndex != -1)
                {
                    row.Efusedefinition = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _efusedefinitionIndex).Trim();
                }

                if (_efusedefinitionWidthIndex != -1)
                {
                    row.Width = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _efusedefinitionWidthIndex).Trim();
                }

                if (_patternIndex != -1)
                {
                    row.Pattern = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _patternIndex).Trim();
                }

                if (_referencedblockIndex != -1)
                {
                    row.Referencedblock = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _referencedblockIndex).Trim();
                }

                if (_referencedmodeIndex != -1)
                {
                    row.Referencedmode = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _referencedmodeIndex).Trim();
                }

                if (_referencedsendbitnameIndex != -1)
                {
                    row.Referencedsendbitname = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _referencedsendbitnameIndex).Trim();
                }

                if (_job != -1)
                {
                    row.Job = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _job).Trim();
                }

                if (_commentIndex != -1)
                {
                    row.Comment = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _commentIndex).Trim();
                }

                _eFuseHardIpTableSheet.Rows.Add(row);
            }
            return _eFuseHardIpTableSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = EpplusExtensions.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(ConHeaderBlock, StringComparison.OrdinalIgnoreCase))
                {
                    _blockIndex = i;
                    _eFuseHardIpTableSheet.HeaderIndex.Add(ConHeaderBlock, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderMode, StringComparison.OrdinalIgnoreCase))
                {
                    _modeIndex = i;
                    _eFuseHardIpTableSheet.HeaderIndex.Add(ConHeaderMode, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderBank, StringComparison.OrdinalIgnoreCase))
                {
                    _bankIndex = i;
                    _eFuseHardIpTableSheet.HeaderIndex.Add(ConHeaderBank, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderEfusedefinition, StringComparison.OrdinalIgnoreCase))
                {
                    _efusedefinitionIndex = i;
                    _eFuseHardIpTableSheet.HeaderIndex.Add(ConHeaderEfusedefinition, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderEfusedefinitionWidth, StringComparison.OrdinalIgnoreCase))
                {
                    _efusedefinitionWidthIndex = i;
                    _eFuseHardIpTableSheet.HeaderIndex.Add(ConHeaderEfusedefinitionWidth, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderPattern, StringComparison.OrdinalIgnoreCase))
                {
                    _patternIndex = i;
                    _eFuseHardIpTableSheet.HeaderIndex.Add(ConHeaderPattern, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderReferencedblock, StringComparison.OrdinalIgnoreCase))
                {
                    _referencedblockIndex = i;
                    _eFuseHardIpTableSheet.HeaderIndex.Add(ConHeaderReferencedblock, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderReferencedmode, StringComparison.OrdinalIgnoreCase))
                {
                    _referencedmodeIndex = i;
                    _eFuseHardIpTableSheet.HeaderIndex.Add(ConHeaderReferencedmode, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderReferencedsendbitname, StringComparison.OrdinalIgnoreCase))
                {
                    _referencedsendbitnameIndex = i;
                    _eFuseHardIpTableSheet.HeaderIndex.Add(ConHeaderReferencedsendbitname, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderInsert, StringComparison.OrdinalIgnoreCase) || lStrHeader.Equals(ConHeaderJob, StringComparison.OrdinalIgnoreCase))
                {
                    _job = i;
                    _eFuseHardIpTableSheet.HeaderIndex.Add(ConHeaderJob, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderComment, StringComparison.OrdinalIgnoreCase))
                {
                    _commentIndex = i;
                    _eFuseHardIpTableSheet.HeaderIndex.Add(ConHeaderComment, i);
                }
            }

            foreach (KeyValuePair<string, int> header in _eFuseHardIpTableSheet.HeaderIndex)
            {
                if (header.Value == -1 && _headerOptional.ContainsKey(header.Key) && _headerOptional[header.Key])
                {
                    return false;
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
                    if (EpplusExtensions.GetCellValue(_excelWorksheet, i, j).Trim().Equals(ConHeaderBlock, StringComparison.OrdinalIgnoreCase))
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
            _blockIndex = -1;
            _modeIndex = -1;
            _bankIndex = -1;
            _efusedefinitionIndex = -1;
            _patternIndex = -1;
            _referencedblockIndex = -1;
            _referencedmodeIndex = -1;
            _referencedsendbitnameIndex = -1;
            _commentIndex = -1;
        }
    }
}
