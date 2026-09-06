using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut
{
    public class SelsrmMappingSheetReader : MySheetReader<SelsrmMappingSheet>
    {
        private const string ConHeaderStage = "Stage";
        private const string ConHeaderBlock = "Block";
        private const string ConHeaderPattern = "Pattern";
        private const string ConHeaderBits = "Bits";
        private const string ConHeaderBitOrder = "bit order";
        private const string ConHeaderAlpha = "Alpha";
        private const string ConHeaderLogicpins = "Logic Pins";
        private const string ConHeaderSrampins = "Sram Pins";
        private const string ConHeaderSelsrm1 = "Selsrm1";
        private const string ConHeaderSelsrm0 = "Selsrm0";
        private const string ConDigSrcAssignment = "DigSrc_Assignment";

        private string _sheetName = "";
        private SelsrmMappingSheet? _selsrmMappingSheet;
        private int _stageIndex = -1;
        private int _blockIndex = -1;
        private int _patternIndex = -1;
        private int _bitsIndex = -1;
        private int _aphaIndex = -1;
        private int _logicpinsIndex = -1;
        private int _srampinsIndex = -1;
        private int _selsrm1Index = -1;
        private int _selsrm0Index = -1;
        private int _digSrcAssignmentIndex = -1;

        public override SelsrmMappingSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;

            _selsrmMappingSheet = new SelsrmMappingSheet(_sheetName);

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition([ConHeaderBlock, ConHeaderStage]))
            {
                return null;
            }

            if (!GetHeaderIndex())
            {
                return null;
            }

            _selsrmMappingSheet = ReadSheetData();

            return _selsrmMappingSheet!;
        }

        private SelsrmMappingSheet ReadSheetData()
        {
            string stage = "";
            string block = "";
            string patteren = "";
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new SelsrmMappingTableRow(_sheetName) { RowNum = i };
                bool needCopy = true;
                if (_bitsIndex != -1)
                {
                    row.Bits = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _bitsIndex).Trim();
                    if (string.IsNullOrEmpty(row.Bits))
                    {
                        needCopy = false;
                    }
                }

                if (_stageIndex != -1)
                {
                    row.Stage = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _stageIndex).Trim();
                    if (!string.IsNullOrEmpty(row.Stage) && needCopy)
                    {
                        stage = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _stageIndex).Trim();
                    }
                    else
                    {
                        if (needCopy)
                        {
                            row.Stage = stage;
                        }
                    }
                }

                if (_blockIndex != -1)
                {
                    row.Block = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _blockIndex).Trim();
                    if (!string.IsNullOrEmpty(row.Block) && needCopy)
                    {
                        block = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _blockIndex).Trim();
                    }
                    else
                    {
                        if (needCopy)
                        {
                            row.Block = block;
                        }
                    }
                }
                if (_patternIndex != -1)
                {
                    row.Pattern = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _patternIndex).Trim();
                    if (!string.IsNullOrEmpty(row.Pattern) && needCopy)
                    {
                        patteren = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _patternIndex).Trim();
                    }
                    else
                    {
                        if (needCopy)
                        {
                            row.Pattern = patteren;
                        }
                    }
                }

                if (_aphaIndex != -1)
                {
                    row.Alpha = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _aphaIndex).Trim();
                }

                if (_logicpinsIndex != -1)
                {
                    row.LogicPins = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _logicpinsIndex).Trim();
                }

                if (_srampinsIndex != -1)
                {
                    row.SramPins = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _srampinsIndex).Trim();
                }

                if (_selsrm1Index != -1)
                {
                    row.Selsrm1 = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _selsrm1Index).Trim();
                }

                if (_selsrm0Index != -1)
                {
                    row.Selsrm0 = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _selsrm0Index).Trim();
                }

                if (_digSrcAssignmentIndex != -1)
                {
                    row.DigSrcAssignment = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, _digSrcAssignmentIndex).Trim();
                }

                _selsrmMappingSheet!.Rows.Add(row);
            }
            return _selsrmMappingSheet!;
        }

        private bool GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string lStrHeader = EpplusExtensions.GetCellValue(ExcelWorksheet, StartRow, i).Trim();
                if (lStrHeader.EqualsIgnoreCase(ConHeaderStage))
                {
                    _stageIndex = i;
                    _selsrmMappingSheet!.HeaderIndex.Add(ConHeaderStage, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderBlock))
                {
                    _blockIndex = i;
                    _selsrmMappingSheet!.HeaderIndex.Add(ConHeaderBlock, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderPattern))
                {
                    _patternIndex = i;
                    _selsrmMappingSheet!.HeaderIndex.Add(ConHeaderPattern, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderBits) ||
                    lStrHeader.EqualsIgnoreCase(ConHeaderBitOrder))
                {
                    _bitsIndex = i;
                    _selsrmMappingSheet!.HeaderIndex.Add(ConHeaderBits, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderAlpha))
                {
                    _aphaIndex = i;
                    _selsrmMappingSheet!.HeaderIndex.Add(ConHeaderAlpha, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderLogicpins))
                {
                    _logicpinsIndex = i;
                    _selsrmMappingSheet!.HeaderIndex.Add(ConHeaderLogicpins, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderSrampins))
                {
                    _srampinsIndex = i;
                    _selsrmMappingSheet!.HeaderIndex.Add(ConHeaderSrampins, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderSelsrm1))
                {
                    _selsrm1Index = i;
                    _selsrmMappingSheet!.HeaderIndex.Add(ConHeaderSelsrm1, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderSelsrm0))
                {
                    _selsrm0Index = i;
                    _selsrmMappingSheet!.HeaderIndex.Add(ConHeaderSelsrm0, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConDigSrcAssignment))
                {
                    _digSrcAssignmentIndex = i;
                    _selsrmMappingSheet!.HeaderIndex.Add(ConDigSrcAssignment, i);
                    continue;
                }
            }
            return true;
        }

        private void Reset()
        {
            StartCol = -1;
            StartRow = -1;
            EndCol = -1;
            EndRow = -1;
            _blockIndex = -1;
            _patternIndex = -1;
            _bitsIndex = -1;
            _logicpinsIndex = -1;
            _srampinsIndex = -1;
            _selsrm1Index = -1;
            _selsrm0Index = -1;
            _digSrcAssignmentIndex = -1;
        }

        public List<Dictionary<string, string>> GenMappingDictionary()
        {
            var dics = new List<Dictionary<string, string>>();
            foreach (SelsrmMappingTableRow row in _selsrmMappingSheet!.Rows)
            {
                var dic = new Dictionary<string, string>
                {
                    { "Block", row.Block }, { "Pattern", row.Pattern }, { "Bits", row.Bits }, { "Logic Pins", row.LogicPins },
                    { "Sram Pins", row.SramPins },
                    { "Selsrm1", row.Selsrm1 },
                    { "Selsrm0", row.Selsrm0 },
                    { "DigSrc_Assignment", row.DigSrcAssignment }
                };
                dics.Add(dic);
            }
            return dics;
        }
    }

    public class SelsrmMappingSheet : MySheet
    {
        #region Properity
        public List<SelsrmMappingTableRow> Rows { get; set; }
        public Dictionary<string, int> HeaderIndex { get; } = [];
        #endregion

        #region Constructor
        public SelsrmMappingSheet(string sheetname)
        {
            SheetName = sheetname;
            Rows = [];
        }

        public SelsrmMappingSheet()
        {
            Rows = [];
        }
        #endregion
    }

    public class SelsrmMappingTableRow : MyRow
    {
        #region Properity
        public string SourceSheetName { get; set; } = string.Empty;
        public string Stage { get; set; }
        public string Block { get; set; }
        public string Pattern { get; set; }
        public string Bits { get; set; }
        public string Alpha { get; set; }
        public string LogicPins { get; set; }
        public string SramPins { get; set; }
        public string Selsrm1 { get; set; }
        public string Selsrm0 { get; set; }
        public string DigSrcAssignment { get; set; }
        #endregion

        #region Constructor
        public SelsrmMappingTableRow()
        {
            Stage = "";
            Block = "";
            Pattern = "";
            Bits = "";
            Alpha = "";
            LogicPins = "";
            SramPins = "";
            Selsrm1 = "";
            Selsrm0 = "";
            DigSrcAssignment = "";
        }

        public SelsrmMappingTableRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Stage = "";
            Block = "";
            Pattern = "";
            Bits = "";
            Alpha = "";
            LogicPins = "";
            SramPins = "";
            Selsrm1 = "";
            Selsrm0 = "";
            DigSrcAssignment = "";
        }
        #endregion
    }
}
