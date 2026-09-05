using System.Text.RegularExpressions;

using OfficeOpenXml;

using TestPlanLib;

namespace RF_PatternTool.PatternGen
{
    public class TestSequenceReader : MySheetReader
    {
        private const string ConHeaderId = "ID";
        private const string ConHeaderTag = "Tag";
        private const string ConHeaderLoop = "Loop";
        private const string ConHeaderCmd = "Cmd";
        private const string ConHeaderApbSlave = "APB Slave";
        private const string ConHeaderRegisterAddressPin1 = "Register Address / Pin 1";
        private const string ConHeaderFieldNamePin2 = "Field Name / Pin 2";
        private const string ConHeaderBitWidths = "Bit Widths";
        private const string ConHeaderValue1 = "Value 1";
        private const string ConHeaderValue2 = "Value 2";
        private const string ConHeaderValue3 = "Value 3";
        private const string ConHeaderValue4 = "Value 4";
        private const string ConHeaderValue5 = "Value 5";
        private const string ConHeaderComment = "Comment";

        private int _indexId = -1;
        private int _indexTag = -1;
        private int _indexLoop = -1;
        private int _indexCmd = -1;
        private int _indexApbSlave = -1;
        private int _indexRegisterAddressPin1 = -1;
        private int _indexFieldNamePin2 = -1;
        private int _indexBitWidths = -1;
        private int _indexValue1 = -1;
        private int _indexValue2 = -1;
        private int _indexValue3 = -1;
        private int _indexValue4 = -1;
        private int _indexValue5 = -1;
        private int _indexComment = -1;

        public TestSequenceSheet ReadSheet(ExcelWorksheet worksheet)
        {
            string sheetName = worksheet.Name;

            var ruleSheet = new TestSequenceSheet(sheetName);

            ExcelWorksheet = worksheet;

            if (!GetDimensions())
            {
                return (TestSequenceSheet)ruleSheet.DimensionError();
            }

            if (!GetFirstHeaderPosition())
            {
                return (TestSequenceSheet)ruleSheet.FirstHeaderError(ConHeaderId);
            }

            GetHeaderIndex();

            ruleSheet = ReadSheet(sheetName);

            return ruleSheet;
        }

        private void ReadRowValues(TestSequenceRow row, int i)
        {
            if (_indexId != -1)
            {
                row.Id = GetMergedCellValue(ExcelWorksheet, i, _indexId).Trim();
            }

            if (_indexTag != -1)
            {
                row.Tag = GetMergedCellValue(ExcelWorksheet, i, _indexTag).Trim();
            }

            if (_indexLoop != -1)
            {
                row.Loop = GetMergedCellValue(ExcelWorksheet, i, _indexLoop).Trim();
            }

            if (_indexCmd != -1)
            {
                row.Cmd = GetMergedCellValue(ExcelWorksheet, i, _indexCmd).Trim();
            }

            if (_indexApbSlave != -1)
            {
                row.ApbSlave = GetMergedCellValue(ExcelWorksheet, i, _indexApbSlave).Trim();
            }

            if (_indexRegisterAddressPin1 != -1)
            {
                row.RegisterAddressPin1 = GetMergedCellValue(ExcelWorksheet, i, _indexRegisterAddressPin1).Trim();
            }

            if (_indexFieldNamePin2 != -1)
            {
                row.FieldNamePin2 = GetMergedCellValue(ExcelWorksheet, i, _indexFieldNamePin2).Trim();
            }

            if (_indexBitWidths != -1)
            {
                row.BitWidths = GetMergedCellValue(ExcelWorksheet, i, _indexBitWidths).Trim();
            }

            if (_indexValue1 != -1)
            {
                row.Value1 = GetMergedCellValue(ExcelWorksheet, i, _indexValue1).Trim();
            }

            if (_indexValue2 != -1)
            {
                row.Value2 = GetMergedCellValue(ExcelWorksheet, i, _indexValue2).Trim();
            }

            if (_indexValue3 != -1)
            {
                row.Value3 = GetMergedCellValue(ExcelWorksheet, i, _indexValue3).Trim();
            }

            if (_indexValue4 != -1)
            {
                row.Value4 = GetMergedCellValue(ExcelWorksheet, i, _indexValue4).Trim();
            }

            if (_indexValue5 != -1)
            {
                row.Value5 = GetMergedCellValue(ExcelWorksheet, i, _indexValue5).Trim();
            }

            if (_indexComment != -1)
            {
                row.Comment = GetMergedCellValue(ExcelWorksheet, i, _indexComment).Trim();
            }
        }

        private TestSequenceSheet ReadSheet(string sheetName)
        {
            var ruleSheet = new TestSequenceSheet(sheetName);
            bool isLoopRecord = false;
            int loopNum = 0;
            var loopInfo = new List<TestSequenceRow>();
            for (int i = StartRowNumber + 1; i <= EndRowNumber; i++)
            {
                var row = new TestSequenceRow(sheetName);
                row.RowNum = i;
                ReadRowValues(row, i);

                if (!string.IsNullOrEmpty(row.Loop))
                {
                    string regloop = @"Loop\((?<num>\d+)\)";
                    if (isLoopRecord && Regex.IsMatch(row.Loop, "LoopEnd", RegexOptions.IgnoreCase))
                    {
                        for (int j = 1; j < loopNum; j++)
                        {
                            ruleSheet.Rows.AddRange(loopInfo);
                        }
                        loopInfo.Clear();
                        isLoopRecord = false;
                    }
                    if (Regex.IsMatch(row.Loop, regloop, RegexOptions.IgnoreCase))
                    {
                        isLoopRecord = true;
                        loopNum = int.Parse(Regex.Match(row.Loop, regloop, RegexOptions.IgnoreCase).Groups["num"].Value);
                    }

                }

                //if (!string.IsNullOrEmpty(row.Cmd) && row.Cmd.Equals("wait", StringComparison.OrdinalIgnoreCase))
                {
                    //ruleSheet.Rows.Last(p=>!string.IsNullOrEmpty(p.Cmd)).Wait = double.Parse(Regex.Replace(row.Value1,"0x","",RegexOptions.IgnoreCase))*1e-6;
                }
                //else
                if (!string.IsNullOrEmpty(row.Cmd))
                {
                    ruleSheet.Rows.Add(row);
                    if (isLoopRecord)
                    {
                        loopInfo.Add(row);
                    }
                }



            }

            ruleSheet.IndexId = _indexId;
            ruleSheet.IndexTag = _indexTag;
            ruleSheet.IndexLoop = _indexLoop;
            ruleSheet.IndexCmd = _indexCmd;
            ruleSheet.IndexApbSlave = _indexApbSlave;
            ruleSheet.IndexRegisterAddressPin1 = _indexRegisterAddressPin1;
            ruleSheet.IndexFieldNamePin2 = _indexFieldNamePin2;
            ruleSheet.IndexBitWidths = _indexBitWidths;
            ruleSheet.IndexValue1 = _indexValue1;
            ruleSheet.IndexValue2 = _indexValue2;
            ruleSheet.IndexValue3 = _indexValue3;
            ruleSheet.IndexValue4 = _indexValue4;
            ruleSheet.IndexValue5 = _indexValue5;
            ruleSheet.IndexComment = _indexComment;

            return ruleSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartColNumber; i <= EndColNumber; i++)
            {
                string header = GetCellValue(ExcelWorksheet, StartRowNumber, i).Trim();
                if (header.Equals(ConHeaderId, StringComparison.OrdinalIgnoreCase))
                {
                    _indexId = i;
                    continue;
                }
                if (header.Equals(ConHeaderTag, StringComparison.OrdinalIgnoreCase))
                {
                    _indexTag = i;
                    continue;
                }
                if (header.Equals(ConHeaderLoop, StringComparison.OrdinalIgnoreCase))
                {
                    _indexLoop = i;
                    continue;
                }
                if (header.Equals(ConHeaderCmd, StringComparison.OrdinalIgnoreCase))
                {
                    _indexCmd = i;
                    continue;
                }
                if (header.Equals(ConHeaderApbSlave, StringComparison.OrdinalIgnoreCase))
                {
                    _indexApbSlave = i;
                    continue;
                }
                if (header.Equals(ConHeaderRegisterAddressPin1, StringComparison.OrdinalIgnoreCase))
                {
                    _indexRegisterAddressPin1 = i;
                    continue;
                }
                if (header.Equals(ConHeaderFieldNamePin2, StringComparison.OrdinalIgnoreCase))
                {
                    _indexFieldNamePin2 = i;
                    continue;
                }
                if (header.Equals(ConHeaderBitWidths, StringComparison.OrdinalIgnoreCase))
                {
                    _indexBitWidths = i;
                    continue;
                }
                if (header.Equals(ConHeaderValue1, StringComparison.OrdinalIgnoreCase))
                {
                    _indexValue1 = i;
                    continue;
                }
                if (header.Equals(ConHeaderValue2, StringComparison.OrdinalIgnoreCase))
                {
                    _indexValue2 = i;
                    continue;
                }
                if (header.Equals(ConHeaderValue3, StringComparison.OrdinalIgnoreCase))
                {
                    _indexValue3 = i;
                    continue;
                }
                if (header.Equals(ConHeaderValue4, StringComparison.OrdinalIgnoreCase))
                {
                    _indexValue4 = i;
                    continue;
                }
                if (header.Equals(ConHeaderValue5, StringComparison.OrdinalIgnoreCase))
                {
                    _indexValue5 = i;
                    continue;
                }
                if (header.Equals(ConHeaderComment, StringComparison.OrdinalIgnoreCase))
                {
                    _indexComment = i;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRowNumber > 10 ? 10 : EndRowNumber;
            int colNum = EndColNumber > 10 ? 10 : EndColNumber;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (GetCellValue(ExcelWorksheet, i, j).Trim().Equals(ConHeaderId, StringComparison.OrdinalIgnoreCase))
                    {
                        StartRowNumber = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public class TestSequenceSheet : MySheet
    {
        public List<TestSequenceRow> Rows { set; get; }
        public string FileName { get; set; }
        public int IndexId = -1;
        public int IndexTag = -1;
        public int IndexLoop = -1;
        public int IndexCmd = -1;
        public int IndexApbSlave = -1;
        public int IndexRegisterAddressPin1 = -1;
        public int IndexFieldNamePin2 = -1;
        public int IndexBitWidths = -1;
        public int IndexValue1 = -1;
        public int IndexValue2 = -1;
        public int IndexValue3 = -1;
        public int IndexValue4 = -1;
        public int IndexValue5 = -1;
        public int IndexComment = -1;

        #region Constructor
        public TestSequenceSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<TestSequenceRow>();
        }
        #endregion
    }

    public class TestSequenceRow : MyRow
    {
        public string Id { set; get; }
        public string Tag { set; get; }
        public string Loop { set; get; }
        public string Cmd { set; get; }
        public string ApbSlave { set; get; }
        public string RegisterAddressPin1 { set; get; }
        public string FieldNamePin2 { set; get; }
        public string BitWidths { set; get; }
        public string Value1 { set; get; }
        public string Value2 { set; get; }
        public string Value3 { set; get; }
        public string Value4 { set; get; }
        public string Value5 { set; get; }
        public double Wait { set; get; }
        public string Comment { set; get; }

        #region Constructor
        public TestSequenceRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion

        public string GetRegisterName()
        {
            return string.Format("{0}__{1}", ApbSlave, RegisterAddressPin1);
        }
    }
}
