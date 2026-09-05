using System.Collections.Generic;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Basic
{
    public partial class UfInstanceTableReader : MySheetReader<UfInstanceTable>
    {
        private const string ConTestName = "Test Name";
        private const string ConDcSpecs = "DC Specs";
        private const string ConAcSpecs = "AC Specs";
        private const string ConTimeset = "Time Sets";
        private const string ConPinLevels = "Pin Levels";
        private const string ConModule = "Module";
        private const string ConArgList = "ArgList";
        private const string ConArg = @"Arg\d+";

        [GeneratedRegex(ConArg, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();

        private int _indexTestName = -1;
        private int _indexDcSpecs = -1;
        private int _indexAcSpecs = -1;
        private int _indexTimeset = -1;
        private int _indexModule = -1;
        private int _indexArgList = -1;
        private int _indexPinLevels = -1;
        private readonly List<int> _indexArg = [];
        private readonly UfInstanceTable _ufInstanceTable = new();

        public override UfInstanceTable? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            ReadSheet();

            if (_ufInstanceTable.CheckPowerUpExist())
            {
                _ = new List<int>();
                List<int> powerUpRowNumList = _ufInstanceTable.GetPowerUpInstanceRowNum();
                foreach (int num in powerUpRowNumList)
                {
                    _ufInstanceTable.AddError(BasicErrorType.E_DefaultPowerUpInstanceExist_01, excelWorksheet.Name, num, _indexTestName, "PowerUp instance exist in UF_Instance sheet, Autogen will not generate default power up instance \"PowerUp\", please make sure to put this instance in Flow_Main");
                }
            }

            return _ufInstanceTable;
        }

        private void ReadSheet()
        {
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var rowTmp = new UfInstanceRow();
                if (_indexTestName != -1)
                {
                    rowTmp.TestName = ExcelWorksheet.GetCellValue(i, _indexTestName).Trim();
                }
                if (_indexDcSpecs != -1)
                {
                    rowTmp.DcSpec = ExcelWorksheet.GetCellValue(i, _indexDcSpecs).Trim();
                }
                if (_indexAcSpecs != -1)
                {
                    rowTmp.AcSpec = ExcelWorksheet.GetCellValue(i, _indexAcSpecs).Trim();
                }
                if (_indexPinLevels != -1)
                {
                    rowTmp.PinLevels = ExcelWorksheet.GetCellValue(i, _indexPinLevels).Trim();
                }
                if (_indexTimeset != -1)
                {
                    rowTmp.TimeSet = ExcelWorksheet.GetCellValue(i, _indexTimeset).Trim();
                }
                if (_indexModule != -1)
                {
                    rowTmp.Module = ExcelWorksheet.GetCellValue(i, _indexModule).Trim();
                }
                if (_indexArgList != -1)
                {
                    rowTmp.ArgList = ExcelWorksheet.GetCellValue(i, _indexArgList).Trim();
                }
                var argTmp = new List<string>();
                foreach (int index in _indexArg)
                {
                    argTmp.Add(ExcelWorksheet.GetCellValue(i, index).Trim());
                }
                rowTmp.Arg = argTmp;
                _ufInstanceTable.Rows.Add(rowTmp);
            }
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConTestName))
                {
                    _indexTestName = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConModule))
                {
                    _indexModule = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConDcSpecs))
                {
                    _indexDcSpecs = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConAcSpecs))
                {
                    _indexAcSpecs = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConPinLevels))
                {
                    _indexPinLevels = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConTimeset))
                {
                    _indexTimeset = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConArgList))
                {
                    _indexArgList = i;
                    continue;
                }
                if (MyRegex().IsMatch(header))
                {
                    _indexArg.Add(i);
                    continue;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow;
            int colNum = EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConTestName))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
