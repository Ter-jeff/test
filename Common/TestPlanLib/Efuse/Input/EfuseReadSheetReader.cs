using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Efuse.Input
{
    public class EfuseReadSheetReader : MySheetReader<EfuseReadSheet>
    {
        private int _testNameColNumber = -1;
        private int _typeColNumber = -1;
        private int _bankColNumber = -1;
        private int _writeReadColNumber = -1;
        private int _purposeColNumber = -1;
        private int _userDefinedColNumber = -1;
        private int _initPatColNumber = -1;
        private int _payloadPatColNumber = -1;
        private readonly List<int> _jobColNumberList = [];

        public override EfuseReadSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            GetDimensions();

            GetFirstHeaderPosition();

            GetHeaders();

            EfuseReadSheet sheet = ReadSheet(sheetName);
            return sheet;
        }

        protected EfuseReadSheet ReadSheet(string sheetName)
        {
            var sheet = new EfuseReadSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new EfuseReadRow
                {
                    RowNum = i,
                    SheetName = sheetName,
                    TestName = _testNameColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _testNameColNumber).Trim(),
                    Type = _typeColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _typeColNumber).Trim(),
                    Bank = _bankColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _bankColNumber).Trim(),
                    WriteRead = _bankColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _writeReadColNumber).Trim(),
                    Purpose = _purposeColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _purposeColNumber).Trim(),
                    UserDefined = _userDefinedColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _userDefinedColNumber).Trim()
                };
                if (_initPatColNumber != -1)
                {
                    foreach (string initPat in ExcelWorksheet.GetCellValue(i, _initPatColNumber).Trim().Split(','))
                    {
                        if (!string.IsNullOrEmpty(initPat))
                        {
                            row.InitList.Add(initPat.Trim());
                        }
                    }
                }
                if (_payloadPatColNumber != -1)
                {
                    foreach (string payload in ExcelWorksheet.GetCellValue(i, _payloadPatColNumber).Trim().Split(','))
                    {
                        if (!string.IsNullOrEmpty(payload))
                        {
                            row.PayloadList.Add(payload.Trim());
                        }
                    }
                }
                foreach (int jobCol in _jobColNumberList)
                {
                    string[] job = ExcelWorksheet.GetCellValue(StartRow, jobCol).Trim().Split('/');
                    foreach (string jobStr in job)
                    {
                        string voltage = ExcelWorksheet.GetCellValue(i, jobCol).Trim();
                        row.JobDictionary.Add(jobStr, voltage);
                    }
                }

                sheet.Rows.Add(row);
            }

            sheet.TestNameColIdx = _testNameColNumber;
            sheet.TypeColIdx = _typeColNumber;
            sheet.BankColIdx = _bankColNumber;
            sheet.WriteReadColIdx = _writeReadColNumber;
            sheet.PurposeColIdx = _purposeColNumber;
            sheet.UserDefinedColIdx = _userDefinedColNumber;
            sheet.InitPatColIdx = _initPatColNumber;
            sheet.PayloadPatColIdx = _payloadPatColNumber;
            sheet.JobColIdx = _jobColNumberList;
            return sheet;
        }

        protected void GetHeaders()
        {
            for (int i = 1; i <= EndCol; i++)
            {
                string lStrHeader = GetCellValue(StartRow, i).ToUpper().Trim();

                if (EfuseHeaderRegex.TestName().IsMatch(lStrHeader))
                {
                    _testNameColNumber = i;
                }
                else if (EfuseHeaderRegex.Type().IsMatch(lStrHeader))
                {
                    _typeColNumber = i;
                }
                else if (EfuseHeaderRegex.Bank().IsMatch(lStrHeader))
                {
                    _bankColNumber = i;
                }
                else if (EfuseHeaderRegex.WriteRead().IsMatch(lStrHeader))
                {
                    _writeReadColNumber = i;
                }
                else if (EfuseHeaderRegex.Purpose().IsMatch(lStrHeader))
                {
                    _purposeColNumber = i;
                }
                else if (EfuseHeaderRegex.UserDefined().IsMatch(lStrHeader))
                {
                    _userDefinedColNumber = i;
                }
                else if (EfuseHeaderRegex.InitPat().IsMatch(lStrHeader))
                {
                    _initPatColNumber = i;
                }
                else if (EfuseHeaderRegex.PayloadPat().IsMatch(lStrHeader))
                {
                    _payloadPatColNumber = i;
                }
                else if (EfuseHeaderRegex.Job().IsMatch(lStrHeader))
                {
                    _jobColNumberList.Add(i);
                }
            }
        }

        protected void GetFirstHeaderPosition()
        {
            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    if (EfuseHeaderRegex.TestName().IsMatch(GetCellValue(i, j).Trim()))
                    {
                        StartRow = i;
                        break;
                    }
                }
            }
        }

        private string GetCellValue(int rowNumber, int columnNumber)
        {
            object value = ExcelWorksheet.Cells[rowNumber, columnNumber].Value;
            return value?.ToString() ?? "";
        }
    }
}
