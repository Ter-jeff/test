using System;

using CommonLib.Extension;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.Rtos.Input
{
    public class InputFormat
    {
        private readonly ExcelWorksheet _worksheet;
        public const string Block = "Block";
        public const string Mode = "Mode";
        public const string TestName = "TestName";

        public InputFormat(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public int GetFormatNumber()
        {
            if (_worksheet.Dimension == null && _worksheet.Dimension == null)
            {
                return -1;
            }

            int formatNumber = -1;

            for (int i = 1; i < _worksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j < _worksheet.Dimension.End.Column; j++)
                {
                    string cellContent = EpplusExtensions.GetCellValue(_worksheet, i, j);
                    if (cellContent.Equals(Block, StringComparison.OrdinalIgnoreCase) ||
                        cellContent.Equals(Mode, StringComparison.OrdinalIgnoreCase))
                    {
                        formatNumber = 0;
                    }

                    if (cellContent.Equals(TestName, StringComparison.OrdinalIgnoreCase))
                    {
                        formatNumber = 1;
                    }
                }
                if (formatNumber != -1)
                {
                    break;
                }
            }
            return formatNumber;
        }
    }
}
