using System.Collections.Generic;
using System.IO;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class ReadBinTableSheet : IgxlSheetReader<BinTableSheet>
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        public override EnumSheetType SupportedType => EnumSheetType.DTBintablesSheet;

        private readonly List<string> _headList = [];
        private int _commentIdx;
        private int _endItemIndex;

        public override BinTableSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var subFlowSheet = new BinTableSheet(excelWorksheet);
            int maxRowCount = excelWorksheet.Dimension.End.Row;
            int maxColumnCount = excelWorksheet.Dimension.End.Column;
            _endItemIndex = maxColumnCount;

            _headList.Clear();
            for (int i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                string value = GetMergeCellValue(excelWorksheet, StartRowIndex, i);
                string head = value.Trim();
                _headList.Add(head);
            }

            _commentIdx = _headList.FindIndex(x => x.EqualsIgnoreCase("Comment"));
            if (_commentIdx != 0)
            {
                _commentIdx += StartColumnIndex;
                _endItemIndex = _commentIdx;
            }

            bool backup = false;
            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                BinTableRow binRow = GetBinTableRow(excelWorksheet, i);
                if (string.IsNullOrEmpty(binRow.Name))
                {
                    backup = true;
                }
                binRow.IsBackup = backup;
                subFlowSheet.AddRow(binRow);
            }
            return subFlowSheet;
        }

        public override BinTableSheet ReadSheet(Stream stream, string sheetName)
        {
            var sheet = new BinTableSheet(sheetName);
            bool isBackup = false;
            int i = 1;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine()!;
                    if (i > StartRowIndex)
                    {
                        BinTableRow binRow = GetBinTableRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(binRow.Name))
                        {
                            isBackup = true;
                            continue;
                        }

                        binRow.IsBackup = isBackup;
                        sheet.AddRow(binRow);
                    }

                    i++;
                }
            }

            return sheet;
        }

        private BinTableRow GetBinTableRow(string line, string sheetName, int row)
        {
            string[] arr = line.Split('\t');
            var binTableRow = new BinTableRow
            {
                RowNum = row,
                SheetName = sheetName
            };
            int index = StartColumnIndex - 1;
            string content = GetCellValue(arr, 0);
            binTableRow.ColumnA = content;
            content = GetCellValue(arr, index);
            binTableRow.Name = content;
            index++;
            content = GetCellValue(arr, index);
            binTableRow.ItemList = content;
            index++;
            content = GetCellValue(arr, index);
            binTableRow.Op = content;
            index++;
            content = GetCellValue(arr, index);
            binTableRow.Sort = content;
            index++;
            content = GetCellValue(arr, index);
            binTableRow.Bin = content;
            index++;
            content = GetCellValue(arr, index);
            binTableRow.Result = content;
            index++;

            var items = new List<string>();
            for (int i = 0; i < 80; i++)
            {
                content = GetCellValue(arr, index + i);
                if (!string.IsNullOrEmpty(content))
                {
                    items.Add(content);
                }
            }
            binTableRow.Items = items;

            return binTableRow;
        }

        private BinTableRow GetBinTableRow(ExcelWorksheet excelWorksheet, int row)
        {
            var binRow = new BinTableRow();
            int index = 1;
            binRow.RowNum = row;

            string content = GetCellText(excelWorksheet, row, index);
            binRow.ColumnA = content;

            index++;
            content = GetCellText(excelWorksheet, row, index);
            binRow.Name = content;

            index++;
            content = GetCellText(excelWorksheet, row, index);
            binRow.ItemList = content;

            index++;
            content = GetCellText(excelWorksheet, row, index);
            binRow.Op = content;

            index++;
            content = GetCellText(excelWorksheet, row, index);
            binRow.Sort = content;

            index++;
            content = GetCellText(excelWorksheet, row, index);
            binRow.Bin = content;

            index++;
            content = GetCellText(excelWorksheet, row, index);
            binRow.Result = content;
            index++;

            var items = new List<string>();
            for (int i = index; i < _endItemIndex; i++)
            {
                if (!string.IsNullOrEmpty(GetCellText(excelWorksheet, row, i)))
                {
                    items.Add(GetCellText(excelWorksheet, row, i));
                }
            }

            binRow.Items = items;

            for (int i = index; i < _endItemIndex; i++)
            {
                binRow.ItemsWithIndex.Add(i - index, GetCellText(excelWorksheet, row, i));
            }

            binRow.Comment = GetCellText(excelWorksheet, row, _commentIdx == 0 ? _endItemIndex - 1 : _commentIdx);

            return binRow;
        }
    }
}
