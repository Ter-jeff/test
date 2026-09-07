using System.Collections.Generic;
using System.IO;

using CommonLib.Utility;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class ReadLevelSheet : IgxlSheetReader<LevelSheet>
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        public override EnumSheetType SupportedType => EnumSheetType.DTLevelSheet;

        private readonly List<string> _headList = [];

        public override LevelSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var levelSheet = new LevelSheet(excelWorksheet);
            int maxRowCount = excelWorksheet.Dimension.End.Row;
            int maxColumnCount = excelWorksheet.Dimension.End.Column;

            _headList.Clear();
            for (int i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                string value = GetMergeCellValue(excelWorksheet, StartRowIndex, i);
                string head = value.Trim();
                _headList.Add(head);
            }

            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                LevelRow levelRow = GetLevelRow(excelWorksheet, i);
                levelSheet.Rows.Add(levelRow);
            }

            return levelSheet;
        }

        public override LevelSheet ReadSheet(Stream stream, string sheetName)
        {
            var sheet = new LevelSheet(sheetName);
            bool isBackup = false;
            int i = 1;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine()!;
                    if (i > StartRowIndex)
                    {
                        LevelRow levelRow = GetLevelRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(levelRow.PinName))
                        {
                            isBackup = true;
                            continue;
                        }

                        levelRow.IsBackup = isBackup;
                        sheet.Rows.Add(levelRow);
                    }

                    i++;
                }
            }

            return sheet;
        }

        private LevelRow GetLevelRow(string line, string sheetName, int row)
        {
            string[] arr = line.Split('\t');
            int index = StartColumnIndex - 1;
            string content = GetCellValue(arr, 0);
            string columnA = content;
            content = GetCellValue(arr, index);
            string pinName = content;
            index++;
            content = GetCellValue(arr, index);
            string seq = content;
            index++;
            content = GetCellValue(arr, index);
            string parameter = content;
            index++;
            content = GetCellValue(arr, index);
            string value = content;
            index++;
            content = GetCellValue(arr, index);
            string comment = content;
            var levelRow = new LevelRow(pinName, parameter, value, comment)
            {
                RowNum = row,
                SheetName = sheetName,
                ColumnA = columnA,
                Seq = seq
            };
            return levelRow;
        }

        private LevelRow GetLevelRow(ExcelWorksheet excelWorksheet, int row)
        {
            string pinName = "";
            string parameter = "";
            string value = "";
            string comment = "";
            for (int i = StartColumnIndex; i <= excelWorksheet.Dimension.End.Column; i++)
            {
                string head = _headList[i - StartColumnIndex];
                string content = GetCellText(excelWorksheet, row, i);
                switch (CellDiff.FormatStringForCompare(head))
                {
                    case "PIN/GROUP":
                        pinName = content;
                        break;
                    case "SEQ.":
                        break;
                    case "PARAMETER":
                        parameter = content;
                        break;
                    case "VALUE":
                        value = content.Replace("=", "");
                        break;
                    case "COMMENT":
                        comment = content;
                        break;
                }
            }
            return new LevelRow(pinName, parameter, value, comment, row);
        }
    }
}
