using System.Collections.Generic;
using System.IO;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class ReadPatSubroutineSheet : IgxlSheetReader<PatSetSubSheet>
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 1;

        public override EnumSheetType SupportedType => EnumSheetType.DTPatternSubroutineSheet;

        private List<string> _headList = [];

        public override PatSetSubSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            _headList = [];
            var patSetSubSheet = new PatSetSubSheet(excelWorksheet);
            int maxRowCount = excelWorksheet.Dimension.End.Row;
            int maxColumnCount = excelWorksheet.Dimension.End.Column;

            for (int i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                string value = GetMergeCellValue(excelWorksheet, StartRowIndex, i);
                string head = value.Trim();
                _headList.Add(head);
            }

            var patSetSubRows = new List<PatSetSubRow>();
            bool isBackup = false;
            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                PatSetSubRow patSetSubRow = GetPatSetSubRow(excelWorksheet, i);
                if (string.IsNullOrEmpty(patSetSubRow.PatternFileName))
                {
                    isBackup = true;
                    continue;
                }
                patSetSubRow.IsBackup = isBackup;

                patSetSubRows.Add(patSetSubRow);
            }
            patSetSubSheet.Rows = patSetSubRows;

            return patSetSubSheet;
        }

        public override PatSetSubSheet ReadSheet(Stream stream, string sheetName)
        {
            var patSetSubSheet = new PatSetSubSheet(sheetName);
            bool isBackup = false;
            int i = 1;
            var patSetSubRows = new List<PatSetSubRow>();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine()!;
                    if (i > StartRowIndex)
                    {
                        PatSetSubRow patSetRow = GetPatSetSubRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(patSetRow.PatternFileName))
                        {
                            isBackup = true;
                            continue;
                        }

                        patSetRow.IsBackup = isBackup;
                        patSetSubRows.Add(patSetRow);
                    }

                    i++;
                }
            }

            patSetSubSheet.Rows = patSetSubRows;
            return patSetSubSheet;
        }

        private PatSetSubRow GetPatSetSubRow(string line, string sheetName, int row)
        {
            return CreateRow<PatSetSubRow>(line, sheetName, row, (rowObj, arr) =>
            {
                rowObj.PatternFileName = GetCellValue(arr, StartColumnIndex);
                rowObj.Comment = GetCellValue(arr, StartColumnIndex + 1);
            });
        }

        private PatSetSubRow GetPatSetSubRow(ExcelWorksheet excelWorksheet, int row)
        {
            return GetRowData<PatSetSubRow>(excelWorksheet, row, StartColumnIndex, _headList, (obj, header, content) =>
            {
                switch (header)
                {
                    case "COLUMNA":
                        obj.ColumnA = content;
                        break;
                    case "PATTERN FILENAME":
                        obj.PatternFileName = content;
                        break;
                    case "COMMENT":
                        obj.Comment = content;
                        break;
                }
            });
        }
    }
}
