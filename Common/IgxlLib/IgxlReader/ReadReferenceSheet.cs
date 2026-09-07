using System.Collections.Generic;
using System.IO;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class ReadReferenceSheet : IgxlSheetReader<ReferenceSheet>
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 1;

        public override EnumSheetType SupportedType => EnumSheetType.DTReferencesSheet;

        private List<string> _headList = [];

        public override ReferenceSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            _headList = [];
            var referenceSheet = new ReferenceSheet(excelWorksheet);
            int maxRowCount = excelWorksheet.Dimension.End.Row;
            int maxColumnCount = excelWorksheet.Dimension.End.Column;

            for (int i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                string value = GetMergeCellValue(excelWorksheet, StartRowIndex, i);
                string head = value.Trim();
                _headList.Add(head);
            }

            var referenceRows = new List<ReferenceRow>();
            bool isBackup = false;
            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                ReferenceRow referenceRow = GetReferenceRow(excelWorksheet, i);
                if (string.IsNullOrEmpty(referenceRow.FilePath))
                {
                    isBackup = true;
                    continue;
                }
                if (isBackup)
                {
                    referenceRow.IsBackup = true;
                }

                referenceRows.Add(referenceRow);
            }
            referenceSheet.Rows = referenceRows;

            return referenceSheet;
        }

        private ReferenceRow GetReferenceRow(ExcelWorksheet excelWorksheet, int row)
        {
            return GetRowData<ReferenceRow>(excelWorksheet, row, StartColumnIndex, _headList, (obj, header, content) =>
            {
                switch (header)
                {
                    case "COLUMNA":
                        obj.ColumnA = content;
                        break;
                    case "FILE PATH":
                        obj.FilePath = content;
                        break;
                    case "COMMENT":
                        obj.Comment = content;
                        break;
                }
            });
        }

        public override ReferenceSheet ReadSheet(Stream stream, string sheetName)
        {
            var referenceSheet = new ReferenceSheet(sheetName);
            bool isBackup = false;
            int i = 1;
            var referenceRows = new List<ReferenceRow>();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine()!;
                    if (i > StartRowIndex)
                    {
                        ReferenceRow patSetRow = GetReferenceRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(patSetRow.FilePath))
                        {
                            isBackup = true;
                            continue;
                        }

                        patSetRow.IsBackup = isBackup;
                        referenceRows.Add(patSetRow);
                    }

                    i++;
                }
            }

            referenceSheet.Rows = referenceRows;
            return referenceSheet;
        }

        private ReferenceRow GetReferenceRow(string line, string sheetName, int row)
        {
            return CreateRow<ReferenceRow>(line, sheetName, row, (rowObj, arr) =>
            {
                rowObj.FilePath = GetCellValue(arr, StartColumnIndex);
                rowObj.Comment = GetCellValue(arr, StartColumnIndex + 1);
            });
        }
    }
}
