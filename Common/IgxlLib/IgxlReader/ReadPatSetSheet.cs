using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class ReadPatSetSheet : IgxlSheetReader<PatSetSheet>
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 1;

        public override EnumSheetType SupportedType => EnumSheetType.DTPatternSetSheet;

        private List<string> _headList = [];

        public override PatSetSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            _headList = [];
            var patSetSheet = new PatSetSheet(excelWorksheet);
            int maxRowCount = excelWorksheet.Dimension.End.Row;
            int maxColumnCount = excelWorksheet.Dimension.End.Column;

            for (int i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                string value = GetMergeCellValue(excelWorksheet, StartRowIndex, i);
                string head = value.Trim();
                _headList.Add(head);
            }

            var patSetRows = new List<PatSetRow>();
            bool isBackup = false;
            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                PatSetRow patSetRow = GetPatSetRow(excelWorksheet, i);
                if (string.IsNullOrEmpty(patSetRow.PatternSet))
                {
                    isBackup = true;
                    continue;
                }
                patSetRow.IsBackup = isBackup;
                patSetRows.Add(patSetRow);
            }

            var group = patSetRows.ChunkBy(x => x.PatternSet).ToList();
            foreach (IGrouping<string, PatSetRow> item in group)
            {
                var patSet = new PatSet
                {
                    PatSetName = item.Key,
                    IsBackup = item.First().IsBackup
                };
                patSet.PatSetRows.AddRange(item);
                patSetSheet.AddRow(patSet);
            }
            return patSetSheet;
        }

        public override PatSetSheet ReadSheet(Stream stream, string sheetName)
        {
            var patSetSheet = new PatSetSheet(sheetName);
            var rows = new List<PatSetRow>();
            bool isBackup = false;
            int i = 1;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine()!;
                    if (i > StartRowIndex)
                    {
                        PatSetRow patSetRow = GetPatSetRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(patSetRow.PatternSet))
                        {
                            isBackup = true;
                            continue;
                        }

                        patSetRow.IsBackup = isBackup;
                        rows.Add(patSetRow);
                    }

                    i++;
                }
            }

            string temp = "";
            foreach (PatSetRow row in rows)
            {
                if (string.IsNullOrEmpty(row.PatternSet))
                {
                    row.PatternSet = temp;
                }

                if (!string.IsNullOrEmpty(row.PatternSet))
                {
                    temp = row.PatternSet;
                }
            }

            IEnumerable<IGrouping<string, PatSetRow>> groups = rows.ChunkBy(x => x.PatternSet);
            foreach (IGrouping<string, PatSetRow> group in groups)
            {
                var patSet = new PatSet { PatSetName = group.Key };
                patSet.PatSetRows.AddRange([.. group]);
                patSetSheet.Rows.Add(patSet);
            }
            return patSetSheet;
        }

        private PatSetRow GetPatSetRow(string line, string sheetName, int row)
        {
            string[] arr = line.Split('\t');
            var patSetRow = new PatSetRow
            {
                RowNum = row,
                SheetName = sheetName
            };
            int index = StartColumnIndex;
            string content = GetCellValue(arr, 0);
            patSetRow.ColumnA = content;
            content = GetCellValue(arr, index);
            patSetRow.PatternSet = content;
            index++;
            content = GetCellValue(arr, index);
            patSetRow.TimeDomain = content;
            index++;
            content = GetCellValue(arr, index);
            patSetRow.Enable = content;
            index++;
            content = GetCellValue(arr, index);
            patSetRow.File = content;
            index++;
            content = GetCellValue(arr, index);
            patSetRow.Burst = content;
            index++;
            content = GetCellValue(arr, index);
            patSetRow.StartLabel = content;
            index++;
            content = GetCellValue(arr, index);
            patSetRow.StopLabel = content;
            index++;
            content = GetCellValue(arr, index);
            patSetRow.Comment = content;
            return patSetRow;
        }

        private PatSetRow GetPatSetRow(ExcelWorksheet excelWorksheet, int row)
        {
            var patSetRow = new PatSetRow
            {
                SheetName = excelWorksheet.Name,
                RowNum = row
            };
            for (int i = StartColumnIndex; i < StartColumnIndex + _headList.Count; i++)
            {
                string head = _headList[i - StartColumnIndex];
                string content = GetCellText(excelWorksheet, row, i);
                if (i == 1)
                {
                    head = "ColumnA";
                }
                switch (CellDiff.FormatStringForCompare(head))
                {
                    case "COLUMNA":
                        patSetRow.ColumnA = content;
                        break;
                    case "PATTERN SET":
                        patSetRow.PatternSet = content;
                        break;
                    case "TD_GROUP":
                        patSetRow.TdGroup = content;
                        break;
                    case "TIME DOMAIN":
                        patSetRow.TimeDomain = content;
                        break;
                    case "ENABLE":
                        patSetRow.Enable = content;
                        break;
                    case "FILE/GROUP NAME":
                        patSetRow.File = content;
                        break;
                    case "BURST":
                        patSetRow.Burst = content;
                        break;
                    case "START LABEL":
                        patSetRow.StartLabel = content;
                        break;
                    case "STOP LABEL":
                        patSetRow.StopLabel = content;
                        break;
                    case "COMMENT":
                        patSetRow.Comment = content;
                        break;
                }
            }
            return patSetRow;
        }
    }
}
