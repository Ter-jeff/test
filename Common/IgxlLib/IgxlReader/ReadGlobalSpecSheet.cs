using System.IO;

using CommonLib.Utility;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class ReadGlobalSpecSheet : IgxlSheetReader<GlobalSpecSheet>
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        public override EnumSheetType SupportedType => EnumSheetType.DTGlobalSpecSheet;

        public override GlobalSpecSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var globalSpecSheet = new GlobalSpecSheet(excelWorksheet, false);
            int maxRowCount = excelWorksheet.Dimension.End.Row;

            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                GlobalSpec globalSpec = GetGlobalSpecRow(excelWorksheet, i);
                if (string.IsNullOrEmpty(globalSpec.Symbol))
                {
                    break;
                }

                globalSpecSheet.AddRow(globalSpec);
            }
            return globalSpecSheet;
        }

        public override GlobalSpecSheet ReadSheet(Stream stream, string sheetName)
        {
            var globalSpecSheet = new GlobalSpecSheet(sheetName, false);
            bool isBackup = false;
            int i = 1;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine()!;
                    if (i > StartRowIndex)
                    {
                        GlobalSpec globalSpec = GetGlobalSpecRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(globalSpec.Symbol))
                        {
                            isBackup = true;
                            continue;
                        }

                        globalSpec.IsBackup = isBackup;
                        globalSpecSheet.AddRow(globalSpec);
                    }

                    i++;
                }
            }

            return globalSpecSheet;
        }

        private GlobalSpec GetGlobalSpecRow(string line, string sheetName, int row)
        {
            string[] arr = line.Split('\t');
            var globalSpec = new GlobalSpec("")
            {
                RowNum = row,
                SheetName = sheetName
            };
            int index = StartColumnIndex - 1;
            string content = GetCellValue(arr, 0);
            globalSpec.ColumnA = content;
            content = GetCellValue(arr, index);
            globalSpec.Symbol = content;
            index++;
            content = GetCellValue(arr, index);
            globalSpec.Job = content;
            index++;
            content = GetCellValue(arr, index);
            globalSpec.Value = content;
            index++;
            content = GetCellValue(arr, index);
            globalSpec.Comment = content;
            return globalSpec;
        }

        private GlobalSpec GetGlobalSpecRow(ExcelWorksheet excelWorksheet, int row)
        {
            string symbol = "";
            string job = "";
            string value = "";
            string comment = "";
            for (int i = StartColumnIndex; i <= excelWorksheet.Dimension.End.Column; i++)
            {
                string head = GetMergeCellValue(excelWorksheet, StartRowIndex, i);
                string content = GetCellText(excelWorksheet, row, i);
                switch (CellDiff.FormatStringForCompare(head))
                {
                    case "SYMBOL":
                        symbol = content;
                        break;
                    case "JOB":
                        job = content;
                        break;
                    case "VALUE":
                        value = content;
                        break;
                    case "COMMENT":
                        comment = content;
                        break;
                }
            }
            return new GlobalSpec(symbol, job, value, comment);
        }
    }
}
