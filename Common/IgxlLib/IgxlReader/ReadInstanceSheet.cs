using System.Collections.Generic;
using System.IO;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class ReadInstanceSheet : IgxlSheetReader<InstanceSheet>
    {
        private const int StartRowIndex = 4;
        private const int StartColumnIndex = 2;

        public override EnumSheetType SupportedType => EnumSheetType.DTTestInstancesSheet;

        public override InstanceSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var instanceSheet = new InstanceSheet(excelWorksheet);
            int maxRowCount = excelWorksheet.Dimension.End.Row;

            bool backup = false;
            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                InstanceRow instanceRow = GetInstanceRow(excelWorksheet, i);
                if (string.IsNullOrEmpty(instanceRow.TestName))
                {
                    backup = true;
                }
                instanceRow.IsBackup = backup;
                if (!string.IsNullOrEmpty(instanceRow.TestName))
                {
                    instanceSheet.AddRow(instanceRow);
                }
            }

            return instanceSheet;
        }

        public override InstanceSheet ReadSheet(Stream stream, string sheetName)
        {
            var instanceSheet = new InstanceSheet(sheetName);
            bool isBackup = false;
            int i = 1;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine()!;
                    if (i > StartRowIndex)
                    {
                        InstanceRow instanceRow = GetInstanceRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(instanceRow.TestName))
                        {
                            isBackup = true;
                            continue;
                        }

                        instanceRow.IsBackup = isBackup;
                        instanceSheet.AddRow(instanceRow);
                    }

                    i++;
                }
            }

            return instanceSheet;
        }

        public List<InstanceRow> GetSheetInstanceRows(ExcelWorksheet excelWorksheet)
        {
            var instanceRows = new List<InstanceRow>();
            int maxRowCount = excelWorksheet.Dimension.End.Row;

            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                InstanceRow instanceRow = GetInstanceRow(excelWorksheet, i);
                if (string.IsNullOrEmpty(instanceRow.TestName))
                {
                    break;
                }

                instanceRows.Add(instanceRow);
            }

            return instanceRows;
        }

        private InstanceRow GetInstanceRow(string line, string sheetName, int row)
        {
            string[] arr = line.Split('\t');
            var instanceRow = new InstanceRow { RowNum = row, SheetName = sheetName };
            int index = StartColumnIndex - 1;
            string content = GetCellValue(arr, 0);
            instanceRow.ColumnA = content;
            content = GetCellValue(arr, index);
            instanceRow.TestName = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.VbtType = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.VbtName = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.CalledAs = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.DcCategory = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.DcSelector = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.AcCategory = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.AcSelector = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.TimeSets = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.EdgeSets = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.PinLevels = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.MixedSignalTiming = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.Overlay = content;
            index++;
            content = GetCellValue(arr, index);
            instanceRow.ArgList = content;
            for (int i = 1; i <= 130; i++)
            {
                content = GetCellValue(arr, index + i);
                instanceRow.Args.Add(content);
            }

            index += 130;
            content = GetCellValue(arr, index);
            instanceRow.Comment = content;
            return instanceRow;
        }

        private InstanceRow GetInstanceRow(ExcelWorksheet excelWorksheet, int row)
        {
            var instanceRow = new InstanceRow { SheetName = excelWorksheet.Name, RowNum = row };
            int index = 2;
            string content = GetCellText(excelWorksheet, row, 1);
            instanceRow.ColumnA = content;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.TestName = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.VbtType = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.VbtName = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.CalledAs = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.DcCategory = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.DcSelector = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.AcCategory = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.AcSelector = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.TimeSets = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.EdgeSets = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.PinLevels = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.MixedSignalTiming = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.Overlay = content;
            index++;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.ArgList = content;
            for (int i = 1; i <= 130; i++)
            {
                content = GetCellText(excelWorksheet, row, index + i);
                instanceRow.Args.Add(content);
            }
            index += 130;
            content = GetCellText(excelWorksheet, row, index);
            instanceRow.Comment = content;
            return instanceRow;
        }
    }
}
