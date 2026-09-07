using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class ReadPortMapSheet : IgxlSheetReader<PortMapSheet>
    {
        private const int StartRowIndex = 4;
        private const int StartColumnIndex = 2;

        public override EnumSheetType SupportedType => EnumSheetType.DTPortMapSheet;

        public override PortMapSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var portMapSheet = new PortMapSheet(excelWorksheet);
            int stopRow = excelWorksheet.Dimension.End.Row;
            var portRows = new List<PortRow>();
            bool flag = false;
            for (int i = StartRowIndex + 1; i <= stopRow; i++)
            {
                PortRow portRow = GetPortMapRow(excelWorksheet, i);
                portRow.RowNum = i;
                if (string.IsNullOrEmpty(portRow.PortName))
                {
                    flag = true;
                }

                portRow.IsBackup = flag;
                portRows.Add(portRow);
            }

            var groups = portRows.ChunkBy(x => x.PortName).ToList();
            foreach (IGrouping<string, PortRow> group in groups)
            {
                var portSet = new PortSet(group.Key)
                {
                    PortRows = [.. group]
                };
                portMapSheet.Rows.Add(portSet);
            }

            return portMapSheet;
        }

        public override PortMapSheet ReadSheet(Stream stream, string sheetName)
        {
            var portMapSheet = new PortMapSheet(sheetName);
            bool isBackup = false;
            int i = 1;
            var portRows = new List<PortRow>();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string? line = sr.ReadLine();
                    if (i > StartRowIndex)
                    {
                        if (line != null)
                        {
                            PortRow portRow = GetPortRow(line, sheetName, i);
                            if (string.IsNullOrEmpty(portRow.PortName))
                            {
                                isBackup = true;
                            }

                            portRow.IsBackup = isBackup;
                            portRows.Add(portRow);
                        }
                    }

                    i++;
                }
            }

            var groups = portRows.ChunkBy(x => x.PortName).ToList();
            foreach (IGrouping<string, PortRow> group in groups)
            {
                var portSet = new PortSet(group.Key)
                {
                    PortRows = [.. group]
                };
                portMapSheet.Rows.Add(portSet);
            }

            return portMapSheet;
        }

        private PortRow GetPortRow(string line, string sheetName, int row)
        {
            string[] arr = line.Split('\t');
            var portRow = new PortRow
            {
                RowNum = row,
                SheetName = sheetName
            };
            int index = StartColumnIndex - 1;
            string content = GetCellValue(arr, 0);
            portRow.ColumnA = content;
            content = GetCellValue(arr, index);
            portRow.PortName = content;
            index++;
            content = GetCellValue(arr, index);
            portRow.ProtocolFamily = content;
            index++;
            content = GetCellValue(arr, index);
            portRow.ProtocolType = content;
            index++;
            content = GetCellValue(arr, index);
            portRow.ProtocolSettings = content;
            index++;
            for (int i = 0; i < 10; i++)
            {
                content = GetCellValue(arr, index);
                portRow.ProtocolSettingValues.Add(content);
                index++;
            }

            content = GetCellValue(arr, index);
            portRow.FunctionName = content;
            index++;
            content = GetCellValue(arr, index);
            portRow.FunctionPin = content;
            index++;
            content = GetCellValue(arr, index);
            portRow.FunctionProperties = content;
            index++;
            for (int i = 0; i < 10; i++)
            {
                content = GetCellValue(arr, index);
                portRow.FunctionPropertyValues.Add(content);
                index++;
            }

            index++;
            content = GetCellValue(arr, index);
            portRow.Comment = content;
            return portRow;
        }

        private static PortRow GetPortMapRow(ExcelWorksheet excelWorksheet, int row)
        {
            var portRow = new PortRow
            {
                SheetName = excelWorksheet.Name,
                RowNum = row
            };
            string content = excelWorksheet.GetCellValue(row, 1);
            portRow.ColumnA = content;
            int index = StartColumnIndex;
            content = excelWorksheet.GetCellValue(row, index);
            portRow.PortName = content;
            index++;
            content = excelWorksheet.GetCellValue(row, index);
            portRow.ProtocolFamily = content;
            index++;
            content = excelWorksheet.GetCellValue(row, index);
            portRow.ProtocolType = content;
            index++;
            content = excelWorksheet.GetCellValue(row, index);
            portRow.ProtocolSettings = content;
            index++;
            for (int i = 0; i < 10; i++)
            {
                portRow.ProtocolSettingValues.Add(excelWorksheet.GetCellValue(row, index + i));
            }

            index += 10;
            content = excelWorksheet.GetCellValue(row, index);
            portRow.FunctionName = content;
            index++;
            content = excelWorksheet.GetCellValue(row, index);
            portRow.FunctionPin = content;
            index++;
            content = excelWorksheet.GetCellValue(row, index);
            portRow.FunctionProperties = content;
            index++;
            for (int i = 0; i < 10; i++)
            {
                portRow.FunctionPropertyValues.Add(excelWorksheet.GetCellValue(row, index + i));
            }

            index += 10;
            content = excelWorksheet.GetCellValue(row, index);
            portRow.Comment = content;
            return portRow;
        }
    }
}
