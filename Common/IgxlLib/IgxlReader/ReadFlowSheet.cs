using System;
using System.Collections.Generic;
using System.IO;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class ReadFlowSheet : IgxlSheetReader<SubFlowSheet>
    {
        private const int StartRowIndex = 3;

        public override EnumSheetType SupportedType => EnumSheetType.DTFlowtableSheet;

        public override SubFlowSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var subFlowSheet = new SubFlowSheet(excelWorksheet);
            if (excelWorksheet.Dimension == null)
            {
                return subFlowSheet;
            }

            int maxRowCount = excelWorksheet.Dimension.End.Row;

            bool isBackup = false;
            for (int i = StartRowIndex + 2; i <= maxRowCount; i++)
            {
                FlowRow flowRow = GetFlowRow(excelWorksheet, i);
                if (string.IsNullOrEmpty(flowRow.Opcode))
                {
                    isBackup = true;
                }
                flowRow.IsBackup = isBackup;
                if (!string.IsNullOrEmpty(flowRow.Opcode))
                {
                    subFlowSheet.AddRow(flowRow);
                }
            }
            return subFlowSheet;
        }

        public override SubFlowSheet ReadSheet(Stream stream, string sheetName)
        {
            var subFlowSheet = new SubFlowSheet(sheetName);
            bool isBackup = false;
            int i = 0;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine()!;
                    if (i > StartRowIndex)
                    {
                        FlowRow flowRow = GetFlowRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(flowRow.Opcode))
                        {
                            isBackup = true;
                        }

                        flowRow.IsBackup = isBackup;
                        subFlowSheet.AddRow(flowRow);
                    }
                    i++;
                }
            }
            return subFlowSheet;
        }

        private FlowRow GetFlowRow(ExcelWorksheet excelWorksheet, int row)
        {
            return InternalGetFlowRow(i => GetCellText(excelWorksheet, row, i + 1), excelWorksheet.Name, row);
        }

        private FlowRow GetFlowRow(string line, string sheetName, int row)
        {
            string[] arr = line.Split('\t');
            return InternalGetFlowRow(i => GetCellValue(arr, i), sheetName, row);
        }

        private static FlowRow InternalGetFlowRow(Func<int, string> getCellContent, string sheetName, int row)
        {
            var flowRow = new FlowRow { SheetName = sheetName, RowNum = row };
            int index = 0;

            // Helper to get content and auto-increment index
            string Next() => getCellContent(index++);

            flowRow.ColumnA = Next();
            flowRow.Label = Next();
            flowRow.Enable = Next();
            flowRow.Job = Next();
            flowRow.Part = Next();
            flowRow.Env = Next();
            flowRow.Opcode = Next();
            flowRow.Parameter = Next();
            flowRow.TName = Next();
            flowRow.TNum = Next();
            flowRow.LoLim = Next();
            flowRow.HiLim = Next();
            flowRow.Scale = Next();
            flowRow.Units = Next();
            flowRow.Format = Next();
            flowRow.BinPass = Next();
            flowRow.BinFail = Next();
            flowRow.SortPass = Next();
            flowRow.SortFail = Next();
            flowRow.Result = Next();
            flowRow.PassAction = Next();
            flowRow.FailAction = Next();
            flowRow.State = Next();
            flowRow.GroupSpecifier = Next();
            flowRow.GroupSense = Next();
            flowRow.GroupCondition = Next();
            flowRow.GroupName = Next();
            flowRow.DeviceSense = Next();
            flowRow.DeviceCondition = Next();
            flowRow.DeviceName = Next();
            flowRow.DebugAssume = Next();
            flowRow.DebugSites = Next();
            flowRow.CtProfileDataElapsedTime = Next();
            flowRow.CtProfileDataBackgroundType = Next();
            flowRow.CtProfileDataSerialize = Next();
            flowRow.CtProfileDataResourceLock = Next();
            flowRow.CtProfileDataFlowStepLocked = Next();
            flowRow.Comment = Next();

            var comment1List = new List<string>
            {
                Next(),
                Next()
            };
            flowRow.Comment1 = string.Join("\t", comment1List);

            return flowRow;
        }
    }
}
