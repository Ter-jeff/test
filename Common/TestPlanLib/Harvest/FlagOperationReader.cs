using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using CommonReaderLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using OpCode = IgxlLib.IgxlConst.OpCode;

namespace TestPlanLib.Harvest
{
    public class FlagOperationReader : MySheetReader<FlagOperationSheet>
    {
        private const string ConHeaderNewFlag = "New Flag";
        private const string ConHeaderOperator = "Operator";
        private const string ConHeaderExistFlags = "Existing Flags";

        private int _indexNewFlag = -1;
        private int _indexOperator = -1;
        private int _indexExistFlags = -1;

        public override FlagOperationSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            FlagOperationSheet flagOperationSheet = ReadSheet(sheetName);

            return flagOperationSheet;
        }

        private FlagOperationSheet ReadSheet(string sheetName)
        {
            var flagOperationSheet = new FlagOperationSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new FlagOperationRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexNewFlag != -1)
                {
                    row.NewFlag = ExcelWorksheet.GetCellValue(i, _indexNewFlag).Trim();
                }

                if (row.NewFlag?.Length == 0)
                {
                    continue;
                }

                if (_indexOperator != -1)
                {
                    row.Operator = ExcelWorksheet.GetCellValue(i, _indexOperator).Trim();
                }

                if (_indexExistFlags != -1)
                {
                    string existFlags = ExcelWorksheet.GetCellValue(i, _indexExistFlags).Trim();
                    row.Value = existFlags;
                    row.ExistFlags = [.. existFlags.Replace(" ", "").Split(',')];
                }
                flagOperationSheet.Rows.Add(row);
            }

            flagOperationSheet.IndexNewFlag = _indexNewFlag;
            flagOperationSheet.IndexOperator = _indexOperator;
            flagOperationSheet.IndexExistFlags = _indexExistFlags;

            return flagOperationSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderNewFlag))
                {
                    _indexNewFlag = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderOperator))
                {
                    _indexOperator = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderExistFlags))
                {
                    _indexExistFlags = i;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderNewFlag))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public class FlagOperationSheet : MySheet
    {
        public List<FlagOperationRow> Rows { set; get; }

        public int IndexNewFlag = -1;
        public int IndexOperator = -1;
        public int IndexExistFlags = -1;
        public string SubFlowName = "";

        #region Constructor
        public FlagOperationSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
            SubFlowName = "Flow_" + SheetName;
        }
        #endregion

        public SubFlowSheet CreateSubFlow()
        {
            var subflow = new SubFlowSheet(SubFlowName, SheetName);
            subflow.AddPrintStartRow(SubFlowName);

            foreach (FlagOperationRow row in Rows)
            {
                switch (row.Operator.ToUpper())
                {
                    case "FLAG-TRUE":
                        subflow.AddRow(new FlowRow { Opcode = OpCode.FlagTrue, Parameter = row.NewFlag });
                        break;
                    case "FLAG-FALSE":
                        subflow.AddRow(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = row.NewFlag });
                        break;
                    case "OR":
                        subflow.AddRows(CreateOpRow("||", row));
                        break;
                    case "AND":
                        subflow.AddRows(CreateOpRow("&&", row));
                        break;
                    case "EQUAL":
                        subflow.AddRows(CreateEqual(row));
                        break;
                }
            }
            subflow.AddPrintEndRow(SubFlowName);
            subflow.AddReturnRow();
            return subflow;
        }

        public List<string> GetAllFlags()
        {
            var flags = new List<string>();
            flags.AddRange(Rows.Select(x => x.NewFlag));
            foreach (IEnumerable<string> existFlags in Rows.Select(x => x.ExistFlags.SelectMany(existFlag => existFlag.Replace("&&", ",").Replace("||", ",").Replace("(", "").Replace(")", "").Replace("!", "").Split(','))))
            {
                foreach (string existFlag in existFlags)
                {
                    if (string.IsNullOrEmpty(existFlag))
                    {
                        continue;
                    }

                    if (flags.Exists(x => x.EqualsIgnoreCase(existFlag)))
                    {
                        continue;
                    }

                    flags.Add(existFlag);
                }
            }
            return flags;
        }

        private static List<FlowRow> CreateOpRow(string operation, FlagOperationRow flagOperationRow)
        {
            var flowRows = new List<FlowRow>();
            List<string> flags = flagOperationRow.ExistFlags;
            flowRows.Add(new FlowRow { Opcode = OpCode.If, Parameter = string.Join(operation, flags) });
            flowRows.Add(new FlowRow { Opcode = OpCode.FlagTrue, Parameter = flagOperationRow.NewFlag });
            flowRows.Add(new FlowRow { Opcode = OpCode.Else });
            flowRows.Add(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = flagOperationRow.NewFlag });
            flowRows.Add(new FlowRow { Opcode = OpCode.EndIf });
            return flowRows;
        }

        private static List<FlowRow> CreateEqual(FlagOperationRow flagOperationRow)
        {
            var flowRows = new List<FlowRow>
            {
                new() { Opcode = OpCode.FlagClear, Parameter = flagOperationRow.NewFlag },
                new() { Opcode = OpCode.FlagTrue, Parameter = flagOperationRow.NewFlag, DeviceCondition = OpCode.FlagTrue, DeviceName = flagOperationRow.Value },
                new() { Opcode = OpCode.FlagFalse, Parameter = flagOperationRow.NewFlag, DeviceCondition = OpCode.FlagFalse, DeviceName = flagOperationRow.Value }
            };
            return flowRows;
        }
    }

    public class FlagOperationRow : MyRow
    {
        public string NewFlag = "";
        public List<string> ExistFlags = [];
        public string Operator = "";
        public string Value = "";
        #region Constructor
        public FlagOperationRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion

    }
}
