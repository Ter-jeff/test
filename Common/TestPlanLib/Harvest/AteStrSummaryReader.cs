using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using CommonReaderLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using static TestPlanLib.Harvest.AteStrSummarySheet;

using OpCode = IgxlLib.IgxlConst.OpCode;

namespace TestPlanLib.Harvest
{
    public class AteStrSummaryReader : MySheetReader<AteStrSummarySheet>
    {
        private const string ConHeaderTestName = "PTR Test Name";
        private const string ConHeaderOperator = "Operator";
        private const string ConHeaderValue = "Value";

        private int _indexTestName = -1;
        private int _indexOperator = -1;
        private int _indexValue = -1;

        public override AteStrSummarySheet? ReadSheet(ExcelWorksheet excelWorksheet)
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

            AteStrSummarySheet ateStrSummarySheet = ReadSheet(sheetName);

            ateStrSummarySheet.FlagsFromValue = ateStrSummarySheet.GetAllFlagFromValue();

            return ateStrSummarySheet;
        }

        private AteStrSummarySheet ReadSheet(string sheetName)
        {
            var ateStrSummarySheet = new AteStrSummarySheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new AteStrSummaryRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexTestName != -1)
                {
                    row.TestName = ExcelWorksheet.GetCellValue(i, _indexTestName).Trim();
                }

                if (row.TestName?.Length == 0)
                {
                    continue;
                }

                if (_indexOperator != -1)
                {
                    row.Operator = ExcelWorksheet.GetCellValue(i, _indexOperator).Trim();
                }

                if (_indexValue != -1)
                {
                    row.Value = ExcelWorksheet.GetCellValue(i, _indexValue).Trim();
                }
                ateStrSummarySheet.Rows.Add(row);
            }

            ateStrSummarySheet.IndexNewFlag = _indexTestName;
            ateStrSummarySheet.IndexOperator = _indexOperator;
            ateStrSummarySheet.IndexExistFlags = _indexValue;

            return ateStrSummarySheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderTestName))
                {
                    _indexTestName = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderOperator))
                {
                    _indexOperator = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderValue))
                {
                    _indexValue = i;
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
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderTestName))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public class AteStrSummarySheet : MySheet
    {
        public enum EnumAteType
        {
            TD, BIST, SA, None
        }

        public List<AteStrSummaryRow> Rows { set; get; }
        public List<string> FlagsFromValue { get; set; } = [];

        public int IndexNewFlag = -1;
        public int IndexOperator = -1;
        public int IndexExistFlags = -1;
        public string SubFlowName = "Flow_ATE_STR_Summary";

        public AteStrSummarySheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }

        public List<string> GetAllFlagFromValue()
        {
            List<string> flags = [];

            foreach (AteStrSummaryRow row in Rows)
            {
                switch (row.Operator)
                {
                    case "OR":
                        flags.AddRange([.. row.Value.Replace(" ", "").Split(',')]);
                        break;
                    case "AND":
                        flags.AddRange([.. row.Value.Replace(" ", "").Split(',')]);
                        break;
                    case "EQUAL":
                        flags.Add(row.Value);
                        break;
                    case "IF":
                        flags.Add(row.Value.Split(',').First().Split('=').First().Replace("(", ""));
                        flags.Add(row.Value.Split(',')[1].Replace(" ", ""));
                        break;
                }
            }

            return [.. flags.Distinct()];
        }

        public List<SubFlowSheet> CreateFlows()
        {
            IEnumerable<IGrouping<EnumAteType, AteStrSummaryRow>> orderByType = Rows.GroupBy(x => x.AteType);
            var subFlows = new List<SubFlowSheet>();
            foreach (IGrouping<EnumAteType, AteStrSummaryRow> rows in orderByType)
            {
                EnumAteType type = rows.Key;
                subFlows.Add(CreateSubFlow($"{SubFlowName}_{type}", [.. rows]));
            }
            subFlows.Add(CreateAteMainFlow([.. subFlows.Select(x => x.Name)]));
            return subFlows;
        }

        private SubFlowSheet CreateAteMainFlow(List<string> subFlowList)
        {
            string sheetName = SubFlowName;
            var subflow = new SubFlowSheet(sheetName, SheetName);
            subflow.AddPrintStartRow(SheetName);
            foreach (string subflowName in subFlowList)
            {
                subflow.AddRow(new FlowRow { Opcode = OpCode.Call, Parameter = subflowName });
            }

            subflow.AddRow(new FlowRow { Opcode = OpCode.Test, Parameter = "ATE_STR_Summary", DeviceName = "F_PrintHarvReport", DeviceCondition = OpCode.FlagTrue });
            subflow.AddPrintEndRow(sheetName);
            subflow.AddReturnRow();
            return subflow;
        }

        private static SubFlowSheet CreateSubFlow(string sheetName, List<AteStrSummaryRow> ateStrSummaryRows)
        {
            var subflow = new SubFlowSheet(sheetName);
            subflow.AddPrintStartRow(sheetName);
            //subflow.AddRow(new FlowRow { Opcode = OpCode.FlagFalseAll, Parameter = string.Join(",", Rows.Select(x => x.TestName)) });
            foreach (AteStrSummaryRow row in ateStrSummaryRows)
            {
                switch (row.Operator.ToUpper())
                {
                    case "FLAG-TRUE":
                        subflow.AddRow(new FlowRow { Opcode = OpCode.FlagTrue, Parameter = row.TestName });
                        break;
                    case "FLAG-FALSE":
                        subflow.AddRow(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = row.TestName });
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
                    case "IF":
                        subflow.AddRows(CreateOpIf(row));
                        break;
                }
            }
            subflow.AddPrintEndRow(sheetName);
            subflow.AddReturnRow();
            return subflow;
        }

        private static List<FlowRow> CreateEqual(AteStrSummaryRow ateStrSummaryRow)
        {
            var flowRows = new List<FlowRow>
            {
                new() { Opcode = OpCode.FlagClear, Parameter = ateStrSummaryRow.TestName },
                new() { Opcode = OpCode.FlagTrue, Parameter = ateStrSummaryRow.TestName, DeviceCondition = OpCode.FlagTrue, DeviceName = ateStrSummaryRow.Value },
                new() { Opcode = OpCode.FlagFalse, Parameter = ateStrSummaryRow.TestName, DeviceCondition = OpCode.FlagFalse, DeviceName = ateStrSummaryRow.Value }
            };
            return flowRows;
        }

        private static List<FlowRow> CreateOpIf(AteStrSummaryRow ateStrSummaryRow)
        {
            var flowRows = new List<FlowRow>();
            string conditionFlag = ateStrSummaryRow.Value.Split(',').First().Split('=').First().Replace("(", "");
            string condition = ateStrSummaryRow.Value.Split(',').First().Split('=').Last().Replace(" ", "") == "1" ? OpCode.FlagTrue : OpCode.FlagFalse;
            string referFlag = ateStrSummaryRow.Value.Split(',')[1].Replace(" ", "");
            string referValue = GetOpCodeByValue(ateStrSummaryRow.Value.Split(',').Last().Replace(" ", ""));
            flowRows.Add(new FlowRow { Opcode = referValue, Parameter = ateStrSummaryRow.TestName });
            flowRows.Add(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = ateStrSummaryRow.TestName, DeviceCondition = condition, DeviceName = conditionFlag });
            flowRows.Add(new FlowRow { Opcode = OpCode.If, Parameter = referFlag, DeviceCondition = condition, DeviceName = conditionFlag });
            flowRows.Add(new FlowRow { Opcode = OpCode.FlagTrue, Parameter = ateStrSummaryRow.TestName, DeviceCondition = condition, DeviceName = conditionFlag });
            flowRows.Add(new FlowRow { Opcode = OpCode.EndIf, DeviceCondition = condition, DeviceName = conditionFlag });
            return flowRows;
        }

        private static List<FlowRow> CreateOpRow(string operation, AteStrSummaryRow ateStrSummaryRow)
        {
            var flowRows = new List<FlowRow>();
            List<string> flags = [.. ateStrSummaryRow.Value.Replace(" ", "").Split(',')];
            flowRows.Add(new FlowRow { Opcode = OpCode.If, Parameter = string.Join(operation, flags) });
            flowRows.Add(new FlowRow { Opcode = OpCode.FlagTrue, Parameter = ateStrSummaryRow.TestName });
            flowRows.Add(new FlowRow { Opcode = OpCode.Else });
            flowRows.Add(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = ateStrSummaryRow.TestName });
            flowRows.Add(new FlowRow { Opcode = OpCode.EndIf });
            return flowRows;
        }

        private static string GetOpCodeByValue(string value)
        {
            switch (value)
            {
                case "1":
                    return OpCode.FlagTrue;
                case "0":
                    return OpCode.FlagFalse;
                case "-1":
                    break;
            }
            return OpCode.FlagClear;
        }

        public List<string> GetAllFlags()
        {
            return Rows.ConvertAll(x => x.TestName);
        }
    }

    public class AteStrSummaryRow : MyRow
    {
        public string TestName = "";
        public string Value = "";
        public string Operator = "";
        private EnumAteType _ateType = EnumAteType.None;
        public EnumAteType AteType
        {
            get
            {
                if (_ateType == EnumAteType.None)
                {
                    if (TestName.Contains("TD", StringComparison.OrdinalIgnoreCase))
                    {
                        _ateType = EnumAteType.TD;
                    }
                    else if (TestName.Contains("SA", StringComparison.OrdinalIgnoreCase))
                    {
                        _ateType = EnumAteType.SA;
                    }
                    else if (TestName.Contains("BIST", StringComparison.OrdinalIgnoreCase))
                    {
                        _ateType = EnumAteType.BIST;
                    }
                    else
                    {
                        _ateType = EnumAteType.None;
                    }
                }
                return _ateType;
            }
        }
        #region Constructor
        public AteStrSummaryRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion

    }
}
