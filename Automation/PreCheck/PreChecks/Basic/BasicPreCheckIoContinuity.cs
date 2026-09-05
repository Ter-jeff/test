using System;
using System.Collections.Generic;

using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.IgxlBase;

using OfficeOpenXml;

using TestPlanLib.DataStruct;

namespace Automation.PreCheck.PreChecks.Basic
{
    public class BasicPreCheckIoContinuity : BasicPreCheckBase
    {
        private const string ConFs = "FS";
        private const string ConDd = "DD";
        private readonly List<string> _grpList = new List<string>();

        public BasicPreCheckIoContinuity(ExcelWorkbook excelWorkbook, string sheetName) : base(excelWorkbook, sheetName)
        {
            FirstHeader = IoContiRow.ConHeaderBumpName;
        }
        protected internal override bool CheckHeaders()
        {
            var headers = new List<string>
            {
                IoContiRow.ConHeaderBumpName,
                IoContiRow.ConHeaderIoVoltage,
                IoContiRow.ConHeaderFsDd
            };

            if (!CheckColumnHeadByList(StartRow, headers))
            {
                return false;
            }

            return true;
        }

        protected internal override bool CheckFormat()
        {
            return true;
        }

        protected internal override bool CheckBusiness()
        {
            bool result = true;
            int pinColumn = 0;
            int fsDdColumn = 0;
            int ioVolColumn = 0;
            if (TestProgram.IgxlWorkBk.PinMapPair.Value == null)
            {
                ErrorReportManager.AddError(BasicErrorType.E_MissingPin_06, ExcelWorksheet.Name, 1, 1, "Missing PinMap in TestPlan");
                result = false;
            }

            for (int i = StartColumn; i <= ExcelWorksheet.Dimension.End.Column; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i);
                if (CellDiff.IsLiked(header, IoContiRow.ConHeaderBumpName))
                {
                    pinColumn = i;
                }
                if (CellDiff.IsLiked(header, IoContiRow.ConHeaderFsDd))
                {
                    fsDdColumn = i;
                }
                if (CellDiff.IsLiked(header, IoContiRow.ConHeaderIoVoltage))
                {
                    ioVolColumn = i;
                }
            }

            for (int j = StartRow + 1; j <= ExcelWorksheet.Dimension.End.Row; j++)
            {
                string bumpName = ExcelWorksheet.GetCellValue(j, pinColumn);
                string fsDd = ExcelWorksheet.GetCellValue(j, fsDdColumn);
                string ioVol = ExcelWorksheet.GetCellValue(j, ioVolColumn);

                bumpName = FormatPinName(bumpName);
                if (bumpName != "" && (CellDiff.IsLiked(fsDd, ConFs) || CellDiff.IsLiked(fsDd, ConDd)) && !IsExistPin(bumpName))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_07, ExcelWorksheet.Name, j, 1, $"Column[{IoContiRow.ConHeaderBumpName}] The pin: {bumpName} does not exist in the PinMap", new string[] { IoContiRow.ConHeaderBumpName, bumpName });
                    result = false;
                }
                string voltage = ioVol.Replace("V", "").Replace("v", "").Replace("m", "").Replace(" ", "").Replace("\t", "");
                if (voltage.Length > 0)
                {
                    string pinGrpName = "Pins_" + voltage.Replace(".", "p") + "v";
                    pinGrpName = pinGrpName.Replace("0v", "v");
                    if (!_grpList.Contains(pinGrpName))
                    {
                        _grpList.Add(pinGrpName);
                    }
                }
            }

            return result;
        }

        private bool IsExistPin(string pinName)
        {
            bool result = false;
            foreach (Pin pin in TestProgram.IgxlWorkBk.PinMapPair.Value.PinList)
            {
                if (string.Equals(pin.PinName, pinName, StringComparison.CurrentCultureIgnoreCase))
                {
                    result = true;
                    break;
                }
            }

            return result;
        }

        private string FormatPinName(string pinName)
        {
            string targetPinName = pinName;
            if (targetPinName.Contains("["))
            {
                targetPinName = targetPinName.Replace("[", "");
            }
            if (targetPinName.Contains("]"))
            {
                targetPinName = targetPinName.Replace("]", "");
            }

            return targetPinName;
        }
    }
}
