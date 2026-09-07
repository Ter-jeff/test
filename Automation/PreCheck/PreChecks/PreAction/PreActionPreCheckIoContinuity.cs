using System.Collections.Generic;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

using OfficeOpenXml;

using TestPlanLib.DataStruct;

namespace Automation.PreCheck.PreChecks.PreAction
{
    public class PreActionPreCheckIoContinuity : PreActionCheckBase
    {
        internal readonly List<string> _grpList = new List<string>();

        public PreActionPreCheckIoContinuity(ExcelWorkbook excelWorkbook, string sheetName) : base(excelWorkbook, sheetName)
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
            int fsDdColumn = 0;
            int ioVolColumn = 0;
            for (int i = StartColumn; i <= ExcelWorksheet.Dimension.End.Column; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i);
                if (CellDiff.IsLiked(header, IoContiRow.ConHeaderBumpName))
                {
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

            bool result = true;
            for (int j = StartRow + 1; j <= ExcelWorksheet.Dimension.End.Row; j++)
            {
                string fsDd = ExcelWorksheet.GetCellValue(j, fsDdColumn);
                string ioVol = ExcelWorksheet.GetCellValue(j, ioVolColumn);

                if (!Regex.IsMatch(fsDd, "(FS|DD)", RegexOptions.IgnoreCase))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_FSDDIssue_01, ExcelWorksheet.Name, j, fsDdColumn, "Needs to fill FS or DD in FS/DD Column !!!");
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
    }
}
