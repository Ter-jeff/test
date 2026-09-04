using System.Collections.Generic;

using Microsoft.Office.Interop.Excel;

namespace CommonLib.ErrorReport.Interop
{
    internal class ReportDefault : ErrorReportInterop
    {
        public ReportDefault(List<Error> errorList)
            : base(errorList)
        {
        }

        public override void Write(Workbook workbook)
        {
            if (ErrorList.Count == 0)
            {
                return;
            }

            WriteReport(workbook);
        }
    }
}
