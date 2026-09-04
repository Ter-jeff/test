using System;
using System.Collections.Generic;

using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport.Interop
{
    public class ErrorReportFactory
    {
        public static ErrorReportInterop GetReport(Type type, List<Error> errorList)
        {
            if (type == HardIpErrorType.DupBitName.GetType())
            {
                return new ReportHardIp(errorList);
            }

            if (type == DuplicateInstance.Duplicate.GetType())
            {
                return new ReportDuplicateInstance(errorList);
            }

            if (type == UnusedDcCategory.UnusedDcCategory.GetType())
            {
                return new ReportUnusedDcCategory(errorList);
            }

            if (type == ZeroVoltageDcCategory.ZeroVoltageDcCategory.GetType())
            {
                return new ReportZeroVoltageDcCategory(errorList);
            }

            if (type == PatValtPinCheckerType.ValueMissingInCategory.GetType())
            {
                return new ReportPatValtPinValueMissInCategory(errorList);
            }

            return new ReportDefault(errorList);
        }
    }
}
