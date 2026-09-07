using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class TestNameChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charRows, string sheetName)
        {
            foreach (Characterization item in charRows.Where(item => item.TpName.Length > 255))
            {
                const string message = "Test Name length over 255 bytes";
                ErrorManager.AddError(ErrorType.TestNameOverLength, item.SheetName, item.RowNum, item.ColNum("tpname(teusesuffix)"), item.Use, message);
                foreach (int col in item.ColNum("tpname(teusesuffix)"))
                {
                    ErrorReportManager.AddError(CharErrorType.E_TestNameOverLength_01, item.SheetName, item.RowNum, col, []);
                }
            }

            foreach (Characterization item in charRows.Where(item => !Regex.IsMatch(item.TpName, "[^_]_{1}$")))
            {
                const string message = "Test Name not ends with single underline";
                ErrorManager.AddError(ErrorType.TestNameNotEndswithSingleUnderline, item.SheetName, item.RowNum, item.ColNum("tpname(teusesuffix)"), item.Use, message);
                foreach (int col in item.ColNum("tpname(teusesuffix)"))
                {
                    ErrorReportManager.AddError(CharErrorType.E_TestNameNotEndswithSingleUnderline_01, item.SheetName, item.RowNum, col, []);
                }
            }
        }
    }
}
