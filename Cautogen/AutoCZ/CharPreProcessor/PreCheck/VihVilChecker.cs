using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class VihVilChecker : PreCheckBase
    {
        /* properties */
        private readonly HashSet<string> _hardIpSheets = new HashSet<string>();

        /* constructor */
        public VihVilChecker(HashSet<string> hardIpSheets)
        {
            _hardIpSheets = hardIpSheets;
        }

        /* methods */
        public override void Check(List<Characterization> charList, string fileName)
        {
            // for item in hard ip sheets and tp name contains vih | vil, their user def 1 should be HFH | HFL | HIO
            var rgxVihVil = new Regex("vih|vil", RegexOptions.IgnoreCase | RegexOptions.Compiled);
            var rgxHflh = new Regex("HFL|HFH|HIO", RegexOptions.IgnoreCase | RegexOptions.Compiled);
            IEnumerable<Characterization> items = from c in charList
                                                  where _hardIpSheets.Contains(c.SheetName.ToUpper())  // for items in HIP sheet
                                                  where rgxVihVil.IsMatch(c.TpName)  // and tp name contains vih | vil
                                                  where !rgxHflh.IsMatch(c.UserDef1)  // but user def 1 is not HFL | HIO | HFL
                                                  select c;

            // raise error 
            foreach (Characterization item in items)
            {
                string errMessage = "TpName Contains VIH/VIL but USERDEF1 is not HFL|HFH|HIO in " + item.SheetName + "  Row: " + item.RowNum;
                ErrorManager.AddError(ErrorType.WrongUserdef1OfVih, item.SheetName, item.RowNum,
                    item.ColNum(new List<string> { "userdef1", "tpname(teusesuffix)" }), item.Use, errMessage);
                foreach (int col in item.ColNum(new List<string> { "userdef1", "tpname(teusesuffix)" }))
                {
                    ErrorReportManager.AddError(CharErrorType.E_WrongUserdef1OfVih_01, item.SheetName, item.RowNum, col, [item.SheetName, $"{item.RowNum}"]);
                }
            }
        }
    }
}
