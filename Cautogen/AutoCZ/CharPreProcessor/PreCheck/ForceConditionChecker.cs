using System.Collections.Generic;
using System.Linq;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    /* For USERDEF6 contains Vih|Vil|Vicm|Vid, should also specified the pin sweep with the corresponding type */
    public class ForceConditionChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            foreach (Characterization item in charList.Where(IsUseItem).Where(item => !_IsCheckPass(item)))
            {
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Error,
                    ErrorType = ErrorType.MissingForceCondition,
                    SheetName = sheetName,
                    RowNum = item.RowNum,
                    Message = "USERDEF6 contains Vih|Vil|Vicm|Vid, But do not has pin sweep " + sheetName,
                });
                ErrorReportManager.AddError(CharErrorType.E_MissingForceCondition_01, sheetName, item.RowNum, 0, [sheetName]);
            }
        }

        private static bool _IsCheckPass(Characterization item)
        {
            if (item.UserDef6.Contains("VIH"))
            {
                return item.Pins.Any(pin => pin.PinType.ToUpper() == "VIH");
            }

            if (item.UserDef6.Contains("VIL"))
            {
                return item.Pins.Any(pin => pin.PinType.ToUpper() == "VIL");
            }

            if (item.UserDef6.Contains("VICM"))
            {
                return item.Pins.Any(pin => pin.PinType.ToUpper() == "VICM");
            }

            if (item.UserDef6.Contains("VID"))
            {
                return item.Pins.Any(pin => pin.PinType.ToUpper() == "VID");
            }

            return true;
        }
    }
}
