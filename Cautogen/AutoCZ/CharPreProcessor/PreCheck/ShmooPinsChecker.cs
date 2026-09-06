using System;
using System.Collections.Generic;

using Automation.Const;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class ShmooPinsChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charRows, string sheetName)
        {
            foreach (Characterization item in charRows)
            {

                // does not allow in VDD fillin in pin column?
                foreach (Pin pin in item.Pins)
                {
                    if (!pin.PinName.Contains("VDD", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    string outString = "Error VDD pins in Pin column in " + item.SheetName + " Row " + item.RowNum;
                    ErrorManager.AddError(ErrorType.WrongVddInPinColumn, item.SheetName, item.RowNum, pin.PlanIndex, item.Use, outString, pin.PinName);
                    ErrorReportManager.AddError(CharErrorType.E_WrongVddInPinColumn_01, item.SheetName, item.RowNum, pin.PlanIndex, [item.SheetName, $"{item.RowNum}"],
                        new ErrorInfo() { Comments = new List<string>() { pin.PinName } });
                    break;
                }

                // if missing pinmap file, ignore missing pins check
                if (UtilityMain.UtilityData.InputParam.PinMapFile == "")
                {
                    if (!UtilityMain.UtilityData.InputParam.CharPreCheckWithoutTp)
                    {
                        var missingPinMap = new ErrorMessage
                        {
                            Message = "Missing PinMap.txt Please Reload PinMap or Test Program!!!",
                            ErrorType = ErrorType.MissingPinMap,
                            ErrorLevel = ErrorLevel.Error
                        };
                        ErrorManager.Add(missingPinMap);
                        ErrorReportManager.AddError(CharErrorType.E_MissingPinMap_01, "", 0, 0, []);

                    }
                    continue;
                }

                foreach (Pin pin in item.Pins)
                {
                    if (pin.PinType.ToUpper() == "AC_SPEC" || pin.PinType.ToUpper() == "AC SPEC" ||
                        UtilityMain.UtilityFunction.PinExistInPinMap(pin.PinName))
                    {
                        continue;
                    }

                    string errMessage = "Missing pin in pinmap file in " + item.SheetName + " Row " + item.RowNum;
                    ErrorManager.AddError(ErrorType.MissingPinName, item.SheetName, item.RowNum, pin.PlanIndex, item.Use, errMessage, pin.PinName);
                    ErrorReportManager.AddError(CharErrorType.E_MissingPinName_02, item.SheetName, item.RowNum, pin.PlanIndex, [item.SheetName, item.RowNum.ToString()],
                                        new ErrorInfo() { Comments = new List<string>() { pin.PinName } });
                }
            }
        }
    }
}
