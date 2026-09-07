using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.Const;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class HacChecker : PreCheckBase
    {
        /* for HAC and HIO, meas2 is the meas type, and meas 3 is the */
        public override void Check(List<Characterization> charList, string sheetName)
        {
            foreach (Characterization item in charList)
            {
                switch (item.UserDef1.ToLower())
                {
                    case "hac":
                        // hac should have correct meas type in user_def2
                        if (!_IsCorrectUserDef2MeasType(item.UserDef2))
                        {
                            ErrorManager.AddError(ErrorType.WrongMeasOfHac, item.SheetName, item.RowNum, item.ColNum("userdef2"),
                                item.Use, "Wrong Meas for HAC in " + item.SheetName + "  Row: " + item.RowNum);
                            foreach (int col in item.ColNum("userdef2"))
                            {
                                ErrorReportManager.AddError(CharErrorType.E_WrongMeasOfHac_01, item.SheetName, item.RowNum, col, ["HAC", item.SheetName, $"{item.RowNum}"]);
                            }
                        }
                        else
                        {
                            _CheckHacUserDef3(item);
                            _IsPinInPinMap(item);
                        }
                        break;

                    case "hio":
                        // hio should have correct meas type in user_def2
                        if (!_IsCorrectUserDef2MeasType(item.UserDef2))
                        {
                            ErrorManager.AddError(ErrorType.WrongMeasOfHac, item.SheetName, item.RowNum, item.ColNum("userdef2"),
                                item.Use, "Wrong MeasType for HIO in " + item.SheetName + "  Row: " + item.RowNum);
                            foreach (int col in item.ColNum("userdef2"))
                            {
                                ErrorReportManager.AddError(CharErrorType.E_WrongMeasOfHac_01, item.SheetName, item.RowNum, col, ["HIO", item.SheetName, $"{item.RowNum}"]);
                            }
                        }

                        break;
                }
            }
        }

        private static void _IsPinInPinMap(Characterization item)
        {
            // If missing pin map file, ignore missing pin check
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
                    ErrorReportManager.AddError(CharErrorType.E_MissingPinMap_01, "Error", 0, 0, []);
                }
                return;
            }

            if (item.UserDef2.ToLower() != "measc" &&
                !UtilityMain.UtilityFunction.PinExistInPinMap(item.UserDef4))
            {
                ErrorManager.AddError(ErrorType.MissingPinName, item.SheetName, item.RowNum, item.ColNum("userdef4"), item.Use,
                    "Missing (pin name)/(pin group) in PinMap.txt/pinList sheet", item.UserDef4);

                foreach (int col in item.ColNum("userdef4"))
                {
                    ErrorReportManager.AddError(CharErrorType.E_MissingPinName_01, item.SheetName, item.RowNum, col, [],
                        new ErrorInfo() { Comments = new List<string>() { item.UserDef4 } });
                }
            }
        }

        private static void _CheckHacUserDef3(Characterization item)
        {
            string userDef3 = item.UserDef3;
            if (userDef3 == "H")
            {
                return;
            }

            if (userDef3 == "N")
            {
                return;
            }

            if (userDef3 == "L")
            {
                return;
            }

            if (userDef3 == "X")
            {
                return;
            }

            if (userDef3 == "Multi")
            {
                return;
            }

            if (Regex.IsMatch(userDef3, "^[0-9][a-z|A-Z|0-9]+|d+$"))
            {
                return;
            }

            ErrorManager.AddError(ErrorType.WrongUserdef3OfHac, item.SheetName, item.RowNum, item.ColNum("userdef3"), item.Use,
                "Wrong USERDEF3 for HAC in " + item.SheetName + "  Row: " + item.RowNum);
            foreach (int col in item.ColNum("userdef3"))
            {
                ErrorReportManager.AddError(CharErrorType.E_WrongUserdef3OfHac_01, item.SheetName, item.RowNum, col, [item.SheetName, $"{item.RowNum}"]);
            }
        }

        private static bool _IsCorrectUserDef2MeasType(string userDef2)
        {
            userDef2 = userDef2.ToLower();
            if (userDef2 == "measi")
            {
                return true;
            }

            if (userDef2 == "measr")
            {
                return true;
            }

            if (userDef2 == "measv")
            {
                return true;
            }

            if (userDef2 == "measf")
            {
                return true;
            }

            if (userDef2 == "measz")
            {
                return true;
            }

            if (userDef2 == "measc")
            {
                return true;
            }

            return false;
        }
    }
}
