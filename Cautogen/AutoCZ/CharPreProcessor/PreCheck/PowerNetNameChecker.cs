using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class PowerNetNameChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charRows, string sheetName)
        {
            // early return if there input contains no power merge
            if (UtilityMain.UtilityData.PowerMergeResult.Rows.Count == 0)
            {
                return;
            }

            var allNetNames =
                UtilityMain.UtilityData.PowerMergeResult.AsEnumerable().Select(a => a.Field<string>("Master")).ToList();

            // Precheck VDD power on the header is not net name.
            Characterization charItem = charRows[0];
            //foreach (var errorMessage in 
            //    from powerPin in charItem.PowerSupplyX 
            //    where !allNetNames.Exists(
            //    pin => pin.Replace("_", "").Equals(powerPin.Name, StringComparison.OrdinalIgnoreCase)) 
            //    select "Header power name \"" + powerPin.Name + "\" is not a Net Name")
            //{
            //    ErrorManager.AddError(ErrorType.WrongNetName, charItem.SheetName, 1, 
            //        charItem.ColNum(),"Use", errorMessage);
            //}
            foreach (ShmooSpec powerSet in charItem.PowerSupplyX)
            {
                if (!allNetNames.Exists(
                pin => pin.Replace("_", "").Equals(powerSet.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    ErrorManager.AddError(ErrorType.WrongNetName, charItem.SheetName, 1,
                    powerSet.ColIndex, "Use", "Header power name \"" + powerSet.Name + "\" is not a Net Name");
                    ErrorReportManager.AddError(CharErrorType.E_WrongNetName_01, charItem.SheetName, 1, powerSet.ColIndex, [powerSet.Name]);
                }
            }


            // Precheck VDD power on USERDEF3 is not net name
            foreach (Characterization charPlan in charRows)
            {
                if (!Regex.IsMatch(charPlan.UserDef3, "^VDD"))
                {
                    continue;
                }

                if (
                    allNetNames.Exists(
                        pin => pin.Replace("_", "").Equals(charPlan.UserDef3, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                string errorMessage = "Header power name \"" + charPlan.UserDef3 + "\" is not a Net Name";
                ErrorManager.AddError(ErrorType.WrongNetName, charPlan.SheetName, charPlan.RowNum, charPlan.ColNum("userdef3"), "Use", errorMessage);
                foreach (int col in charPlan.ColNum("userdef3"))
                {
                    ErrorReportManager.AddError(CharErrorType.E_WrongNetName_01, charPlan.SheetName, charPlan.RowNum, col, [charPlan.UserDef3]);
                }
            }

            #region Precheck VDD power in the hardip pattern info is not net name.
            // Maybe can check "UserDef4" to instead of checking patInfo.
            // If charplan uses net name but patInfo not, it can be checked by "PinSeqChecker"
            #endregion
        }
    }
}
