using System.Collections.Generic;

using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using LogLib.Static;

namespace Automation.PreCheck.PreChecks.Basic
{
    public class TimeSetChecker
    {
        private TimeSetSheets _multiTimeSetSheets;

        public void CheckTimeSet(TimeSetSheets multiTimeSet)
        {
            if (multiTimeSet == null)
            {
                return;
            }

            _multiTimeSetSheets = multiTimeSet;

            CheckMissingGroupName();

            CheckMultiShiftTSet();
        }

        private void CheckMissingGroupName()
        {
            PinMapSheet pinMapSheet = TestProgram.IgxlWorkBk.PinMapPair.Value;
            if (pinMapSheet == null)
            {
                return;
            }

            foreach (ComTimeSetBasicSheet sheet in _multiTimeSetSheets)
            {
                foreach (TSet timeSet in sheet.Rows)
                {
                    foreach (TimingRow timingRow in timeSet.TimingRows)
                    {
                        string groupName = timingRow.PinGrpName;
                        if (!pinMapSheet.IsPinExist(groupName) && !pinMapSheet.IsGroupExist(groupName))
                        {
                            //string outString = string.Format("Pin (group):{2} in Time Set file:{0};TSet:{1} is missing. ", sheet.Name, timeSet.Name, groupName);
                            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_20, timeSet.Name, 1, 0, string.Format("Pin (group):{2} in Time Set file:{0};TSet:{1} is missing. ", sheet.Name, timeSet.Name, groupName), new string[] { sheet.Name, timeSet.Name, groupName }, new ErrorInfo() { Comments = new List<string>() { groupName } });
                        }

                    }
                }
            }
        }

        private void CheckMultiShiftTSet()
        {
            foreach (ComTimeSetBasicSheet igxlSheet in _multiTimeSetSheets)
            {
                if (igxlSheet.IsMultiShiftInTSet)
                {
                    string alarmStr = $"Multi-ShiftInFreq TimeSet Sheet problem @ {igxlSheet.Name} : {igxlSheet.GetMultiShiftInStr}";
                    igxlSheet.InsertAlarmDataInFirstRow(alarmStr);
                    Response.Report(alarmStr, EnumMessageLevel.Error, 100);
                    ErrorReportManager.AddError(BasicErrorType.E_FormatError_22, igxlSheet.Name, 1, 0, $"Multi-ShiftInFreq TimeSet Sheet problem @ {igxlSheet.Name} : {igxlSheet.GetMultiShiftInStr}", new string[] { igxlSheet.Name, igxlSheet.GetMultiShiftInStr }, new ErrorInfo() { Comments = new List<string>() { "Multi-ShiftInFreq" } });
                }
            }
        }
    }
}
