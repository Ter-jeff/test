using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.PreCheck.AllParaData;
using Automation.PreCheck.PreChecks.Basic;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;
using Automation.Static.Result;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.DataStruct;
using TestPlanLib.Static;

using BasicInputData = Automation.InputManager.Data.BasicInputData;

namespace Automation.PreCheck.PreCheckManager
{
    public class BasicPreCheckManager : PreCheckManager<ParaData>
    {
        private readonly BasicInputData _basicInputData;
        private static readonly Regex _regex = new Regex(@"^\w+::\w+$");

        public BasicPreCheckManager(ExcelWorkbook excelWorkbook, ParaData paraData, BasicInputData basicInputData) : base(excelWorkbook, paraData)
        {
            _basicInputData = basicInputData;
        }

        public override bool PreCheckAll()
        {
            foreach (ExcelWorksheet worksheet in EpWorkbook.TestPlanWorkbook.Worksheets)
            {
                string sheetName = worksheet.Name;
                if (sheetName.StartsWith(NeededSheets.ContiDcTest, StringComparison.OrdinalIgnoreCase))
                {
                    if (EpWorkbook.TestPlanWorkbook.Worksheets[sheetName].Cells[1, 1].Text.Equals("DcTest_Continuity_rev1", StringComparison.CurrentCultureIgnoreCase) || EpWorkbook.TestPlanWorkbook.Worksheets[sheetName].Cells[1, 1].Text.Equals("DcTest_Continuity_rev2", StringComparison.CurrentCultureIgnoreCase))
                    {
                        BasicResult.CheckDcTestContinuity = true;
                    }
                    else if (EpWorkbook.TestPlanWorkbook.Worksheets[sheetName].Cells[1, 1].Text.Equals("DcTest_Continuity_rev3", StringComparison.CurrentCultureIgnoreCase))
                    {
                        BasicResult.CheckDcTestContinuity = true;
                    }
                    else if (EpWorkbook.TestPlanWorkbook.Worksheets[sheetName].Cells[1, 1].Text.Equals("DcTest_Continuity_rev4", StringComparison.CurrentCultureIgnoreCase))
                    {
                        BasicResult.CheckDcTestContinuity = true;
                    }
                }
                else if (sheetName.Equals(NeededSheets.ContiIo, StringComparison.OrdinalIgnoreCase))
                {
                    var basicPreCheckIoContinuity = new BasicPreCheckIoContinuity(ExcelWorkbook, sheetName);
                    BasicResult.CheckIoContinuity = basicPreCheckIoContinuity.CheckMain();
                }
                else if (sheetName.Equals(NeededSheets.PowerInfo, StringComparison.OrdinalIgnoreCase))
                {
                    var basicPreCheckPowerInfo = new BasicPreCheckPowerInfo(ExcelWorkbook, sheetName);
                    BasicResult.CheckPowerInfo = basicPreCheckPowerInfo.CheckMain();
                }
                else if (sheetName.Equals(NeededSheets.IoInfo, StringComparison.OrdinalIgnoreCase))
                {
                    var basicPreCheckIoInfo = new BasicPreCheckIoInfo(ExcelWorkbook, sheetName);
                    BasicResult.CheckIoInfo = basicPreCheckIoInfo.CheckMain();
                }
            }

            BasicResult.CheckCsvPattenList = new BasicPreCheckPattenListCsv(LocalSpecs.PatternListCsvFileName).CheckMain();
            BasicResult.HasDirPatSet = new BasicDirectoryCheck(LocalSpecs.PatternFolder).CheckMain();
            BasicResult.CheckDirTimeSet = new BasicDirectoryCheck(LocalSpecs.TimeSetFolder).CheckMain();
            BasicResult.HasTimeSet = new BasicPrecheckTimeSet(LocalSpecs.PatternListCsvFileName, LocalSpecs.TimeSetFolder).CheckMain();

            CheckConti(_basicInputData.DcTestContiSheets, TestProgram.IgxlWorkBk.PinMapPair.Value);

            CheckPowerInfoSheet(TestPlanStatic.PowerInfoSheet, MultiTestSettingSheetsSingleton.Instance());

            CheckTestSettingPinExistInPinMap(TestProgram.IgxlWorkBk.PinMapPair.Value, MultiTestSettingSheetsSingleton.Instance());

            var timeSetChecker = new TimeSetChecker();
            timeSetChecker.CheckTimeSet(_basicInputData.MultiTimeSetSheets);

            return false;
        }

        private void CheckPowerInfoSheet(PowerInfoSheet powerInfoSheet, MultiTestSettingSheetsSingleton multiTestSettingSheetsSingleton)
        {
            List<string> volTablePinList = multiTestSettingSheetsSingleton.PowerPinList;
            List<ProtocolAwarePin> nWirePin = NwireSingleton.Instance().SettingInfo.NwirePins;
            foreach (PowerInfoRow row in powerInfoSheet.Rows.Where(x => x.IsPowerPin))
            {
                string pinName = row.PinName;
                if (_regex.IsMatch(pinName))
                {
                    pinName = pinName.Split(':')[0] + "_Diff";
                }
                if (!volTablePinList.Contains(pinName) && !nWirePin.Any(x => pinName.Contains(x.OutClk)))
                {
                    string message = $"Missing pin in voltage table: {pinName}";
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_17, NeededSheets.PowerInfo, row.RowNum, 0, $"Missing pin in voltage table: {pinName}", new string[] { pinName }, new ErrorInfo() { Comments = new List<string>() { pinName } });
                }
            }
        }

        private void CheckTestSettingPinExistInPinMap(PinMapSheet pinMap, MultiTestSettingSheetsSingleton multiTestSettingSheetsSingleton)
        {
            if (!pinMap.PinList.Any() && !pinMap.GroupList.Any())
            {
                return;
            }

            foreach (string powerPin in multiTestSettingSheetsSingleton.PowerPinList)
            {
                if (!pinMap.IsGroupExist(powerPin) && !pinMap.IsPinExist(powerPin))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_21, NeededSheets.PinMap, 1, 1, $"Can't generate power level for pin/pin group: {powerPin} from voltage table, if it just be used to crete variable in DC Specs, please ignore this or make sure {NeededSheets.PinMap} had been define this pin/pin group!!!", new string[] { powerPin, NeededSheets.PinMap });
                }
            }
        }

        private void CheckConti(List<DcTestContiSheet> dcTestContiSheets, PinMapSheet pinMapSheet)
        {
            foreach (DcTestContiSheet dcTestContiSheet in dcTestContiSheets)
            {
                foreach (DcTestContiRow dcTestContiRow in dcTestContiSheet.DcTestContiRows)
                {
                    List<string> pinList = dcTestContiRow.PinGroup.Split(';', ',').ToList();
                    foreach (string pin in pinList)
                    {
                        if (pinMapSheet.GetGroup(pin) == null &&
                            pinMapSheet.GetPin(pin) == null &&
                            !dcTestContiRow.Category.Equals("UserFunction", StringComparison.OrdinalIgnoreCase) &&
                            !dcTestContiRow.Category.Equals("Opcode", StringComparison.OrdinalIgnoreCase))
                        {
                            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_16, NeededSheets.ContiDcTest, 1, 1, $"Missing conti pin/pin group: {pin} in pin map!", new string[] { pin });
                        }
                    }
                }
            }
        }
    }
}
