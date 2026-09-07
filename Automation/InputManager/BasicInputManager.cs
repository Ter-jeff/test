using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.DataStruct;
using TestPlanLib.Static;

using BasicInputData = Automation.InputManager.Data.BasicInputData;

namespace Automation.InputManager
{
    public class BasicInputManager : InputManagerBase<BasicInputData>
    {
        private static readonly Regex _regex1 = new Regex("MuxPinSheet", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex2 = new Regex(@"(?<num>\d*)\.(?<digi>\d*)");

        public BasicInputManager(ExcelWorkbook excelWorkbook) : base(excelWorkbook)
        {
        }

        public override BasicInputData Read()
        {
            var result = new BasicInputData();

            if (!string.IsNullOrEmpty(NeededSheets.PSet))
            {
                foreach (ExcelWorksheet ws in ExcelWorkbook.Worksheets.Where(p => Regex.IsMatch(p.Name, NeededSheets.PSet, RegexOptions.IgnoreCase)))
                {
                    ws.ExportToTxt(Path.Combine(FolderStructure.DirCommon, ws.Name + ".txt"));
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, ws.Name);
                }
            }

            if (LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                ExcelWorksheet ioLevelsSheet = ExcelWorkbook.Worksheets[NeededSheets.IoLevels];
                if (ioLevelsSheet != null)
                {
                    result.IoLevelsSheet = new IoLevelsSheetReader().ReadSheet(ioLevelsSheet);
                }
            }

            ExcelWorksheet ifoldPwrTable = ExcelWorkbook.Worksheets[NeededSheets.IfoldPwrTable];
            if (ifoldPwrTable != null)
            {
                result.IfoldPowerTableSheet = new IfoldPowerTableReader().ReadSheet(ifoldPwrTable);
            }

            ExcelWorksheet ioContiSheet = ExcelWorkbook.Worksheets[NeededSheets.ContiIo];
            if (ioContiSheet != null)
            {
                result.IoContiSheet = new IoContiReader().ReadSheet(ioContiSheet);
            }

            if (result.IoContiSheet != null && TestPlanStatic.IoInfoSheet != null)
            {
                CheckIoInfoLevelExists(result.IoContiSheet, TestPlanStatic.IoInfoSheet, TestPlanStatic.IoInfoConcurrentSheet);
            }
            if (result.IoContiSheet != null)
            {
                CompareCpFtPins(result.IoContiSheet);
            }

            result.PinMapSheet = TestProgram.IgxlWorkBk.PinMapPair.Value;

            foreach (ExcelWorksheet ws in EpWorkbook.TestPlanWorkbook.Worksheets)
            {
                if (_regex1.IsMatch(ws.Name))
                {
                    if (TestProgram.NonIgxlSheetsList.SheetList.Exists(p => Path.GetFileNameWithoutExtension(p).Equals(ws.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    ws.ExportToTxt(Path.Combine(FolderStructure.DirConti, ws.Name + ".txt"));
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirConti, ws.Name);
                }
            }

            #region conti
            foreach (ExcelWorksheet sheet in EpWorkbook.TestPlanWorkbook.Worksheets)
            {
                string sheetName = sheet.Name;
                if (Regex.IsMatch(sheetName, NeededSheets.ContiDcTest, RegexOptions.IgnoreCase))
                {
                    if (sheet.Cells[1, 1].Text.Trim().Equals("DcTest_Continuity_rev4", StringComparison.CurrentCultureIgnoreCase))
                    {
                        DcTestContiSheet dcTestContiSheet = new DcTestContiReaderRev4().ReadSheet(sheet);
                        result.DcTestContiSheets.Add(dcTestContiSheet);
                    }
                }
            }
            #endregion

            var patList = AcTSetCategoryMapSingleton.Instance().PatternList.Select(x => x.Value).ToList();
            result.MultiTimeSetSheets = new TimeSetGenerator().GenerateFlow(patList, LocalSpecs.TimeSetFolder, FolderStructure.DirTimings);

            return result;
        }

        private void CompareCpFtPins(IoContiSheet ioContiSheet)
        {
            var ftPinList = ioContiSheet.Rows.Select(p => p.BallName).Distinct().ToList();
            var cpPinList = ioContiSheet.Rows.Select(p => p.BumpName).Distinct().ToList();
            foreach (string pin in ftPinList)
            {
                if (!cpPinList.Any(p => p.Equals(pin, StringComparison.OrdinalIgnoreCase)))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_22, ioContiSheet.SheetName, 1, 1, $"BallName : {pin} Not exist in BumpName!!!", new string[] { pin });
                }
            }
        }

        private void CheckIoInfoLevelExists(IoContiSheet ioContiSheet, IoInfoSheet ioInfoSheet, IoInfoSheet ioInfoConcurrentSheet)
        {
            var groupByVoltage = ioContiSheet.Rows.GroupBy(x => x.IoVoltage).ToList();
            foreach (IGrouping<string, IoContiRow> voltage in groupByVoltage)
            {
                string numTok = _regex2.Match(voltage.Key).Groups["num"].ToString();
                string digTok = _regex2.Match(voltage.Key).Groups["digi"].ToString();
                string pinGroupName = $"Pins_{numTok}p{digTok}v";

                bool isExistData = ioInfoSheet.GetBlockPins("Common").Contains(pinGroupName);
                if (!isExistData)
                {
                    int elementRow = voltage.First() != null ? voltage.First().RowNum : 0;
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_04, ioContiSheet.SheetName, elementRow, 1, $"Should define io infos in {NeededSheets.IoInfo} for {pinGroupName}!!!", new string[] { NeededSheets.IoInfo, pinGroupName });
                }

                if (ioInfoConcurrentSheet == null)
                {
                    continue;
                }

                isExistData = ioInfoConcurrentSheet.GetBlockPins("Common").Contains(pinGroupName);
                if (!isExistData)
                {
                    int elementRow = voltage.First() != null ? voltage.First().RowNum : 0;
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_05, ioContiSheet.SheetName, elementRow, 1, $"Should define io infos in {NeededSheets.IoInfoConcurrent} for {pinGroupName}!!!", new string[] { NeededSheets.IoInfoConcurrent, pinGroupName });
                }
            }
        }
    }
}
