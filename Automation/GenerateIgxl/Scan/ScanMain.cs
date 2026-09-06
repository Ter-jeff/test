using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.Scan.CPM;
using Automation.GenerateIgxl.Scan.Harvest;
using Automation.GenerateIgxl.Scan.NonBinCut;
using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.PreCheck.PreCheckManager;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Reader.TestPlan.ClockOut;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LogLib.Static;
using LogLib.Utility;

using OfficeOpenXml;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Scan;
using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Scan
{
    public class ScanMain : WorkFlowBase<ParaData>
    {
        private ScanInputData _scanInputData;

        public override bool PreCheckFlow(ParaData paraData)
        {
            try
            {
                _scanInputData = new ScanInputManager(EpWorkbook.TestPlanWorkbook).Read();

                if (!LocalSpecs.Options.BypassPreCheck)
                {
                    new ScanCheckManager(EpWorkbook.TestPlanWorkbook, paraData).PreCheckAll();
                }

                return true;
            }
            catch (Exception e)
            {
                Response.Report("Scan Action has errors : " + e.StackTrace, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
                return false;
            }
        }

        public override void WorkFlow(ParaData paraData)
        {
            try
            {
                if (EpWorkbook.ScghWorkbook != null)
                {
                    Response.Report("Generating Scan Characterization ...", percentage: 70);
                    GenSelSram(_scanInputData);
                }
                if (EpWorkbook.TestPlanWorkbook.Worksheets[NeededSheets.ClockPllMeas] != null)
                {
                    GenClockCheck(_scanInputData);
                }
                if (EpWorkbook.TestPlanWorkbook.Worksheets["TurboModeCheck"] != null)
                {
                    GenerateTurboModeCheckFlow(_scanInputData);
                    Response.Report("TurboModeCheck Completed!");
                }
                if (EpWorkbook.TestPlanWorkbook.Worksheets["Instance_CPM"] != null && LocalSpecs.IsModuleIncluded("CPM"))
                {
                    Response.Report("Generating Cpm Instance sheet ...", percentage: 5);
                    ScanConfig config = SettingStatic.ScanConfig;
                    var cpmInstanceFlow = new CpmInstanceMain(config);
                    cpmInstanceFlow.WorkFlow();
                    if (EpWorkbook.TestPlanWorkbook.Worksheets["EFUSE_cpm"] != null)
                    {
                        Response.Report("Generating CPM_Table ...", percentage: 5);
                        var cpmTable = new CpmTable();
                        cpmTable.WorkFlow(_scanInputData.EfuseCpmSheet);
                        Response.Report("CPM_Table Completed!");
                    }
                    Response.Report("CPM Completed!", percentage: 100);
                }
                if (TestPlanStatic.ScanInstanceSheets != null && TestPlanStatic.ScanInstanceSheets.Any() && LocalSpecs.IsModuleIncluded(BlockStatus.Scan))
                {
                    ScanConfig config = SettingStatic.ScanConfig;
                    var scanNonBinCutInstanceMain = new ScanNonBinCutInstanceMain(config);
                    Response.Report("Generating Non BinCut Instance sheet ...", percentage: 75);
                    scanNonBinCutInstanceMain.WorkFlow();
                }

                Response.Report("Scan Completed!", percentage: 100);

            }
            catch (Exception e)
            {
                string message = "Scan AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }

        private static void GenSelSram(ScanInputData scanInputData)
        {
            var selSramWriter = SelSramWriterSingleton.GetInstance();
            selSramWriter.PayloadTypeTable = scanInputData.ScanConfig.PayloadType;

            if (AcTSetCategoryMapSingleton.Instance().PatternList == null)
            {
                const string lStrMessage = "PatterList.csv does not exist!";
                Response.Report(lStrMessage, EnumMessageLevel.Warning, 35);
            }
            else
            {
                SelSramWriterSingleton.GetInstance().PatSetTimeSetDictionary = AcTSetCategoryMapSingleton.Instance().PatternList;
            }
        }

        private void GenClockCheck(ScanInputData scanInputData)
        {
            Response.Report("Generate ClockOut items ...", percentage: 10);

            ClockMeasSheet clockOutSheet = scanInputData.ClockMeasSheet;

            var binCutInstanceSheets = new List<BinCutInstanceSheet>();
            foreach (ExcelWorksheet sheet in EpWorkbook.TestPlanWorkbook.Worksheets)
            {
                if (!sheet.Name.ContainsIgnoreCase("clock"))
                {
                    continue;
                }

                if (sheet.Name.StartsWith("Instance_", StringComparison.CurrentCultureIgnoreCase))
                {
                    var binCutInstanceSheetReader = new BinCutInstanceSheetReader();
                    BinCutInstanceSheet binCutInstanceSheet = binCutInstanceSheetReader.ReadSheet(sheet);
                    if (binCutInstanceSheet != null)
                    {
                        binCutInstanceSheets.Add(binCutInstanceSheet);
                    }
                }
            }
            if (binCutInstanceSheets.Count > 0)
            {
                BinCutInstanceNamingSheet binCutInstanceNamingSheet = SettingStatic.BinCutInstanceNamingSheet;
                BinCutFinalInstanceRows binCutInstanceRows = new InstSheetPreProcess(scanInputData.ScanConfig).InitialInstance(binCutInstanceSheets, binCutInstanceNamingSheet);
                var writer = new ClockOutMeasGenerator(NeededSheets.ClockPllMeas, clockOutSheet.Rows, binCutInstanceRows);
                try
                {

                    writer.WorkFlow();
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }
        }
        public void GenerateTurboModeCheckFlow(ScanInputData scanInputData)
        {
            var instanceSheet = new InstanceSheet("TestInst_TurboModeCheck");
            var flowGroups = scanInputData.TurboModeInstanceSheet.Rows.GroupBy(x => x.FlowName).ToList();
            foreach (IGrouping<string, BinCutInstanceRow> flowGroup in flowGroups)
            {
                var flowSheet = new SubFlowSheet(flowGroup.Key, "TurboModeCheck:" + flowGroup.Key);
                foreach (BinCutInstanceRow instance in flowGroup)
                {
                    InstanceRow row = new InstanceRow
                    {
                        TestName = instance.Instance,
                        VbtType = ".NET",
                    };
                    Function function = TestProgram.VbtFunctionLib.GetFunctionByName(instance.FunctionName, "bincut", true);
                    row.VbtName = function.FullFunctionName;
                    row.ArgList = function.Parameters;
                    if (!string.IsNullOrEmpty(instance.UserFunction))
                    {
                        TestPlanStatic.UserFunctionSheet.ArgumentSetting(instance.UserFunction, function);
                    }
                    row.Args = function.ArgList;
                    instanceSheet.Rows.Add(row);

                    flowSheet.AddRow(new FlowRow { Opcode = OpCode.If, Parameter = instance.SiteVar, PassAction = instance.PassFlag, FailAction = instance.FailFlag });
                    flowSheet.AddRow(new FlowRow { Opcode = OpCode.Test, Parameter = instance.Instance });
                    flowSheet.AddRow(new FlowRow { Opcode = OpCode.EndIf });
                }
                flowSheet.AddReturnRow();
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirNonBinCut, flowSheet);
            }

            if (instanceSheet.Rows.Count != 0)
            {
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirNonBinCut, instanceSheet);
            }
        }
    }
}
