using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.PostAction.GenMainFlow;
using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.GenerateIgxl.PostAction.GenMainFlow.Business;
using Automation.GenerateIgxl.PostAction.Relay;
using Automation.GenerateIgxl.PostAction.SelSram;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using TestPlanLib.Singleton;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class MainFlowMain1Tests : FunctionTestBase
    {
        [DataTestMethod]
        [DataRow("IGXL", "MySheet", "SubFlow", "IGXL_Branch")]
        [DataRow("OPCODE", "SomeOpcode", "SubFlow", "OPCODE_Branch")]
        [DataRow("CHAR", "CharSheet", "CharSub", "CHAR_Branch")]
        [DataRow("CUSTOM", "FlowA", "", "CUSTOM_Branch")]
        [DataRow("Relay", "Relay", "RelayTest", "Relay_Branch")]
        [DataRow("UF_Instance", "UF_Instance", "UFTest", "UF_Instance_Branch")]
        [DataRow("HarvestingTruthTable_A", "HarvestingTruthTable_A", "", "Harvest_Branch")]
        [DataRow("Bintable", "Bin_A_B_C", "", "Bintable_Logical", "F_A&&!F_B&&F_C")]
        [DataRow("Bintable", "Bin_A", "", "Bintable_WithoutLogical", "F_A")]
        public void GenMainFlow_BranchTests(string source, string sheetName, string subFlow, string caseName, string option = "ID: 5001")
        {
            string subName = "MainFlowGenBranch";
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            TestProgram.Clear();
            LocalSpecs.Options.Device = EnumDevice.LCD;
            BlockStatus.GetAutomationBlockStatus(BlockStatus.Otp).Down = true;
            LocalSpecs.TarFolder = outputPath;

            var seq = new FlowSequenceNew
            {
                Enable = true,
                Source = source,
                Module = source,
                SheetName = sheetName,
                SubFlowName = subFlow,
                OriSheetName = "MySheet",
                Comment = "Extra.txt",
                FailFlag = "Bin1",
                Option = option
            };
            var flowMain = new MainFlowBase { JobName = "JOB1", SequencesNew = [seq] };
            var subFlowSheets = new Dictionary<string, SubFlowSheet>();
            var insSheets = new Dictionary<string, InstanceSheet>();
            var mainFlowMain = new MainFlowMain(null, subFlowSheets, insSheets);

            // Act
            SubFlowSheet result = mainFlowMain.GenMainFlow(flowMain);

            // Assert
            Assert.IsTrue(result.Rows.Count > 0, $"{caseName}: Expected rows were not generated.");
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void MainFlowGenTest()
        {
            string subName = "MainFlowGen";
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            TestProgram.Clear();
            HardIpStatic.Clear();
            HardIpStatic.FlowUsedInteger.Add("SrcCodeIndx1");
            LocalSpecs.TarFolder = outputPath;
            BlockStatus.GetAutomationBlockStatus(BlockStatus.Efuse).Down = true;
            LocalSpecs.SetEnableModules([BlockStatus.Efuse]);
            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);
            SelSramPatternSingleton.GetInstance().AddData(new SelSramData("pcSheetName", "pattern", "ConHc"));

            // Arrange
            var subFlowSheets = new Dictionary<string, SubFlowSheet>
            {
                {"Flow_efuse_BankRead", new SubFlowSheet("Flow_efuse_BankRead") { JobNames = ["CP1"] } },
                {"Flow_efuse_EVS_All_Blank_Check", new SubFlowSheet("Flow_efuse_EVS_All_Blank_Check")},
                {"module", new SubFlowSheet("module")},
                {"ShiftInVs", new SubFlowSheet("ShiftInVs")},
                {"Flow_Post_IDS_EQN_Voltage", new SubFlowSheet("Flow_Post_IDS_EQN_Voltage")},
                {"Flow_nWire_XXX", new SubFlowSheet("Flow_nWire_XXX")},
                {"Flow_CZ2", new SubFlowSheet("Flow_CZ2")},
                {"Flow_Table_Main_Init_Var", new SubFlowSheet("Flow_Table_Main_Init_Var")
                {
                    Rows = new FlowRows ()
                    {
                        new FlowRow (){ Opcode = OpCode.BinTable },
                    }
                }},
                {"Flow_efuse_ECID_Deid", new SubFlowSheet("Flow_efuse_ECID_Deid")},
                {"Flow_efuse_ECID", new SubFlowSheet("Flow_efuse_ECID")},
            };

            var insSheets = new Dictionary<string, InstanceSheet>()
            {
                {"TestInst_eFuse", new InstanceSheet("TestInst_eFuse") },
            };
            List<RelayItemNew> relays = [new("module")];
            var mainFlowMain = new MainFlowMain(null, subFlowSheets, insSheets, relays);
            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet flowMainSheet = excelPackage.Workbook.Worksheets.Add("Flow_Main");
                flowMainSheet.Cells[2, 1].Value = "Sequence";
                flowMainSheet.Cells[1, 1].Value = "CP1";
                flowMainSheet.Cells[1, 2].Value = "CP2";

                flowMainSheet.Cells[3, 1].Value = "module";
                flowMainSheet.Cells[3, 2].Value = "CP2_D";
                flowMainSheet.Cells[4, 1].Value = "Flow_DC_Conti:G1";
                flowMainSheet.Cells[4, 2].Value = "CP1_B";
                flowMainSheet.Cells[5, 1].Value = "Flow_Scan";
                flowMainSheet.Cells[5, 2].Value = "CP2_D";
                flowMainSheet.Cells[6, 1].Value = "Flow_Mbist";
                flowMainSheet.Cells[6, 2].Value = "CP2_D";
                flowMainSheet.Cells[7, 1].Value = "Flow_DCTEST_GPIO";
                flowMainSheet.Cells[7, 2].Value = "CP2_D";
                flowMainSheet.Cells[8, 1].Value = "Flow_SPI";
                flowMainSheet.Cells[8, 2].Value = "CP2_D";
                flowMainSheet.Cells[9, 1].Value = "Flow_Vddbinning";
                flowMainSheet.Cells[9, 2].Value = "CP2_D";
                flowMainSheet.Cells[10, 1].Value = "Flow_eFuse_ECID_All_Blank_Check";
                flowMainSheet.Cells[10, 2].Value = "CP2_D";
                flowMainSheet.Cells[11, 1].Value = "Flow_eFuse";
                flowMainSheet.Cells[11, 2].Value = "CP2_D";
                flowMainSheet.Cells[12, 1].Value = "Flow_AAA";
                flowMainSheet.Cells[12, 2].Value = "CP2_D";
                flowMainSheet.Cells[13, 1].Value = "Flow_SelSram";
                flowMainSheet.Cells[13, 2].Value = "CP2_D";
                flowMainSheet.Cells[14, 1].Value = "Flow_eFuse_BC_4_EVS";
                flowMainSheet.Cells[14, 2].Value = "CP2_D";
                flowMainSheet.Cells[15, 1].Value = "Flow_EVS";
                flowMainSheet.Cells[15, 2].Value = "CP2_D";
                flowMainSheet.Cells[16, 1].Value = "Flow_eFuse_IDS_Bincut_PreCheck";
                flowMainSheet.Cells[16, 2].Value = "CP2_D";
                flowMainSheet.Cells[17, 1].Value = "Flow_Post_IDS_EQN_Voltage";
                flowMainSheet.Cells[17, 2].Value = "CP2_D";
                //flowMainSheet.Cells[18, 1].Value = "CONCURRENT_AAA";
                //flowMainSheet.Cells[18, 2].Value = "CP2_D";
                flowMainSheet.Cells[18, 1].Value = "Flow_HARDIP_Test_Run_01";
                flowMainSheet.Cells[18, 2].Value = "CP2_D";
                mainFlowMain.MainFlowGen(flowMainSheet);
            }
            TestProgram.Print();

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }

            SelSramPatternSingleton.Initialize();
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void MainFlowGenTest_1()
        {
            string subName = "MainFlowGen_1";
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            TestProgram.Clear();
            LocalSpecs.TarFolder = outputPath;
            BlockStatus.GetAutomationBlockStatus(BlockStatus.Efuse).Down = true;
            LocalSpecs.SetEnableModules([BlockStatus.Efuse, "Defer"]);
            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);
            MainFlowSheet? mainFlowSheet = TestPlanStatic.MainFlowSheet ?? TestPlanStatic.MainFlowSheet;
            foreach (MainFlowBase row in mainFlowSheet!.Rows)
            {
                foreach (FlowSequenceNew sequencesNew in row.SequencesNew)
                {
                    sequencesNew.Group = "Group1";
                }
            }
            MainFlowBase mainFlowBase = mainFlowSheet.Rows.First();
            FlowSequenceNew flowSequenceNew = new FlowSequenceNew();
            flowSequenceNew.OptionDict["Defer"] = "Flow_efuse_BankRead";
            flowSequenceNew.Module = "Defer";
            flowSequenceNew.Enable = true;
            flowSequenceNew.SheetName = "Flow_efuse_BankRead";
            flowSequenceNew.Source = "IGXL";
            flowSequenceNew.Comment = "A,B";
            mainFlowBase.SequencesNew.Add(flowSequenceNew);

            // Arrange
            var subFlowSheets = new Dictionary<string, SubFlowSheet>
            {
                {"Flow_efuse_BankRead", new SubFlowSheet("Flow_efuse_BankRead")
                {
                    Rows = new FlowRows ()
                    {
                        new FlowRow (){ Opcode = OpCode.BinTable },
                        new FlowRow (){ Opcode = OpCode.BinTable, Parameter = "Bin_HIP" },
                        new FlowRow (){ Opcode = OpCode.BinTable, Parameter = "Bin_HIP" ,Enable = "Enable" }
                    }
                }
                },
                {"Flow_efuse_EVS_All_Blank_Check", new SubFlowSheet("Flow_efuse_EVS_All_Blank_Check")},
                {"module", new SubFlowSheet("module")}
            };
            var insSheets = new Dictionary<string, InstanceSheet>();
            List<RelayItemNew> relays = [new("module")];
            var mainFlowMain = new MainFlowMain(null, subFlowSheets, insSheets, relays);
            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet flowMainSheet = excelPackage.Workbook.Worksheets.Add("Flow_Main");
                flowMainSheet.Cells[1, 1].Value = "XXX";
                mainFlowMain.MainFlowGen(flowMainSheet);
            }
            TestProgram.Print();

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
