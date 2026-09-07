using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.PostAction.GenMainFlow;
using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.GenerateIgxl.PostAction.GenMainFlow.Business;
using Automation.Static;
using Automation.Test.Static;

using CommonLib.Extension;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using TestPlanLib.Basic;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class MainFlowMainTests : FunctionTestBase
    {
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestProgram.VbtFunctionLib.AddVbtFunctionRange(TestSuiteInitialize.Functions);
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void GenMainFlow_BySubProgramTest()
        {
            AssertOnlyWindowsOS("expected output files were generated on Windows with backslash paths");
            string subName = "MainFlow";
            string inputPath = Path.Combine(InputPath, "PostAction", subName);
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            Dictionary<string, SubFlowSheet> subFlowSheets = BuildIgxlSubFlowSheets();

            LocalSpecs.TarFolder = outputPath;
            string testPlanDir = Path.Combine(InputPath, "komodo_documents", "A0_V06A", "TestPlan");
            string csvPath = Path.Combine(InputPath, "PostAction", "MainFlow", "Flow_Main.csv");
            string variableInitializePath = Path.Combine(testPlanDir, "Other", "Variable_Initialize.csv");
            string subprogramMappingPath = Path.Combine(InputPath, "PostAction", "MainFlow", "SubprogramMapping.csv");

            InstanceSheet inst1 = new ReadInstanceSheet().GetSheet(Path.Combine(inputPath, "TestInst_Common.txt"));
            InstanceSheet inst2 = new ReadInstanceSheet().GetSheet(Path.Combine(inputPath, "TestInst_UF.txt"));
            InstanceSheet inst3 = new ReadInstanceSheet().GetSheet(Path.Combine(inputPath, "TestInst_eFuse.txt"));
            Dictionary<string, InstanceSheet> insSheets = new Dictionary<string, InstanceSheet>
            {
                { "1", inst1 },
                { "2", inst2 },
                { "3", inst3 }
            };

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet variableInitializeSheet = excelPackage.Workbook.Worksheets.Add("Variable_Initialize");
                _ = variableInitializeSheet.Cells[1, 1].PrintExcelRowByList(variableInitializePath.CsvConvertToLists());
                VariableInitTable? variableInitTable = new VariableInitTableReader().ReadSheet(variableInitializeSheet);

                ExcelWorksheet flowMainSheet = excelPackage.Workbook.Worksheets.Add("Flow_Main");
                _ = flowMainSheet.Cells[1, 1].PrintExcelRowByList(csvPath.CsvConvertToLists());

                ExcelWorksheet subprogramMappingSheet = excelPackage.Workbook.Worksheets.Add("SubprogramMapping");
                _ = subprogramMappingSheet.Cells[1, 1].PrintExcelRowByList(subprogramMappingPath.CsvConvertToLists());
                TestPlanStatic.SubprogramMappingSheet = new SubprogramMappingReader().ReadSheet(subprogramMappingSheet);

                var mainFlowSheetReader = new MainFlowSheetReaderNew();
                MainFlowSheet mainFlowSheet = mainFlowSheetReader.ReadSheet(flowMainSheet);
                MainFlowBase mainFlowBase = mainFlowSheet.Rows.Last();
                FlowSequenceNew flowSequenceNew = new FlowSequenceNew
                {
                    Enable = true,
                    SheetName = "DCTEST_IDS",
                    OriSheetName = "DCTEST_IDS",
                    Source = "IGXL",
                    Module = "",
                    Comment = "A,B"
                };
                mainFlowBase.SequencesNew.Add(flowSequenceNew);
                FlowSequenceNew flowSequenceNew1 = new FlowSequenceNew
                {
                    Enable = true,
                    SheetName = "BINTABLE",
                    Option = "Option",
                    Source = "",
                    Module = "BINTABLE"
                };
                mainFlowBase.SequencesNew.Add(flowSequenceNew1);
                LocalSpecs.SetEnableModules(["CHAR", "IGXL"]);
                LocalSpecs.CustomPath = new List<string> { outputPath };
                FlowSequenceNew flowSequenceNew2 = new FlowSequenceNew
                {
                    Enable = true,
                    SheetName = "CHAR",
                    SubFlowName = "CHAR",
                    Source = "CHAR",
                    Module = "CHAR"
                };
                mainFlowBase.SequencesNew.Add(flowSequenceNew2);
                FlowSequenceNew flowSequenceNew3 = new FlowSequenceNew
                {
                    Enable = true,
                    Source = "CUSTOM",
                    Module = ""
                };
                mainFlowBase.SequencesNew.Add(flowSequenceNew3);

                FlowSequenceNew flowSequenceNew4 = new FlowSequenceNew
                {
                    Enable = true,
                    Source = "",
                    Module = "",
                    SheetName = "HarvestingTruthTable"
                };
                mainFlowBase.SequencesNew.Add(flowSequenceNew4);

                var mainFlowGen = new MainFlowMain(variableInitTable, subFlowSheets, insSheets, null);
                (_, List<MainFlow> mainFlows, List<SubFlowSheet> flows) = mainFlowGen.GenMainFlow(mainFlowSheet);
                foreach (SubFlowSheet mainFlow in mainFlows)
                {
                    string file = Path.Combine(outputPath, mainFlow.Name + ".txt");
                    mainFlow.Write(file);
                }
                foreach (SubFlowSheet flow in flows)
                {
                    string file = Path.Combine(outputPath, flow.Name + ".txt");
                    flow.Write(file);
                }
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void GenMainFlowsNewTest()
        {
            string subName = "MainFlowNew";
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            LocalSpecs.TarFolder = outputPath;
            string csvPath = Path.Combine(InputPath, "PostAction", "MainFlow", "Flow_Main.csv");
            string subprogramMappingPath = Path.Combine(InputPath, "PostAction", "MainFlow", "SubprogramMapping.csv");
            var subFlowSheets = new Dictionary<string, SubFlowSheet>();
            var insSheets = new Dictionary<string, InstanceSheet>();
            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet flowMainSheet = excelPackage.Workbook.Worksheets.Add("Flow_Main");
                _ = flowMainSheet.Cells[1, 1].PrintExcelRowByList(csvPath.CsvConvertToLists());

                ExcelWorksheet subprogramMappingSheet = excelPackage.Workbook.Worksheets.Add("SubprogramMapping");
                _ = subprogramMappingSheet.Cells[1, 1].PrintExcelRowByList(subprogramMappingPath.CsvConvertToLists());
                TestPlanStatic.SubprogramMappingSheet = new SubprogramMappingReader().ReadSheet(subprogramMappingSheet);

                var mainFlowSheetReader = new MainFlowSheetReaderNew();
                MainFlowSheet mainFlowSheet = mainFlowSheetReader.ReadSheet(flowMainSheet);
                var mainFlowMain = new MainFlowMain(null, subFlowSheets, insSheets);
                (List<string> _, List<MainFlow> mainFlows, List<SubFlowSheet> flows) = mainFlowMain.GenMainFlowsNew(ref mainFlowSheet);
                foreach (MainFlow flow in mainFlows)
                {
                    string file = Path.Combine(outputPath, flow.Name + ".txt");
                    flow.Write(file);
                }
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        private static Dictionary<string, SubFlowSheet> BuildIgxlSubFlowSheets()
        {
            var files = new List<(string, string)>
            {
                 (Path.Combine("C:\\", "Module", "Main", "Flow_Table_Main_Init_Flows"),"Initialize"),
                 (Path.Combine("C:\\", "Module", "Main", "Flow_Table_Main_Init_EnableWd"),"Initialize"),
                 (Path.Combine("C:\\", "Module", "DC_Conti", "Flow_DC_Conti"),"DCTEST_Continuity"),
                 (Path.Combine("C:\\", "Common", "Common_Sheets", "Flow_nWire_Default"),"Autogen"),
                 (Path.Combine("C:\\", "Common", "Common_Sheets", "Flow_nWire_Default_Disable"),"Autogen"),
                 (Path.Combine("C:\\", "Common", "Common_Sheets", "Flow_nWire_Default_Enable"),"Autogen"),
                 (Path.Combine("C:\\", "Common", "Common_Sheets", "Flow_nWire_Enable_RT_CLK32768"),"Autogen"),
                 (Path.Combine("C:\\", "Module", "eFuse", "Flow_efuse_BankRead"),"Instance_EFUSE:Flow_efuse_BankRead"),
                 (Path.Combine("C:\\", "Module", "eFuse", "Flow_efuse_Config_Early"),"Instance_EFUSE:Flow_efuse_Config_Early"),
                 (Path.Combine("C:\\", "Module", "eFuse", "Flow_efuse_Monitor_Early"),"Instance_EFUSE:Flow_efuse_Monitor_Early"),
                 (Path.Combine("C:\\", "Module", "eFuse", "Flow_efuse_ECID"),"Instance_EFUSE:Flow_efuse_ECID"),
                 (Path.Combine("C:\\", "Module", "eFuse", "Flow_efuse_Post"),"Instance_EFUSE:Flow_efuse_Post"),
                 (Path.Combine("C:\\", "Module", "eFuse", "Flow_efuse_PreWrite"),"Instance_EFUSE:Flow_efuse_PreWrite"),
                 (Path.Combine("C:\\", "Module", "eFuse", "Flow_efuse_Config"),"Instance_EFUSE:Flow_efuse_Config"),
                 (Path.Combine("C:\\", "Module", "eFuse", "Flow_efuse_MON"),"Instance_EFUSE:Flow_efuse_MON"),
                 (Path.Combine("C:\\", "Module", "eFuse", "Flow_efuse_UDRP"),"Instance_EFUSE:Flow_efuse_UDRP"),
                 (Path.Combine("C:\\", "Module", "eFuse", "Flow_efuse_UDRE"),"Instance_EFUSE:Flow_efuse_UDRE"),
                 (Path.Combine("C:\\", "Module", "eFuse", "Flow_efuse_BinCheck"),"Instance_EFUSE : Flow_efuse_BinCheck"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HardIP_Init_Flag"),"Initialize"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_DCTEST_IDS"),"DCTEST_IDS"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_DCTEST_IDS_CORR"),"DCTEST_IDS_CORR"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_DCTEST_IDS_POST"),"DCTEST_IDS_POST"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_ABSMIN"),"HARDIP_ABSMIN"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_CLKMON"),"HARDIP_CLKMON"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_DAPS"),"HARDIP_DAPS"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_DRAM"),"HARDIP_DRAM"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_DSGRINGS_CZ"),"HARDIP_DSGRINGS_CZ"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_DSGRINGS"),"HARDIP_DSGRINGS"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_DSGVTHSENSER"),"HARDIP_DSGVTHSENSER"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_IVDM"),"HARDIP_IVDM"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_LLPSMN"),"HARDIP_LLPSMN"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_NCKM"),"HARDIP_NCKM"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_NPMS"),"HARDIP_NPMS"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_PWRDET"),"HARDIP_PWRDET"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_LAPLL_CZ"),"HARDIP_LAPLL_CZ"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_LAPLL"),"HARDIP_LAPLL"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_EUSB"),"HARDIP_EUSB"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_ADCLK_CZ"),"HARDIP_ADCLK_CZ"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_ADCLK"),"HARDIP_ADCLK"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_AMPCMN"),"HARDIP_AMPCMN"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_AMP"),"HARDIP_AMP"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_AMP_VT"),"HARDIP_AMP_VT"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_CPLL_CZ"),"HARDIP_CPLL_CZ"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_CPLL"),"HARDIP_CPLL"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_LPDDRPLL"),"HARDIP_LPDDRPLL"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_ACIPHYCMN"),"HARDIP_ACIPHYCMN"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_ACIOPHYCMN"),"HARDIP_ACIOPHYCMN"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_LPDPTXCMN"),"HARDIP_LPDPTXCMN"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_PCIE5CMN"),"HARDIP_PCIE5CMN"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_CIODPPLL"),"HARDIP_CIODPPLL"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_CIOCIOPLL_CZ"),"HARDIP_CIOCIOPLL_CZ"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_CIOCIOPLL"),"HARDIP_CIOCIOPLL"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_LPDPDPPLL_CZ"),"HARDIP_LPDPDPPLL_CZ"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_LPDPDPPLL"),"HARDIP_LPDPDPPLL"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_PCIEPLL_CZ"),"HARDIP_PCIEPLL_CZ"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_PCIEPLL"),"HARDIP_PCIEPLL"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_PCIEREFPLL_CZ"),"HARDIP_PCIEREFPLL_CZ"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_PCIEREFPLL"),"HARDIP_PCIEREFPLL"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_ACIPHY_CZ"),"HARDIP_ACIPHY_CZ"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_ACIPHY"),"HARDIP_ACIPHY"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_ACIOPHY"),"HARDIP_ACIOPHY"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_LPDPTX"),"HARDIP_LPDPTX"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_PCIE5_CZ"),"HARDIP_PCIE5_CZ"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_PCIE5"),"HARDIP_PCIE5"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_ULPPLL"),"HARDIP_ULPPLL"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_FLPPLL"),"HARDIP_FLPPLL"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_SEPF2M"),"HARDIP_SEPF2M"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_SEPVMC"),"HARDIP_SEPVMC"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_PCIEREFGEN_CZ"),"HARDIP_PCIEREFGEN_CZ"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_PCIEREFGEN"),"HARDIP_PCIEREFGEN"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_IO_JTAG"),"HARDIP_IO_JTAG"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_MIDBIASLDO"),"HARDIP_MIDBIASLDO"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_ADCDFT"),"HARDIP_ADCDFT"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_VMADC"),"HARDIP_VMADC"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_MISC"),"HARDIP_MISC"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_FRO"),"HARDIP_FRO"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_PCIEREFBUF"),"HARDIP_PCIEREFBUF"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_HSCSAROSC"),"HARDIP_HSCSAROSC"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_NTSR"),"HARDIP_NTSR"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_NTSE"),"HARDIP_NTSE"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_NTSX"),"HARDIP_NTSX"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_HARDIP_MTRICTS"),"HARDIP_MTRICTS"),
                 (Path.Combine("C:\\", "Module", "Spi", "Flow_Rtos_UART"),"RTOS_PC"),
                 (Path.Combine("C:\\", "Module", "Spi", "Flow_Rtos_IDS"),"RTOS_IDS"),
                 (Path.Combine("C:\\", "Module", "SPI_ROM", "Flow_Table_Write_SPIROM_main"),"Flow_Table_Write_SPIROM_main"),
                 (Path.Combine("C:\\", "Module", "SPI_ROM", "Flow_Table_Write_SPIROM"),"Flow_Table_Write_SPIROM"),
                 (Path.Combine("C:\\", "Module", "HardIP", "Flow_CLOCKCHECK"),"CLOCKCHECK"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_SOC_TD_MS005_EQN_HV"),"Instance_OutsideBincut:FLOW_SOC_TD_MS005_EQN_HV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_GFX_TD_MG001_EQN_LV"),"Instance_OutsideBincut:FLOW_GFX_TD_MG001_EQN_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_GFX_TD_MG00D_EQN_LV"),"Instance_OutsideBincut:FLOW_GFX_TD_MG00D_EQN_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_ECPU_TD_ME001_EQN_LV"),"Instance_OutsideBincut:FLOW_ECPU_TD_ME001_EQN_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_ECPU_TD_ME008_EQN_LV"),"Instance_OutsideBincut:FLOW_ECPU_TD_ME008_EQN_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_PCPU_TD_MP001_EQN_LV"),"Instance_OutsideBincut:FLOW_PCPU_TD_MP001_EQN_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_PCPU_TD_MP00F_EQN_LV"),"Instance_OutsideBincut:FLOW_PCPU_TD_MP00F_EQN_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_SOC_TD_MP001_EQN_LV"),"Instance_OutsideBincut:FLOW_SOC_TD_MP001_EQN_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_SOC_TD_MP00D_EQN_LV"),"Instance_OutsideBincut:FLOW_SOC_TD_MP00D_EQN_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_CPU_BIST_MP001_LV"),"Instance_OutsideBincut:FLOW_CPU_BIST_MP001_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_CPU_BIST_MP000_LV"),"Instance_OutsideBincut:FLOW_CPU_BIST_MP000_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_CPU_SAA_EQN_LV"),"Instance_SA_CPU:FLOW_CPU_SAA_EQN_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_CPU_SAA_EQN_HV"),"Instance_SA_CPU:FLOW_CPU_SAA_EQN_HV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_GFX_SAA_EQN_LV"),"Instance_SA_GFX:FLOW_GFX_SAA_EQN_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_GFX_SAA_EQN_HV"),"Instance_SA_GFX:FLOW_GFX_SAA_EQN_HV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_SOC_SAA_EQN_LV"),"Instance_SA_SOC:FLOW_SOC_SAA_EQN_LV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "FLOW_SOC_SAA_EQN_HV"),"Instance_SA_SOC:FLOW_SOC_SAA_EQN_HV"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_SSN_CoreMask"),""),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_FlagOperation"),"FlagOperation"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_FlagOperation_Post"),"FlagOperation_Post"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_ATE_STR_Summary_SA"),""),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_ATE_STR_Summary_BIST"),""),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_ATE_STR_Summary_TD"),""),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_ATE_STR_Summary"),"ATE_STR_Summary"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_Harv_Dec_CP1_OB"),"HarvestingTruthTable"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_Harv_Dec_CP1_BC"),"HarvestingTruthTable"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_SetBinFuse"),""),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_Harv_Dec_CP2_RD"),"HarvestingTruthTable"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_Harv_Dec_CP2_OB"),"HarvestingTruthTable"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_Harv_Dec_CP2_BC"),"HarvestingTruthTable"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_Harv_Dec_FT1_RD"),"HarvestingTruthTable"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_Harv_Dec_FT1_OB"),"HarvestingTruthTable"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_Harv_Dec_FT2_RD"),"HarvestingTruthTable"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_Harv_Dec_FT2_OB"),"HarvestingTruthTable"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_Harv_Dec_FT3_OB"),"HarvestingTruthTable"),
                 (Path.Combine("C:\\", "Module", "NonBinCut", "Flow_Harvest_eFuse_Read"),""),
                 (Path.Combine("C:\\", "Module", "cpuMbist", "Flow_C_BI_PP_CP1_BURST"),"C_BI_PP_CP1_BURST"),
                 (Path.Combine("C:\\", "Module", "cpuMbist", "Flow_C_BI_PP_CP1"),"C_BI_PP_CP1"),
                 (Path.Combine("C:\\", "Module", "cpuMbist", "Flow_C_BI_PP_CP2_BURST"),"C_BI_PP_CP2_BURST"),
                 (Path.Combine("C:\\", "Module", "cpuMbist", "Flow_C_BI_PP_CP2"),"C_BI_PP_CP2"),
                 (Path.Combine("C:\\", "Module", "cpuMbist", "Flow_C_BI_PP_FT1_BURST"),"C_BI_PP_FT1_BURST"),
                 (Path.Combine("C:\\", "Module", "cpuMbist", "Flow_C_BI_PP_FT1"),"C_BI_PP_FT1"),
                 (Path.Combine("C:\\", "Module", "cpuMbist", "Flow_C_BI_PP_FT2_BURST"),"C_BI_PP_FT2_BURST"),
                 (Path.Combine("C:\\", "Module", "cpuMbist", "Flow_C_BI_PP_FT2"),"C_BI_PP_FT2"),
                 (Path.Combine("C:\\", "Module", "gfxMbist", "Flow_L_BI_PP_CP1_BURST"),"L_BI_PP_CP1_BURST"),
                 (Path.Combine("C:\\", "Module", "gfxMbist", "Flow_L_BI_PP_CP1"),"L_BI_PP_CP1"),
                 (Path.Combine("C:\\", "Module", "gfxMbist", "Flow_L_BI_PP_CP2_BURST"),"L_BI_PP_CP2_BURST"),
                 (Path.Combine("C:\\", "Module", "gfxMbist", "Flow_L_BI_PP_CP2"),"L_BI_PP_CP2"),
                 (Path.Combine("C:\\", "Module", "gfxMbist", "Flow_L_BI_PP_FT1_BURST"),"L_BI_PP_FT1_BURST"),
                 (Path.Combine("C:\\", "Module", "gfxMbist", "Flow_L_BI_PP_FT1"),"L_BI_PP_FT1"),
                 (Path.Combine("C:\\", "Module", "gfxMbist", "Flow_L_BI_PP_FT2_BURST"),"L_BI_PP_FT2_BURST"),
                 (Path.Combine("C:\\", "Module", "gfxMbist", "Flow_L_BI_PP_FT2"),"L_BI_PP_FT2"),
                 (Path.Combine("C:\\", "Module", "socMbist", "Flow_S_BI_PP_CP1_BURST"),"S_BI_PP_CP1_BURST"),
                 (Path.Combine("C:\\", "Module", "socMbist", "Flow_S_BI_PP_CP1"),"S_BI_PP_CP1"),
                 (Path.Combine("C:\\", "Module", "socMbist", "Flow_S_BI_PP_CP2_BURST"),"S_BI_PP_CP2_BURST"),
                 (Path.Combine("C:\\", "Module", "socMbist", "Flow_S_BI_PP_CP2"),"S_BI_PP_CP2"),
                 (Path.Combine("C:\\", "Module", "socMbist", "Flow_S_BI_PP_FT1_BURST"),"S_BI_PP_FT1_BURST"),
                 (Path.Combine("C:\\", "Module", "socMbist", "Flow_S_BI_PP_FT1"),"S_BI_PP_FT1"),
                 (Path.Combine("C:\\", "Module", "socMbist", "Flow_S_BI_PP_FT2_BURST"),"S_BI_PP_FT2_BURST"),
                 (Path.Combine("C:\\", "Module", "socMbist", "Flow_S_BI_PP_FT2"),"S_BI_PP_FT2"),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_ECPU_mbist_w0"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_ECPU_mbist_w1"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_PCPU_PCORES_mbist_w0"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_PCPU_PCORES_mbist_w1"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_PCPU_CPM_mbist_w0"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_PCPU_CPM_mbist_w1"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_PCPU_Scan_EVS00"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_ECPU_Scan_EVS00"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_PCPU_Scan_EVS11"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_ECPU_Scan_EVS11"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_GFX_MBIST_INVERSE"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_GFX_MBIST_TRUE"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_SOC_Scan_EVS0"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_SOC_Scan_EVS1"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_GFX_Scan_EVS0000"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_GFX_Scan_EVS1111"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_SOC_mbist_WD"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_SOC_mbist_WB"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS"),""),
                 (Path.Combine("C:\\", "Module", "Evs", "Flow_EVS_PowerReset"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MS001_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MS005_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MS002_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MS003_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MS004_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MD001_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MD009_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MD002_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MD003_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MD004_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MD005_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MD006_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MD007_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MD008_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MI001_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MI004_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MI002_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MI003_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG001_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG00D_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG00E_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG00F_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG002_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG003_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG004_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG005_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG006_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG007_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG008_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG009_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG00A_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG00B_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MG00C_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_ME001_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_ME007_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_ME002_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_ME003_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_ME004_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_ME005_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_ME006_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP001_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP00D_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP00J_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP002_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP003_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP004_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP005_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP006_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP007_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP008_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP009_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP00A_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP00B_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_MP00C_TD_Mbist_BV"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_VddBinning_HVCC"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_Vddbinning"),"BV"),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_T0TX_PreCall"),""),
                 (Path.Combine("C:\\", "Module", "BinCut", "Flow_PostBincut"),""),
                 (Path.Combine("C:\\", "Common", "Common_Sheets", "Flow_SelSram"),"SLESRM"),
                 (Path.Combine("C:\\", "Module", "Main", "Flow_Table_Main_EndFlow"),"END")
            };

            var subFlowSheets = new Dictionary<string, SubFlowSheet>();
            foreach ((string, string) file in files)
            {
                string name = Path.GetFileNameWithoutExtension(file.Item1);
                var subFlowSheet = new SubFlowSheet(name)
                {
                    SourceInfo = { Name = file.Item2 }
                };
                subFlowSheets[file.Item1] = subFlowSheet;
            }

            return subFlowSheets;
        }

        [DataTestMethod]
        [DataRow("ABC", new[] { "xxABCyy" }, true, DisplayName = "Contains")]
        [DataRow("ABC", new[] { "xxx", "yyy" }, false, DisplayName = "NotContains")]
        public void IsContain_ChecksCaseInsensitiveSubstringAcrossList(string str, string[] list, bool expected)
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);

            // Act
            bool result = mainFlowMain.IsContain(str, [.. list]);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("Flow_HARDIP_ABC_DEF_GHI", "FLOW_HARDIP_ABCDEFGHI", DisplayName = "ThreeSegmentsCollapsed")]
        [DataRow("Flow_HARDIP_ABC", "Flow_HARDIP_ABC", DisplayName = "TooFewSegmentsUnchanged")]
        [DataRow("Flow_HARDIP_ABC_DEF_init", "Flow_HARDIP_ABC_DEF_init", DisplayName = "InitSuffixUnchanged")]
        [DataRow("Flow_HARDIP_ABC_DEF_GHI_CZ2", "Flow_HARDIP_ABC_DEF_GHI_CZ2", DisplayName = "CzSuffixUnchanged")]
        [DataRow("Flow_Basic_Test", "Flow_Basic_Test", DisplayName = "NonHardIpUnchanged")]
        public void ModifyMainFlowParameter_CollapsesThreeSegmentHardIpNamesOnly(string flowSheet, string expected)
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);

            // Act
            string result = mainFlowMain.ModifyMainFlowParameter(flowSheet);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GenMainFlowRow_EmptyEnv_UsesGivenOpcode()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);

            // Act
            FlowRow row = mainFlowMain.GenMainFlowRow("Flow_Basic_Test", "", "MyComment", "Call", "EnableWd");

            // Assert
            Assert.AreEqual("Call", row.Opcode);
            Assert.AreEqual("", row.Env);
            Assert.AreEqual("EnableWd", row.Enable);
            Assert.AreEqual("MyComment", row.Comment);
            Assert.AreEqual("Flow_Basic_Test", row.Parameter);
        }

        [TestMethod]
        public void GenMainFlowRow_NonEmptyEnv_UsesNopOpcodeAndCollapsesParameter()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);

            // Act
            FlowRow row = mainFlowMain.GenMainFlowRow("Flow_HARDIP_ABC_DEF_GHI", "X", "", "Call", "");

            // Assert
            Assert.AreEqual("Nop", row.Opcode);
            Assert.AreEqual("X", row.Env);
            Assert.AreEqual("FLOW_HARDIP_ABCDEFGHI", row.Parameter);
        }

        [TestMethod]
        public void GenNewItemToMainFlow_SetsAllFields()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);

            // Act
            FlowRow row = mainFlowMain.GenNewItemToMainFlow("Flow_HARDIP_ABC_DEF_GHI", "CP1", "Call", "EnableWd", "ColA", "Bin1");

            // Assert
            Assert.AreEqual("Call", row.Opcode);
            Assert.AreEqual("EnableWd", row.Enable);
            Assert.AreEqual("CP1", row.Job);
            Assert.AreEqual("Bin1", row.FailAction);
            Assert.AreEqual("ColA", row.ColumnA);
            Assert.AreEqual("FLOW_HARDIP_ABCDEFGHI", row.Parameter);
        }

        [TestMethod]
        public void GenNewItemToMainFlow_EmptyFlag_SetsEmptyFailAction()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);

            // Act
            FlowRow row = mainFlowMain.GenNewItemToMainFlow("Flow_Basic_Test", "CP1", "Nop", "", "", "");

            // Assert
            Assert.AreEqual(string.Empty, row.FailAction);
        }

        [DataTestMethod]
        [DataRow("Flow_efuse_BankRead", true, DisplayName = "EfuseSubstring")]
        [DataRow("Flow_Something_Fuse", true, DisplayName = "FuseSubstring")]
        [DataRow("Flow_Scan", false, DisplayName = "NoMatch")]
        public void IsEfuseSubFlow_DetectsEfuseOrFuseSubstring(string sheetName, bool expected)
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);

            // Act
            bool result = mainFlowMain.IsEfuseSubFlow(sheetName);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GenerateEnv_EfuseSubFlow_ReturnsX()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);
            var flowSequence = new FlowSequence("Flow_efuse_BankRead");

            // Act
            string result = mainFlowMain.GenerateEnv(flowSequence, []);

            // Assert
            Assert.AreEqual("X", result);
        }

        [TestMethod]
        public void GenerateEnv_GroupSheetNamePresent_ReturnsEmpty()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);
            var flowSequence = new FlowSequence("Group1:Flow_Scan");

            // Act
            string result = mainFlowMain.GenerateEnv(flowSequence, []);

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GenerateEnv_NoMatchingSubFlowInList_ReturnsX()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);
            var flowSequence = new FlowSequence("Flow_Basic_Test");
            var subFlowList = new List<MainFlowData> { new() { FullSheetName = "Flow_Other" } };

            // Act
            string result = mainFlowMain.GenerateEnv(flowSequence, subFlowList);

            // Assert
            Assert.AreEqual("X", result);
        }

        [TestMethod]
        public void GenerateEnv_MatchingSubFlowInList_ReturnsEmpty()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);
            var flowSequence = new FlowSequence("Flow_Basic_Test");
            var subFlowList = new List<MainFlowData> { new() { FullSheetName = "Flow_Basic_Test" } };

            // Act
            string result = mainFlowMain.GenerateEnv(flowSequence, subFlowList);

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GenerateEnv_TdSubFlowWithVddBinningPresent_ReturnsX()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);
            var flowSequence = new FlowSequence("Flow_Scan_Td");
            var subFlowList = new List<MainFlowData>
            {
                new() { FullSheetName = "Flow_Scan_Td" },
                new() { FullSheetName = "Flow_Vddbinning" }
            };

            // Act
            string result = mainFlowMain.GenerateEnv(flowSequence, subFlowList);

            // Assert
            Assert.AreEqual("X", result);
        }

        [TestMethod]
        public void GetBlockDisabledComment_BlockDisabled_ReturnsUnableComment()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);
            Status status = BlockStatus.GetAutomationBlockStatus(BlockStatus.Basic);
            bool originalEnable = status.Enable;
            bool originalDown = status.Down;
            try
            {
                status.Enable = false;

                // Act
                string result = mainFlowMain.GetBlockDisabledComment(BlockStatus.Basic);

                // Assert
                Assert.AreEqual("Unable generated from Current Input files", result);
            }
            finally
            {
                status.Enable = originalEnable;
                status.Down = originalDown;
            }
        }

        [TestMethod]
        public void GetBlockDisabledComment_BlockEnabledButNotDown_ReturnsDisabledByUiComment()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);
            Status status = BlockStatus.GetAutomationBlockStatus(BlockStatus.Basic);
            bool originalEnable = status.Enable;
            bool originalDown = status.Down;
            try
            {
                status.Enable = true;
                status.Down = false;

                // Act
                string result = mainFlowMain.GetBlockDisabledComment(BlockStatus.Basic);

                // Assert
                Assert.AreEqual("Disabled by User Interface", result);
            }
            finally
            {
                status.Enable = originalEnable;
                status.Down = originalDown;
            }
        }

        [TestMethod]
        public void GetBlockDisabledComment_BlockEnabledAndDown_ReturnsEmpty()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);
            Status status = BlockStatus.GetAutomationBlockStatus(BlockStatus.Basic);
            bool originalEnable = status.Enable;
            bool originalDown = status.Down;
            try
            {
                status.Enable = true;
                status.Down = true;

                // Act
                string result = mainFlowMain.GetBlockDisabledComment(BlockStatus.Basic);

                // Assert
                Assert.AreEqual(string.Empty, result);
            }
            finally
            {
                status.Enable = originalEnable;
                status.Down = originalDown;
            }
        }

        [TestMethod]
        public void GenerateComment_HardIpFlowWithBlockDisabled_ReturnsUnableComment()
        {
            // Arrange
            var subFlowSheets = new Dictionary<string, SubFlowSheet> { { "Flow_HARDIP_ABC", new SubFlowSheet("Flow_HARDIP_ABC") } };
            var mainFlowMain = new MainFlowMain(null, subFlowSheets, []);
            Status status = BlockStatus.GetAutomationBlockStatus(BlockStatus.HardIp);
            bool originalEnable = status.Enable;
            try
            {
                status.Enable = false;

                // Act
                string result = mainFlowMain.GenerateComment("Flow_HARDIP_ABC", null);

                // Assert
                Assert.AreEqual("Unable generated from Current Input files", result);
            }
            finally
            {
                status.Enable = originalEnable;
            }
        }

        [TestMethod]
        public void GenerateComment_FlowNotInSubFlowSheets_ReturnsMissingComment()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);

            // Act
            string result = mainFlowMain.GenerateComment("Flow_Unknown_Sheet", null);

            // Assert
            Assert.AreEqual("Missing this Flow during AutoGen", result);
        }

        [TestMethod]
        public void GetHarvDecisionFlow_MatchingJobAndType_ReturnsSheetName()
        {
            // Arrange
            var subFlowSheets = new Dictionary<string, SubFlowSheet>
            {
                { "Flow_Harv_Dec_CP1_OB", new SubFlowSheet("Flow_Harv_Dec_CP1_OB") { SourceInfo = { Name = "HarvestingTruthTable" } } }
            };
            var mainFlowMain = new MainFlowMain(null, subFlowSheets, []);
            mainFlowMain.SetupSheetNameMappingTable();

            // Act
            string result = mainFlowMain.GetHarvDecisionFlow("OB", "CP1");

            // Assert
            Assert.AreEqual("Flow_Harv_Dec_CP1_OB", result);
        }

        [TestMethod]
        public void GetHarvDecisionFlow_NoMappingTable_ReturnsEmpty()
        {
            // Arrange
            var mainFlowMain = new MainFlowMain(null, [], []);

            // Act
            string result = mainFlowMain.GetHarvDecisionFlow("OB", "CP1");

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void AddCondition_SiteVarPrefixed_SetsConditionAndValue()
        {
            // Arrange
            var flowRow = new FlowRow();

            // Act
            MainFlowMain.AddCondition("site-var MyVar 1", ref flowRow);

            // Assert
            Assert.AreEqual("site-var", flowRow.DeviceCondition);
            Assert.AreEqual("MyVar 1", flowRow.DeviceName);
        }

        [TestMethod]
        public void AddCondition_NegatedFlag_SetsFlagFalseCondition()
        {
            // Arrange
            var flowRow = new FlowRow();

            // Act
            MainFlowMain.AddCondition("!F_Something", ref flowRow);

            // Assert
            Assert.AreEqual("Flag-false", flowRow.DeviceCondition);
            Assert.AreEqual("F_Something", flowRow.DeviceName);
        }

        [TestMethod]
        public void AddCondition_PositiveFlag_SetsFlagTrueCondition()
        {
            // Arrange
            var flowRow = new FlowRow();

            // Act
            MainFlowMain.AddCondition("F_Something", ref flowRow);

            // Assert
            Assert.AreEqual("Flag-true", flowRow.DeviceCondition);
            Assert.AreEqual("F_Something", flowRow.DeviceName);
        }

        [TestMethod]
        public void AddCondition_EndsWithFalse_SetsFlagFalseCondition()
        {
            // Arrange
            var flowRow = new FlowRow();

            // Act
            MainFlowMain.AddCondition("F_Something False", ref flowRow);

            // Assert
            Assert.AreEqual("Flag-false", flowRow.DeviceCondition);
        }

        [TestMethod]
        public void SetDeviceCondition_NoSiteFlag_ReturnsSingleRowWithJobAndRowNum()
        {
            // Arrange
            var sequence = new FlowSequenceNew { SiteFlagPerSite = "", RowNum = 5 };
            var row = new FlowRow();

            // Act
            List<FlowRow> result = MainFlowMain.SetDeviceCondition(sequence, row, "CP1");

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreSame(row, result[0]);
            Assert.AreEqual(5, result[0].RowNum);
            Assert.AreEqual("CP1", result[0].Job);
        }

        [TestMethod]
        public void SetDeviceCondition_SimpleSiteFlag_AddsConditionInline()
        {
            // Arrange
            var sequence = new FlowSequenceNew { SiteFlagPerSite = "F_Something", RowNum = 3 };
            var row = new FlowRow();

            // Act
            List<FlowRow> result = MainFlowMain.SetDeviceCondition(sequence, row, "CP2");

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("Flag-true", result[0].DeviceCondition);
            Assert.AreEqual("CP2", result[0].Job);
        }

        [TestMethod]
        public void SetDeviceCondition_CompoundSiteFlag_WrapsWithIfEndIf()
        {
            // Arrange
            var sequence = new FlowSequenceNew { SiteFlagPerSite = "F_A&&F_B", RowNum = 7 };
            var row = new FlowRow();

            // Act
            List<FlowRow> result = MainFlowMain.SetDeviceCondition(sequence, row, "CP1");

            // Assert
            Assert.AreEqual(3, result.Count);
            Assert.AreEqual("If", result[0].Opcode);
            Assert.AreEqual("(F_A&&F_B) || F_Debug_all", result[0].Parameter);
            Assert.AreSame(row, result[1]);
            Assert.AreEqual("EndIf", result[2].Opcode);
            Assert.AreEqual(7, result[0].RowNum);
            Assert.AreEqual("CP1", result[1].Job);
        }
    }
}
