using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.Scan.Harvest;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Static;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Concurrent;
using TestPlanLib.Harvest;
using TestPlanLib.Singleton;

namespace Automation.Test.UT.Scan
{
    [TestClass]
    public class ScanNonBinCutInstanceMainTests : FunctionTestBase
    {
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void WorkFlow()
        {
            AssertOnlyWindowsOS("MissingPinReport output differs on macOS (Extra harvest pins content mismatch)");

            string subName = "ScanNonBinCutInstanceMain_WorkFlows";
            string outputPath = Path.Combine(OutputPath, "Scan", subName);
            string expectPath = Path.Combine(ExpectPath, "Scan", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);
            BinNumberSingletonLegacy.Initialize(LocalSpecs.ConfigFiles.BinNumberConfig, LocalSpecs.CurrentProject);

            PinMapSheet pinMapSheet = new PinMapSheet("");
            var pinGroup = new PinGroup("PatternPinGroup", "I/O");
            pinGroup.AddPin(new Pin("TX_P", "I/O"));
            pinGroup.AddPin(new Pin("TX_N", "I/O"));
            pinMapSheet.AddGroup(pinGroup);
            Pin pin = new Pin("AON_SLEEP1_RESETN", PinMapConst.TypeIo);
            pinMapSheet.AddPin(pin);
            TestProgram.IgxlWorkBk.PinMapPair = new KeyValuePair<string, PinMapSheet>(FolderStructure.DirPinMap, pinMapSheet);
            DigitalFlagsSheet digitalFlags = new DigitalFlagsSheet("Digital_Flags");
            digitalFlags.Rows.Add(new DigitalFlagsSheetRow()
            {
                Ip = "",
                FlagName = "F_A1",
                Statement = "F_A2||F_A3",
                Print = "",
                ScanFlag = "",
                MbistFlag = "",
                Comments = "",
            });
            TestPlanStatic.DigitalFlagsSheet = digitalFlags;
            ConcurrentFlowSheet concurrentFlowSheet = new ConcurrentFlowSheet("Concurrent_Flow");
            concurrentFlowSheet.Rows.Add(new ConcurrentFlowSheetRow()
            {
                SequenceName = "EVS_FLOW",
                Subflows = ["Flow_Harv_Out_TD_min"]
            });
            TestPlanStatic.ConcurrentFlowSheet = concurrentFlowSheet;

            List<BinCutInstanceSheet> scanInstanceSheets = TestPlanStatic.ScanInstanceSheets;
            var row = new BinCutInstanceRow("HTOL_Sheet")
            {
                RowNum = 1,
                FlowNameOri = "EVS_FLOW",
                Instance = "Flow_EVS_MAIN_FLOW_01",
                EnableAndDevice = "Enable_CPU",
                SubFlow = "Sub1",
                EnableFlow = "EnableFlow1",
                JobTestStage = "CP1",
                SiteVar = "Site1",
                FailFlag = "X",
                BinOutStage = "CP1",
                DCcategory = "EVS_X_X_X_Eqn",
                TimeSet = "TS_EVS_MAIN_FLOW:Param",
                ShiftSpeed = "Normal",
                PatternPinGroup = "name:PatternPinGroup(A);EnableCoreHarvest:TRUE;",
                PinGroupBinoutFlag = "F_HV,F_LV,F_NV,F_HV||F_LV,F_HV&&F_LV",
                PatternList =
                    [
                        "cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t10pd",
                        "cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t10",
                        "PP_BRNA0_L_INLP_SC_CFXX_SAA_COM_AUT_ALLFRV_DM_XORDSSC",
                        "PP_BRNA0_XORDSSCAA_X_X",
                        "dp_brna0_c_fulp_an_aa22_frq_jtg_dyc_allfrv_si_ciopll_t2p1r+dp_brna0_c_fulp_an_aa22_frq_jtg_dyc_allfrv_si_ciopll_t2p1r",
                        "cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t6pd",
                        "cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t6pd",
                    ],
                InitList = ["INIT_1_2_3_4_5_6_7_8_MA009", "INIT_2"],
                PayloadList =
                [
                    "cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t6pd",
                    "cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t6pd",
                ],
                Type = BincutInstanceType.Pattern,
                Char = "Char"
            };
            scanInstanceSheets.First().Rows.Insert(0, row);
            ScanConfig config = SettingStatic.ScanConfig;
            var testClass = new ScanNonBinCutInstanceMain(config);
            testClass.WorkFlow();

            TestProgram.Print();

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
            TestPlanStatic.DigitalFlagsSheet = null;
            TestPlanStatic.ConcurrentFlowSheet = null;
        }

        [TestMethod]
        public void GenAllFailFlagToMainFlow_Should_Collect_GroupFlags()
        {
            var sheet = new BinTableSheet("Test");
            sheet.AddRow(new BinTableRow { ItemList = "F_CPU_Group", Items = ["T"] });
            sheet.AddRow(new BinTableRow { ItemList = "F_CPU_Group", Items = ["T"] });
            sheet.AddRow(new BinTableRow { ItemList = "OtherFlag", Items = ["T"] });

            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());
            List<string> result = testClass.GenAllFailFlagToMainFlow(sheet);

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("F_CPU_Group", result[0]);
        }

        [TestMethod]
        public void GenPatSets_Should_Create_PatSet_Per_Row()
        {
            var binCutRows = new List<BinCutFinalInstanceRow>
            {
                new()
                {
                    Domain = "CPU",
                    PatSetName = "PS1",
                    PatternList = ["PAT1", "PAT2"],
                    BinCutInstanceRow = new BinCutInstanceRow { SheetName = "S1", RowNum = 1 }
                }
            };

            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());
            List<PatSet> result = testClass.GenPatSets(binCutRows);

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("PS1", result[0].PatSetName);
            Assert.AreEqual("CPU", result[0].Domain);
            Assert.AreEqual(2, result[0].PatSetRows.Count);
        }

        [TestMethod]
        public void AddCommandAndFlagInPatSet_Should_Mark_Unused_As_Backup()
        {
            // Arrange
            var patSets = new List<PatSet>
            {
                new() { PatSetName="Unused", PatSetRows = { new PatSetRow { File = "cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t10" } } },
                new() { PatSetName="Used", PatSetRows = { new PatSetRow { File = "cz_brna0_c_fulp_an_aa00_dll_jtg_vix_allfrv_si_cpllds_t10pd" } } }
            };

            var binCutFinalInstanceRows = new List<BinCutFinalInstanceRow>
            {
                new()
            };

            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());

            // Act
            List<PatSet> result = testClass.AddCommandAndFlagInPatSet(patSets, binCutFinalInstanceRows);

            // Assert
            Assert.IsTrue(result[0].IsBackup);
            Assert.AreEqual(", dont_useInFlow, no_pattern", result[0].PatSetRows.First().Comment);
            Assert.IsTrue(result[1].IsBackup);
            Assert.AreEqual(", dont_useInFlow, no_pattern", result[1].PatSetRows.First().Comment);
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void GenBinTableRows_Should_Create_Logical_And_HLV_Bins()
        {
            string subName = "GenBinTableRows";
            string outputPath = Path.Combine(OutputPath, "Scan", subName);
            string expectPath = Path.Combine(ExpectPath, "Scan", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);
            BinNumberSingletonLegacy.Initialize(LocalSpecs.ConfigFiles.BinNumberConfig, LocalSpecs.CurrentProject);

            var rows = new List<BinCutFinalInstanceRow>
            {
                new()
                {
                    Domain = "CPU",
                    Block = "CORE",
                    PatternList = ["PAT_1"],
                    PayloadList = ["P1","P2"],
                    BinCutInstanceRow =new BinCutInstanceRow
                    {
                        FailFlag= "F_CPU_A && F_CPU_B || F_SOC_C, F_SOC_D || F_SOC_E",
                        BinOutStage = "CP1"
                    }
                },
                new()
                {
                    Domain = "CPU",
                    Block = "CORE",
                    PatternList = ["PAT_1"],
                    PayloadList = ["P1","P2"],
                    BinCutInstanceRow =new BinCutInstanceRow
                    {
                        FailFlag= "F_CPU_A && F_CPU_B || F_SOC_C, F_SOC_D || F_SOC_E",
                        BinOutStage = "CP2"
                    }
                }
            };

            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());

            // Act
            List<BinTableRow> binTableRows = testClass.GetBinTableRows(rows);

            IEnumerable<FlowRow> flowRows = testClass.GenFlowBinTable(rows);

            // Assert
            var subFlowSheet = new SubFlowSheet("Flow_Scan");
            subFlowSheet.Rows.AddRange(flowRows);
            subFlowSheet.Write(Path.Combine(outputPath, subFlowSheet.Name + ".txt"));
            var binTableSheet = new BinTableSheet("BinTable_Scan");
            binTableSheet.AddRows(binTableRows);
            binTableSheet.Write(Path.Combine(outputPath, binTableSheet.Name + ".txt"));

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void SetPatSetSheetTest()
        {
            // Arrange
            var patSets = new List<PatSet>
            {
                new() { Domain = "Soc" },
                new() { Domain = "Cpu" },
                new() { Domain = "Gfx" },
                new() { Domain = "Other" }
            };

            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());
            LocalSpecs.TarFolder = OutputPath;

            // Act
            testClass.SetPatSetSheet(patSets);

            // Assert
            var sheets = TestProgram.IgxlWorkBk.PatSetSheets.Values.ToList();
            Assert.IsTrue(sheets.Exists(s => s.Name == "PatSets_Non_Bincut"));
        }

        [DataTestMethod]
        [DataRow("A_B_C", "full", "A_B_C", DisplayName = "FullRuleReturnsWholeName")]
        [DataRow("A_B_C", "0,2", "A_C", DisplayName = "IndexListSelectsWords")]
        [DataRow("A_B_C", "", "", DisplayName = "EmptyRuleReturnsEmpty")]
        [DataRow("", "0", "", DisplayName = "EmptyNameReturnsEmpty")]
        public void GetSubName_SelectsWordsByRuleOrReturnsWholeName(string name, string rule, string expected)
        {
            // Act
            string result = ScanNonBinCutInstanceMain.GetSubName(name, rule);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetPayloadType_NullPayloadTypeTable_ReturnsEmpty()
        {
            // Arrange
            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());

            // Act
            string result = testClass.GetPayloadType("cz_some_pattern");

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetPayloadType_WhitespacePattern_ReturnsEmpty()
        {
            // Arrange
            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());

            // Act
            string result = testClass.GetPayloadType("   ");

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void OtherFlags_BinoutFlagsNotInPatternPinGroup_ReturnsExtras()
        {
            // Arrange
            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());
            var row = new BinCutInstanceRow
            {
                PatternPinGroup = "name:GroupA(F_HV);name:GroupB(F_LV)",
                PinGroupBinoutFlag = "F_HV,F_LV,F_NV"
            };

            // Act
            List<string> result = testClass.OtherFlags(row);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "F_NV" }, result);
        }

        [DataTestMethod]
        [DataRow("SocDomain", "TestInst_SocScan")]
        [DataRow("CpuDomain", "TestInst_CpuScan")]
        [DataRow("GfxDomain", "TestInst_GfxScan")]
        [DataRow("OtherDomain", "TestInst_Non_Bincut")]
        public void GetInstanceSheetNameByDomain_MapsDomainToSheetName(string domain, string expected)
        {
            // Arrange
            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());

            // Act
            string result = testClass.GetInstanceSheetNameByDomain(domain);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("HV", "Max")]
        [DataRow("LV", "Min")]
        [DataRow("NV", "Typ")]
        [DataRow("Unknown", "Typ")]
        public void GetDcSelector_MapsSelectorNameToDcSelector(string selectorName, string expected)
        {
            // Arrange
            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());

            // Act
            string result = testClass.GetDcSelector(selectorName);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("Scan_X_X_X_HV", "Scan_X_X_X")]
        [DataRow("Scan_X_X_X_LV", "Scan_X_X_X")]
        [DataRow("Scan_X_X_X", "Scan_X_X_X")]
        public void GetDcCategory_StripsVoltageSuffix(string category, string expected)
        {
            // Arrange
            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());

            // Act
            string result = testClass.GetDcCategory(category);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("Scan_EVS_X_X", "Levels_EVS_Scan", DisplayName = "EvsCategory")]
        [DataRow("Scan_BIST_X_X", "Levels_Mbist", DisplayName = "BistCategory")]
        [DataRow("Scan_BIRA_X_X", "Levels_Mbist", DisplayName = "BiraCategory")]
        [DataRow("Scan_X_X_X", "Levels_Scan", DisplayName = "DefaultCategory")]
        public void GenerateLevel_NoUserDefinedLevel_MapsCategoryToLevels(string dcCategory, string expected)
        {
            // Arrange
            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());

            // Act
            string result = testClass.GenerateLevel(dcCategory, "", "");

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GenerateLevel_UserDefinedLevelProvided_ReturnsItUnchanged()
        {
            // Arrange
            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());

            // Act
            string result = testClass.GenerateLevel("Scan_BIST_X_X", "", "Levels_Custom");

            // Assert
            Assert.AreEqual("Levels_Custom", result);
        }

        [TestMethod]
        public void RemoveDuplicateBinTableRows_ReturnsCopiesOfAllInputRows()
        {
            // Arrange - the method compares by reference against fresh Copy() instances, so
            // duplicate name/item-list pairs are never actually removed from the returned list.
            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());
            var rows = new List<BinTableRow>
            {
                new() { Name = "Bin_A", ItemList = "F_1,F_2" },
                new() { Name = "Bin_A", ItemList = "F_1,F_2" },
                new() { Name = "Bin_B", ItemList = "F_3" }
            };

            // Act
            List<BinTableRow> result = testClass.RemoveDuplicateBinTableRows(rows);

            // Assert
            Assert.AreEqual(2, result.Count(x => x.Name == "Bin_A"));
            Assert.AreEqual(1, result.Count(x => x.Name == "Bin_B"));
            Assert.IsFalse(result.Any(rows.Contains), "Result rows should be copies, not the original instances");
        }

        [DataTestMethod]
        [DataRow("Instance_ScanA", true, DisplayName = "InstancePrefixKept")]
        [DataRow("BinCut_Instance_A", false, DisplayName = "BinCutInstancePrefixExcluded")]
        [DataRow("Instance_BinCut_A", false, DisplayName = "InstanceBinCutPrefixExcluded")]
        [DataRow("Instance_Post_BinCut_A", false, DisplayName = "InstancePostBinCutExcluded")]
        [DataRow("Instance_EvsScan", false, DisplayName = "EvsExcluded")]
        [DataRow("Instance_HtolScan", false, DisplayName = "HtolExcluded")]
        [DataRow("Instance_CpmScan", false, DisplayName = "CpmExcluded")]
        [DataRow("Other_Sheet", false, DisplayName = "NonInstancePrefixExcluded")]
        public void GetBinCutInstanceSheets_FiltersByPrefixAndKeyword(string sheetName, bool expectedIncluded)
        {
            // Act
            List<string> result = ScanNonBinCutInstanceMain.GetBinCutInstanceSheets([sheetName]);

            // Assert
            Assert.AreEqual(expectedIncluded, result.Contains(sheetName));
        }

        [TestMethod]
        public void BinCutInstanceNamingSheet_ReturnsSettingStaticValue()
        {
            // Arrange
            var testClass = new ScanNonBinCutInstanceMain(new ScanConfig());

            // Act
            BinCutInstanceNamingSheet result = testClass.BinCutInstanceNamingSheet();

            // Assert
            Assert.AreSame(SettingStatic.BinCutInstanceNamingSheet, result);
        }
    }
}
