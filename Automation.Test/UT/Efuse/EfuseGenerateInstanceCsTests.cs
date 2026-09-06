using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Business;
using Automation.GenerateIgxl.EFuse.Enums;
using Automation.Static;
using Automation.Test.Static;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Harvest;

namespace Automation.Test.UT.Efuse
{
    [TestClass]
    public class EfuseGenerateInstanceCsTests : FunctionTestBase
    {
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestProgram.VbtFunctionLib.AddVbtFunctionRange(TestSuiteInitialize.Functions);
        }

        [TestMethod]
        public void GenerateAllBlankCheckInstanceTestCs()
        {
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            EfuseFinalInstanceRow ecidItem = new EfuseFinalInstanceRow();
            InstanceRow row = efuseGenerateInstance.GenerateAllBlankCheckInstance("Name", ecidItem);
            Assert.AreEqual("IgxlWrapper.CoreTestLibrary.EFuse.EFuseMain.Bank_Read", row.VbtName);
        }

        [TestMethod]
        public void GenerateReadInstanceRow()
        {
            string subName = "GenerateReadInstanceRow";
            string outputPath = Path.Combine(OutputPath, "Efuse", subName);
            string expectPath = Path.Combine(ExpectPath, "Efuse", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            EfuseFinalInstanceRow bankItem = new EfuseFinalInstanceRow();
            InstanceRow result = efuseGenerateInstance.GenerateReadInstanceRow(bankItem);

            // Assert
            string json = JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [DataTestMethod]
        [DataRow(BankType.UdrE, "CMP_E")]
        [DataRow(BankType.UdrP, "CMP_P")]
        [DataRow(BankType.UdrP0, "CMP_P0")]
        [DataRow(BankType.UdrP1, "CMP_P1")]
        public void GenerateReadInstanceRow_Ver2BankMapping_SetsCorrectCmpName(string inputBank, string expectedCmpName)
        {
            // 1. Arrange: Ensure TestName contains "Ver2" and the BankName matches the test case
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "MyTest_" + EFuseConst.Ver2,
                BankName = inputBank,
                EfusePatternRow = new EfusePatternRow
                {
                    PatternType = new EfusePatternType(),
                    PayloadList = [] // Empty list to avoid additional logic
                }
            };

            // 2. Act
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            InstanceRow result = efuseGenerateInstance.GenerateReadInstanceRow(bankItem);

            // 3. Assert: Verify SetParamValue was called with the specific "CMP_" string
            Assert.AreEqual(expectedCmpName, result.Args[4]);
        }

        [TestMethod]
        public void GenerateAllBlankCheckInstance_DefaultBankItem_SetsBaseParams()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var ecidItem = new EfuseFinalInstanceRow { BankName = "SomeBank", TestName = "SomeTest" };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateAllBlankCheckInstance("InstanceName_HV", ecidItem);

            // Assert
            Assert.AreEqual(".NET", row.VbtType);
            Assert.AreEqual("Efuse_X_X_X", row.DcCategory);
            Assert.AreEqual("Max", row.DcSelector);
            Assert.AreEqual("Typ", row.AcSelector);
            Assert.AreEqual("Levels_Efuse", row.PinLevels);
            Assert.AreEqual("SomeBank", row.GetArgument("bankName"));
            Assert.AreEqual("FALSE", row.GetArgument("earlyfuse"));
            Assert.AreEqual("FALSE", row.GetArgument("isEcid"));
            Assert.AreEqual("FALSE", row.GetArgument("blankCheckCurrentStage"));
            Assert.AreEqual("FALSE", row.GetArgument("rvOnly"));
            Assert.AreEqual("TRUE", row.GetArgument("printDecode"));
            Assert.AreEqual("TRUE", row.GetArgument("PrintDspWave"));
        }

        [TestMethod]
        public void GenerateAllBlankCheckInstance_EcidEarlyBlankCheckWithReadWritePinAndRvSuffix_SetsConditionalParams()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var ecidItem = new EfuseFinalInstanceRow
            {
                BankName = BankType.Ecid,
                TestName = "Test_" + EFuseConst.BlankCheck + "_RV",
                ExtraType = EfuseExtraType.Early,
                EfusePatternRow = new EfusePatternRow { ReadWritePin = "PIN_A" }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateAllBlankCheckInstance("InstanceName_LV", ecidItem);

            // Assert
            Assert.AreEqual("Min", row.DcSelector);
            Assert.AreEqual(BankType.Ecid, row.GetArgument("bankName"));
            Assert.AreEqual("TRUE", row.GetArgument("isEcid"));
            Assert.AreEqual("TRUE", row.GetArgument("earlyfuse"));
            Assert.AreEqual("TRUE", row.GetArgument("blankCheckCurrentStage"));
            Assert.AreEqual("TRUE", row.GetArgument("rvOnly"));
            Assert.AreEqual("PIN_A", row.GetArgument("pinName"));
        }

        [TestMethod]
        public void GenerateReadInstanceRow_NonVer2WithReadWritePinAndFlags_SetsExpectedParams()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_" + EFuseConst.BlankCheck,
                BankName = BankType.Ecid,
                ExtraType = EfuseExtraType.Early,
                EfusePatternRow = new EfusePatternRow
                {
                    ReadWritePin = "PIN_R",
                    PatternType = new EfusePatternType { IsDvrv = true }
                }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateReadInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("PIN_R", row.GetArgument("pinName"));
            Assert.AreEqual(BankType.Ecid, row.GetArgument("bankName"));
            Assert.AreEqual("TRUE", row.GetArgument("isEcid"));
            Assert.AreEqual("TRUE", row.GetArgument("earlyFuse"));
            Assert.AreEqual("TRUE", row.GetArgument("blankCheckCurrentStage"));
            Assert.AreEqual("TRUE", row.GetArgument("printdecode"));
            Assert.AreEqual("TRUE", row.GetArgument("PrintDspWave"));
            Assert.AreEqual("TRUE", row.GetArgument("RvOnly"));
        }

        [TestMethod]
        public void GenerateReadInstanceRow_JtagReadModeWithCsLibraryFolderSet_UsesNonJtagBranch()
        {
            // Arrange - the isJtag branch also requires LocalSpecs.CsLibraryFolder to be empty; this
            // fixture (TestSuiteInitialize) always sets it, so JtagRead mode still resolves "Bank_read".
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_Jtag",
                BankName = "BankX",
                EfusePatternRow = new EfusePatternRow { PatternType = new EfusePatternType { TestMode = EfuseTestMode.JtagRead } }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateReadInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("IgxlWrapper.CoreTestLibrary.EFuse.EFuseMain.Bank_Read", row.VbtName);
            Assert.AreEqual("TRUE", row.GetArgument("printdecode"));
        }

        [TestMethod]
        public void GenerateInstanceRowByInstanceSheet_SetsExpectedFields()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var instanceSheetRow = new EfuseFinalInstanceRow
            {
                EfuseInstanceRow = new BinCutInstanceRow("EfuseSheet")
                {
                    Instance = "InstanceA",
                    DCcategory = "Efuse_X_X_X extra",
                    FunctionName = "Bank_read",
                    RowNum = 7
                },
                EfusePatternRow = new EfusePatternRow { PayloadList = ["Pat1"] }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateInstanceRowByInstanceSheet(instanceSheetRow);

            // Assert
            Assert.AreEqual("InstanceA", row.TestName);
            Assert.AreEqual(".NET", row.VbtType);
            Assert.AreEqual("Efuse_X_X_X", row.DcCategory);
            Assert.AreEqual(7, row.RowNum);
            Assert.AreEqual("EfuseSheet", row.SheetName);
            Assert.AreEqual("EfuseSheet,Row7;DC Category:Efuse_X_X_X", row.ColumnA);
        }

        [TestMethod]
        public void GenerateBinCheckInstance_AddsRowPerHardBin()
        {
            // Arrange
            List<HarvestingTruthTableSheet> original = SwapHarvestingTruthTableSheets(
            [
                new HarvestingTruthTableSheet("Harvesting_Truth_Table_1")
                {
                    Rows = [new HarvestingTruthTableRow { OutputBinTables = new Dictionary<string, string> { { "Bin07", "1" } } }]
                }
            ]);
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var instanceSheet = new InstanceSheet("Inst_Efuse");
            try
            {
                // Act
                List<string> result = efuseGenerateInstance.GenerateBinCheckInstance(instanceSheet);

                // Assert
                Assert.AreEqual(1, result.Count);
                Assert.AreEqual("EFUSE_BinCheck_Bin07", result[0]);
                InstanceRow row = instanceSheet.Rows.Single();
                Assert.AreEqual("EFUSE_BinCheck_Bin07", row.TestName);
                Assert.AreEqual(".NET", row.VbtType);
                Assert.AreEqual("FuseCheck instance for Bin07 from Harvest truth tables", row.ColumnA);
                Assert.AreEqual("07", row.GetArgument("hardBinNum"));
            }
            finally
            {
                SwapHarvestingTruthTableSheets(original);
            }
        }

        [TestMethod]
        public void GenerateInitInstanceRow_NetFunctionFound_SetsExpectedParams()
        {
            // Arrange
            LocalSpecs.Options.MergeInitPatterns = false;
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_Init_HV",
                InitPatName = "Pattern_DSSC_1",
                EfusePatternRow = new EfusePatternRow { InitList = ["InitPat0"] }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateInitInstanceRow(bankItem);

            // Assert
            Assert.AreEqual(".NET", row.VbtType);
            Assert.AreEqual("Efuse_X_X_X", row.DcCategory);
            Assert.AreEqual("Max", row.DcSelector);
            Assert.AreEqual("InitPat0", row.GetArgument("Patterns"));
            Assert.AreEqual("0", row.GetArgument("ResultMode"));
            Assert.AreEqual("1", row.GetArgument("RelayMode"));
        }

        [TestMethod]
        public void GenerateCrcInstanceRow_NetFunctionFound_SetsExpectedParams()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_CRC",
                EfusePatternRow = new EfusePatternRow { PayloadList = ["CrcPat"] }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateCrcInstanceRow(bankItem);

            // Assert
            Assert.AreEqual(".NET", row.VbtType);
            Assert.AreEqual("Efuse_X_X_X", row.DcCategory);
            Assert.AreEqual("Typ", row.DcSelector);
            Assert.AreEqual("CrcPat", row.GetArgument("Patterns"));
            Assert.AreEqual("0", row.GetArgument("ResultMode"));
            Assert.AreEqual("1", row.GetArgument("RelayMode"));
        }

        [TestMethod]
        public void GenerateCompareWrInstanceRow_DefaultBank_SetsFalseFlags()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var bankItem = new EfuseFinalInstanceRow { TestName = "Bank_CompareWr", BankName = "BankY" };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateCompareWrInstanceRow(bankItem);

            // Assert
            Assert.AreEqual(".NET", row.VbtType);
            Assert.AreEqual("BankY", row.GetArgument("bankName"));
            Assert.AreEqual("FALSE", row.GetArgument("earlyfuse"));
            Assert.AreEqual("FALSE", row.GetArgument("isEcid"));
            Assert.AreEqual("FALSE", row.GetArgument("blankCheckCurrentStage"));
            Assert.AreEqual("FALSE", row.GetArgument("rvOnly"));
            Assert.AreEqual(string.Empty, row.GetArgument("initPatternSet"));
        }

        [DataTestMethod]
        [DataRow(EfuseExtraType.Early, DisplayName = "Early")]
        [DataRow(EfuseExtraType.Deid, DisplayName = "Deid")]
        public void GenerateCompareWrInstanceRow_EarlyOrDeidWithCfgBankAndRvSuffix_SetsTrueFlags(EfuseExtraType extraType)
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_CompareWr_" + EFuseConst.BlankCheck + "_RV",
                BankName = BankType.Cfg,
                ExtraType = extraType
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateCompareWrInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("Config", row.GetArgument("bankName"));
            Assert.AreEqual("TRUE", row.GetArgument("earlyfuse"));
            Assert.AreEqual("TRUE", row.GetArgument("blankCheckCurrentStage"));
            Assert.AreEqual("TRUE", row.GetArgument("rvOnly"));
        }

        [TestMethod]
        public void GenerateCompareWrInstanceRow_EcidBank_SetsIsEcidTrue()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var bankItem = new EfuseFinalInstanceRow { TestName = "Bank_CompareWr", BankName = BankType.Ecid };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateCompareWrInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("TRUE", row.GetArgument("isEcid"));
        }

        [TestMethod]
        public void GenerateWriteInstanceRow_PrgPayloadWithPowerPinAndCfgBank_SetsPowerAndFlags()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_Write_HV",
                BankName = BankType.Cfg,
                ExtraType = EfuseExtraType.Early,
                EfusePatternRow = new EfusePatternRow
                {
                    PayloadList = ["PRG_Write"],
                    ReadWritePin = "PIN_W",
                    PatternType = new EfusePatternType { IsDvrv = true }
                }
            };
            var powerPin = new List<string> { "VDD1", "VDD2" };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateWriteInstanceRow(bankItem, powerPin);

            // Assert
            Assert.AreEqual("Max", row.DcSelector);
            Assert.AreEqual("Config", row.GetArgument("bankName"));
            Assert.AreEqual("TRUE", row.GetArgument("earlyFuse"));
            Assert.AreEqual("FALSE", row.GetArgument("isEcid"));
            Assert.AreEqual("VDD1,VDD2", row.GetArgument("powerPin"));
            Assert.AreEqual("1.8", row.GetArgument("powerPinVoltage"));
            Assert.AreEqual("TRUE", row.GetArgument("rvOnly"));
            Assert.AreEqual("PIN_W", row.GetArgument("pinName"));
        }

        [TestMethod]
        public void GenerateWriteInstanceRow_NonPrgPayloadWithEcidBank_SkipsPowerParams()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_Write_LV",
                BankName = BankType.Ecid,
                EfusePatternRow = new EfusePatternRow { PayloadList = ["WritePat"] }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateWriteInstanceRow(bankItem, []);

            // Assert
            Assert.AreEqual("Min", row.DcSelector);
            Assert.AreEqual("TRUE", row.GetArgument("isEcid"));
            Assert.AreEqual("FALSE", row.GetArgument("earlyFuse"));
            Assert.AreEqual("FALSE", row.GetArgument("rvOnly"));
            Assert.AreEqual(string.Empty, row.GetArgument("powerPin"));
        }

        [TestMethod]
        public void GenerateVer1InstanceRow_NetFunctionFound_SetsPatternsParam()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_Ver1_LV",
                EfusePatternRow = new EfusePatternRow { PayloadList = ["Ver1Pat"] }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateVer1InstanceRow(bankItem);

            // Assert
            Assert.AreEqual("Min", row.DcSelector);
            Assert.AreEqual("Efuse_X_X_X", row.DcCategory);
            Assert.AreEqual("Ver1Pat", row.GetArgument("patterns"));
        }

        [TestMethod]
        public void GenerateUdrUfpInstanceRow_NetFunctionFound_SetsPatternsParam()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_UdrUfp_HV",
                EfusePatternRow = new EfusePatternRow { PayloadList = ["UfpPat"] }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateUdrUfpInstanceRow(bankItem, []);

            // Assert
            Assert.AreEqual("Max", row.DcSelector);
            Assert.AreEqual("UfpPat", row.GetArgument("patterns"));
        }

        [TestMethod]
        public void GenerateUdrUfrInstanceRow_NetFunctionFound_SetsPatternsParam()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_UdrUfr_NV",
                EfusePatternRow = new EfusePatternRow { PayloadList = ["UfrPat"] }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateUdrUfrInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("Typ", row.DcSelector);
            Assert.AreEqual("UfrPat", row.GetArgument("patterns"));
        }

        [DataTestMethod]
        [DataRow("HV", "Max", DisplayName = "HV")]
        [DataRow("LV", "Min", DisplayName = "LV")]
        [DataRow("NV", "Typ", DisplayName = "NV")]
        [DataRow("Unknown", "Typ", DisplayName = "Unknown")]
        public void GetDcSelector_MapsSelectorNameToDcSelector(string selectorName, string expected)
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();

            // Act
            string result = efuseGenerateInstance.GetDcSelector(selectorName);

            // Assert
            Assert.AreEqual(expected, result);
        }

        private static List<HarvestingTruthTableSheet> SwapHarvestingTruthTableSheets(List<HarvestingTruthTableSheet> newValue)
        {
            List<HarvestingTruthTableSheet> original = TestPlanStatic._harvestingTruthTable;
            TestPlanStatic._harvestingTruthTable = newValue;
            return original;
        }
    }
}
