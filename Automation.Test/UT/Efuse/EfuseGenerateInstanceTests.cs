using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Business;
using Automation.GenerateIgxl.EFuse.Enums;
using Automation.Static;
using Automation.Test.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.VbtLib;

namespace Automation.Test.UT.Efuse
{
    [TestClass]
    public class EfuseGenerateInstanceTests : FunctionTestBase
    {
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestProgram.VbtFunctionLib.AddVbtFunctionRange(TestSuiteInitialize.Functions);
        }

        [DataTestMethod]
        [DataRow("InitTest", false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, true, null, DisplayName = "01_InitTest_Skipped")]
        [DataRow("InitTest", false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "Functional_T_updated", DisplayName = "02_InitTest_Generate")]
        [DataRow("", false, EfuseTestMode.Crc, DvRvType.None, BankType.Unknow, false, "Functional_T_updated", DisplayName = "03_CRC_Instance")]
        [DataRow("write", true, EfuseTestMode.None, DvRvType.Dv, BankType.Unknow, false, "auto_ConfigWrite_CFG_DV", DisplayName = "04_DV_Instance")]
        [DataRow("", true, EfuseTestMode.FlatCheck, DvRvType.None, BankType.Unknow, false, "EFUSE_Flat_Pattern_Check", DisplayName = "05_FlatCheck_Instance")]
        [DataRow(EFuseConst.CompareWr, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "Bank_CompareWRData", DisplayName = "06_CompareWR_Instance")]
        [DataRow(EFuseConst.SyntaxCheck, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "Bank_SyntaxCheck", DisplayName = "07_SyntaxCheck_Instance")]
        [DataRow(EFuseConst.ShowEcid, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "auto_ShowECIDData", DisplayName = "08_ShowECID_Instance")]
        [DataRow(EFuseConst.EcidSorting, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "BinSorting_Compare_FT_ECID_S", DisplayName = "09_ECIDSorting_Instance")]
        [DataRow(EFuseConst.BkmFt, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "IgxlWrapper.CoreTestLibrary.EFuse.EFuseMain.Set_BKM", DisplayName = "10_BKMFT_Instance")]
        [DataRow(EFuseConst.IedaBkmSet, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "WriteIEDARegistry", DisplayName = "11_IEDA_Instance")]
        [DataRow(EFuseConst.Write, false, EfuseTestMode.Usi, DvRvType.None, BankType.Unknow, false, "Bank_Write", DisplayName = "12_Write_Instance")]
        [DataRow(EFuseConst.SwitchFlag, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "SwitchFlag", DisplayName = "13_SwitchFlag_Instance")]
        [DataRow("", false, EfuseTestMode.Ver1, DvRvType.None, BankType.Unknow, false, "auto_Function_Test", DisplayName = "14_Ver1_Instance")]
        [DataRow("", false, EfuseTestMode.Ufp, DvRvType.None, BankType.UdrE, false, "auto_UDR_UFP", DisplayName = "15_UdrUfp_Instance")]
        [DataRow("", false, EfuseTestMode.Ufr, DvRvType.None, BankType.UdrE, false, "auto_UDR_UFR", DisplayName = "16_UdrUfr_Instance")]
        [DataRow("NV", false, EfuseTestMode.Ufr, DvRvType.None, BankType.UdrE, true, "auto_UDR_UFR", DisplayName = "17_UdrUfr_Instance")]
        [DataRow("LV", false, EfuseTestMode.Ufr, DvRvType.None, BankType.UdrE, true, "auto_UDR_UFR", DisplayName = "18_UdrUfr_Instance")]
        [DataRow("HV", false, EfuseTestMode.Ufr, DvRvType.None, BankType.UdrE, true, "auto_UDR_UFR", DisplayName = "19_UdrUfr_Instance")]
        [DataRow(EFuseConst.BlankCheck, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "Bank_Read", DisplayName = "20_BlankCheck_Instance")]
        [DataRow("BlankCheck_VER2", true, EfuseTestMode.JtagRead, DvRvType.None, BankType.UdrE, false, "Bank_Read", DisplayName = "21_BlankCheck_Instance")]
        [DataRow("BlankCheck_VER2", false, EfuseTestMode.None, DvRvType.None, BankType.UdrP, false, "Bank_Read", DisplayName = "22_BlankCheck_Instance")]
        [DataRow("BlankCheck_VER2", false, EfuseTestMode.None, DvRvType.None, BankType.UdrP0, false, "Bank_Read", DisplayName = "23_BlankCheck_Instance")]
        [DataRow("BlankCheck_VER2", false, EfuseTestMode.None, DvRvType.None, BankType.UdrP1, false, "Bank_Read", DisplayName = "24_BlankCheck_Instance")]
        [DataRow("BlankCheck_VER2", true, EfuseTestMode.JtagRead, DvRvType.None, BankType.Ecid, false, "Bank_Read", DisplayName = "25_BlankCheck_Instance")]
        public void EfuseGenerateInstanceTest(string testName, bool isDvrv, EfuseTestMode efuseTestMode, DvRvType dvRvType, string bank, bool mergeInitPatterns, string expectedVbtName)
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstance();

            LocalSpecs.Options.MergeInitPatterns = mergeInitPatterns;

            var row = new EfuseFinalInstanceRow
            {
                TestName = testName,
                BankName = bank,
                EfusePatternRow = new EfusePatternRow
                {
                    PatternType = new EfusePatternType
                    {
                        IsDvrv = isDvrv,
                        TestMode = efuseTestMode,
                        DvrvType = dvRvType
                    },
                    InitList = ["_DSSC_"],
                    ReadWritePin = "ReadWritePin"
                },
                InitPatName = "_DSSC_"
            };

            // Act
            List<InstanceRow> result = efuseGenerateInstance.GenerateInstanceRows([row]);

            // Assert
            if (string.IsNullOrEmpty(expectedVbtName))
            {
                Assert.AreEqual(0, result.Count, "Row should have been skipped");
            }
            else
            {
                Assert.AreEqual(expectedVbtName, result[0].VbtName);
            }
        }

        [DataTestMethod]
        [DataRow("InitTest", false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, true, null, DisplayName = "01_InitTest_Skipped")]
        [DataRow("InitTest", false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "IgxlWrapper.CoreTestLibrary.FunctionalTestMain.FuncTestMain", DisplayName = "02_InitTest_Generate")]
        [DataRow("", false, EfuseTestMode.Crc, DvRvType.None, BankType.Unknow, false, "IgxlWrapper.CoreTestLibrary.FunctionalTestMain.FuncTestMain", DisplayName = "03_CRC_Instance")]
        [DataRow("write", true, EfuseTestMode.None, DvRvType.Dv, BankType.Unknow, false, "auto_ConfigWrite_CFG_DV", DisplayName = "04_DV_Instance")]
        [DataRow("", true, EfuseTestMode.FlatCheck, DvRvType.None, BankType.Unknow, false, "EFUSE_Flat_Pattern_Check", DisplayName = "05_FlatCheck_Instance")]
        [DataRow(EFuseConst.CompareWr, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "IgxlWrapper.CoreTestLibrary.EFuse.EFuseMain.Bank_Read", DisplayName = "06_CompareWR_Instance")]
        [DataRow(EFuseConst.SyntaxCheck, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "Bank_SyntaxCheck", DisplayName = "07_SyntaxCheck_Instance")]
        [DataRow(EFuseConst.ShowEcid, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "auto_ShowECIDData", DisplayName = "08_ShowECID_Instance")]
        [DataRow(EFuseConst.EcidSorting, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "BinSorting_Compare_FT_ECID_S", DisplayName = "09_ECIDSorting_Instance")]
        [DataRow(EFuseConst.BkmFt, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "IgxlWrapper.CoreTestLibrary.EFuse.EFuseMain.Set_BKM", DisplayName = "10_BKMFT_Instance")]
        [DataRow(EFuseConst.IedaBkmSet, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "WriteIEDARegistry", DisplayName = "11_IEDA_Instance")]
        [DataRow(EFuseConst.Write, false, EfuseTestMode.Usi, DvRvType.None, BankType.Unknow, false, "IgxlWrapper.CoreTestLibrary.EFuse.EFuseMain.Bank_Write", DisplayName = "12_Write_Instance")]
        [DataRow(EFuseConst.SwitchFlag, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "SwitchFlag", DisplayName = "13_SwitchFlag_Instance")]
        [DataRow("", false, EfuseTestMode.Ver1, DvRvType.None, BankType.Unknow, false, "IgxlWrapper.CoreTestLibrary.FunctionalTestMain.FuncTestMain", DisplayName = "14_Ver1_Instance")]
        [DataRow("", false, EfuseTestMode.Ufp, DvRvType.None, BankType.UdrE, false, "IgxlWrapper.CoreTestLibrary.FunctionalTestMain.FuncTestMain", DisplayName = "15_UdrUfp_Instance")]
        [DataRow("", false, EfuseTestMode.Ufr, DvRvType.None, BankType.UdrE, false, "IgxlWrapper.CoreTestLibrary.FunctionalTestMain.FuncTestMain", DisplayName = "16_UdrUfr_Instance")]
        [DataRow(EFuseConst.BlankCheck, false, EfuseTestMode.None, DvRvType.None, BankType.Unknow, false, "IgxlWrapper.CoreTestLibrary.EFuse.EFuseMain.Bank_Read", DisplayName = "17_BlankCheck_Instance")]
        public void EfuseGenerateInstanceCsTest(string testName, bool isDvrv, EfuseTestMode efuseTestMode, DvRvType dvRvType, string bank, bool mergeInitPatterns, string expectedVbtName)
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstanceCs();

            LocalSpecs.Options.MergeInitPatterns = mergeInitPatterns;

            var row = new EfuseFinalInstanceRow
            {
                TestName = testName,
                BankName = bank,
                EfusePatternRow = new EfusePatternRow
                {
                    PatternType = new EfusePatternType
                    {
                        IsDvrv = isDvrv,
                        TestMode = efuseTestMode,
                        DvrvType = dvRvType
                    }
                }
            };

            // Act
            List<InstanceRow> result = efuseGenerateInstance.GenerateInstanceRows([row]);

            // Assert
            if (string.IsNullOrEmpty(expectedVbtName))
            {
                Assert.AreEqual(0, result.Count, "Row should have been skipped");
            }
            else
            {
                Assert.AreEqual(expectedVbtName, result[0].VbtName);
            }
        }

        [TestMethod]
        public void AddCsharpInstanceItemTest()
        {
            InstanceRow row = new EfuseGenerateInstance().AddCsharpInstanceItem("testName", "vbtName");
            Assert.AreEqual("vbtName", row.VbtName);

            InstanceRow row1 = new EfuseGenerateInstance().AddCsharpInstanceItem("testName", "Set_BKM");
            Assert.AreEqual("IgxlWrapper.CoreTestLibrary.EFuse.EFuseMain.Set_BKM", row1.VbtName);

            InstanceRow row2 = new EfuseGenerateInstance().AddCsharpInstanceItem("testName", EFuseConst.PseudoFuseReadItem);
            Assert.AreEqual("IgxlWrapper.CoreTestLibrary.EFuse.EFuseMain.PseudoFuse_ReadFromFile", row2.VbtName);
        }

        [TestMethod]
        public void AddInstanceItem()
        {
            InstanceRow row = new EfuseGenerateInstance().AddInstanceItem("testName", "vbtName");
            Assert.AreEqual("vbtName", row.VbtName);

            InstanceRow row1 = new EfuseGenerateInstance().AddInstanceItem("testName", "BKM_Update");
            Assert.AreEqual("BKM_Update", row1.VbtName);

            InstanceRow row2 = new EfuseGenerateInstance().AddInstanceItem("testName", EFuseConst.PseudoFuseReadItem);
            Assert.AreEqual("PseudoFuse_ReadFromFile", row2.VbtName);
        }

        [TestMethod]
        public void GenerateAllBlankCheckInstanceTest()
        {
            var efuseGenerateInstance = new EfuseGenerateInstance();
            EfuseFinalInstanceRow ecidItem = new EfuseFinalInstanceRow();
            InstanceRow row = efuseGenerateInstance.GenerateAllBlankCheckInstance("Name", ecidItem);
            Assert.AreEqual("Bank_Read", row.VbtName);
        }

        [TestMethod]
        public void GenerateAllBlankCheckInstance_DefaultBankItem_SetsBaseParams()
        {
            // Arrange - force resolution of the VBT "Bank_Read" (matching this code's param names)
            // instead of the .NET override that also exists in the test fixture library.
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var ecidItem = new EfuseFinalInstanceRow { BankName = "SomeBank", TestName = "SomeTest" };
            List<Function> removed = RemoveNetFunction("Bank_Read");
            try
            {
                // Act
                InstanceRow row = efuseGenerateInstance.GenerateAllBlankCheckInstance("InstanceName_HV", ecidItem);

                // Assert
                Assert.AreEqual("Bank_Read", row.VbtName);
                Assert.AreEqual("Efuse_X_X_X", row.DcCategory);
                Assert.AreEqual("Max", row.DcSelector);
                Assert.AreEqual("Typ", row.AcSelector);
                Assert.AreEqual("Levels_Efuse", row.PinLevels);
                Assert.AreEqual("SomeBank", row.GetArgument("bank"));
                Assert.AreEqual(string.Empty, row.GetArgument("ecid"));
                Assert.AreEqual(string.Empty, row.GetArgument("earlyfuse"));
                Assert.AreEqual(string.Empty, row.GetArgument("blankCheck"));
                Assert.AreEqual(string.Empty, row.GetArgument("PinRead"));
                Assert.AreEqual("TRUE", row.GetArgument("printdecode"));
                Assert.AreEqual("TRUE", row.GetArgument("PrintDspWave"));
                Assert.AreEqual("TRUE", row.GetArgument("blankCheckAll"));
            }
            finally
            {
                TestProgram.VbtFunctionLib.VbtLib.AddRange(removed);
            }
        }

        [TestMethod]
        public void GenerateAllBlankCheckInstance_EcidEarlyBlankCheckWithReadWritePin_SetsConditionalParams()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var ecidItem = new EfuseFinalInstanceRow
            {
                BankName = BankType.Ecid,
                TestName = "Test_" + EFuseConst.BlankCheck,
                ExtraType = EfuseExtraType.Early,
                EfusePatternRow = new EfusePatternRow { ReadWritePin = "PIN_A" }
            };
            List<Function> removed = RemoveNetFunction("Bank_Read");
            try
            {
                // Act
                InstanceRow row = efuseGenerateInstance.GenerateAllBlankCheckInstance("InstanceName_LV", ecidItem);

                // Assert
                Assert.AreEqual("Min", row.DcSelector);
                Assert.AreEqual(BankType.Ecid, row.GetArgument("bank"));
                Assert.AreEqual("TRUE", row.GetArgument("ecid"));
                Assert.AreEqual("TRUE", row.GetArgument("earlyfuse"));
                Assert.AreEqual("TRUE", row.GetArgument("blankCheck"));
                Assert.AreEqual("PIN_A", row.GetArgument("PinRead"));
            }
            finally
            {
                TestProgram.VbtFunctionLib.VbtLib.AddRange(removed);
            }
        }

        [TestMethod]
        public void GenerateFlatCheckInstanceRow_SetsExpectedParams()
        {
            // Arrange
            LocalSpecs.Options.MergeInitPatterns = false;
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_FlatCheck_LV",
                BankName = "BankX",
                EfusePatternRow = new EfusePatternRow { PayloadList = ["FlatPat"] }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateFlatCheckInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("EFUSE_Flat_Pattern_Check", row.VbtName);
            Assert.AreEqual("Efuse_X_X_X", row.DcCategory);
            Assert.AreEqual("Min", row.DcSelector);
            Assert.AreEqual("Typ", row.AcSelector);
            Assert.AreEqual("Levels_Efuse", row.PinLevels);
            Assert.AreEqual("FlatPat", row.GetArgument("Flat_Pattern"));
            Assert.AreEqual("BankX", row.GetArgument("bank"));
        }

        [TestMethod]
        public void GenerateDvInstanceRow_WriteBranch_NoPrg_SetsWriteParams()
        {
            // Arrange
            LocalSpecs.Options.MergeInitPatterns = false;
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_Write_LV",
                EfusePatternRow = new EfusePatternRow { PayloadList = ["WritePat"] }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateDvInstanceRow(bankItem, []);

            // Assert
            Assert.AreEqual("auto_ConfigWrite_CFG_DV", row.VbtName);
            Assert.AreEqual("Min", row.DcSelector);
            Assert.AreEqual("WritePat", row.GetArgument("CFG_DV_pat"));
            Assert.AreEqual(string.Empty, row.GetArgument("PwrPin"));
            Assert.AreEqual(string.Empty, row.GetArgument("vpwr"));
        }

        [TestMethod]
        public void GenerateDvInstanceRow_WriteBranch_WithPrg_SetsPowerParams()
        {
            // Arrange
            LocalSpecs.Options.MergeInitPatterns = false;
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_Write_HV",
                EfusePatternRow = new EfusePatternRow { PayloadList = ["PRG_Write"] }
            };
            var powerPin = new List<string> { "VDD1", "VDD2" };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateDvInstanceRow(bankItem, powerPin);

            // Assert
            Assert.AreEqual("Max", row.DcSelector);
            Assert.AreEqual("VDD1,VDD2", row.GetArgument("PwrPin"));
            Assert.AreEqual("1.8", row.GetArgument("vpwr"));
        }

        [TestMethod]
        public void GenerateDvInstanceRow_ReadBranch_SetsPatternsParam()
        {
            // Arrange
            LocalSpecs.Options.MergeInitPatterns = false;
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_Read_Typ",
                EfusePatternRow = new EfusePatternRow { PayloadList = ["ReadPat"] }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateDvInstanceRow(bankItem, []);

            // Assert
            Assert.AreEqual("Functional_T_updated", row.VbtName);
            Assert.AreEqual("ReadPat", row.GetArgument("Patterns"));
        }

        [TestMethod]
        public void GenerateInitInstanceRow_SetsExpectedParams()
        {
            // Arrange
            LocalSpecs.Options.MergeInitPatterns = false;
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_Init_HV",
                InitPatName = "Pattern_DSSC_1",
                EfusePatternRow = new EfusePatternRow { InitList = ["InitPat0"] }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateInitInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("Functional_T_updated", row.VbtName);
            Assert.AreEqual("Efuse_X_X_X", row.DcCategory);
            Assert.AreEqual("Max", row.DcSelector);
            Assert.AreEqual("InitPat0", row.GetArgument("Patterns"));
            Assert.AreEqual("Test_AutoSwitch:JTAG_TDI", row.GetArgument("DigSource"));
            Assert.AreEqual("0", row.GetArgument("ResultMode"));
            Assert.AreEqual("1", row.GetArgument("RelayMode"));
        }

        [TestMethod]
        public void GenerateCrcInstanceRow_SetsExpectedParams()
        {
            // Arrange
            LocalSpecs.Options.MergeInitPatterns = false;
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_CRC",
                InitPatName = "Pattern_DSSC_1",
                EfusePatternRow = new EfusePatternRow { PayloadList = ["CrcPat"] }
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateCrcInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("Functional_T_updated", row.VbtName);
            Assert.AreEqual("Efuse_X_X_X", row.DcCategory);
            Assert.AreEqual("Typ", row.DcSelector);
            Assert.AreEqual("Typ", row.AcSelector);
            Assert.AreEqual("Levels_Efuse", row.PinLevels);
            Assert.AreEqual("CrcPat", row.GetArgument("Patterns"));
            Assert.AreEqual("0", row.GetArgument("ResultMode"));
            Assert.AreEqual("1", row.GetArgument("RelayMode"));
            Assert.AreEqual("Test_AutoSwitch:JTAG_TDI", row.GetArgument("DigSource"));
        }

        [TestMethod]
        public void GenerateCompareWrInstanceRow_Default_SetsBankOnly()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow { TestName = "Bank_CompareWr", BankName = "BankY" };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateCompareWrInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("Bank_CompareWRData", row.VbtName);
            Assert.AreEqual("BankY", row.GetArgument("bank"));
            Assert.AreEqual(string.Empty, row.GetArgument("earlyfuse"));
            Assert.AreEqual(string.Empty, row.GetArgument("RvOnly"));
        }

        [DataTestMethod]
        [DataRow(EfuseExtraType.Early, DisplayName = "Early")]
        [DataRow(EfuseExtraType.Deid, DisplayName = "Deid")]
        public void GenerateCompareWrInstanceRow_EarlyOrDeidExtraTypeWithRvSuffix_SetsConditionalParams(EfuseExtraType extraType)
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow
            {
                TestName = "Bank_CompareWr_RV",
                BankName = "BankY",
                ExtraType = extraType
            };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateCompareWrInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("TRUE", row.GetArgument("earlyfuse"));
            Assert.AreEqual("TRUE", row.GetArgument("RvOnly"));
        }

        [TestMethod]
        public void GenerateSyntaxInstanceRow_NonVer2Default_SetsBankAndCheckAll()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow { TestName = "Bank_SyntaxCheck", BankName = "BankZ" };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateSyntaxInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("Bank_SyntaxCheck", row.VbtName);
            Assert.AreEqual("BankZ", row.GetArgument("bank"));
            Assert.AreEqual("TRUE", row.GetArgument("compareWR"));
            Assert.AreEqual("TRUE", row.GetArgument("checkAll"));
            Assert.AreEqual(string.Empty, row.GetArgument("earlyfuse"));
            Assert.AreEqual(string.Empty, row.GetArgument("RvOnly"));
            Assert.AreEqual(string.Empty, row.GetArgument("ecid"));
        }

        [DataTestMethod]
        [DataRow(BankType.UdrE, "CMP_E", DisplayName = "UdrE")]
        [DataRow(BankType.UdrP, "CMP_P", DisplayName = "UdrP")]
        [DataRow(BankType.UdrP0, "CMP_P0", DisplayName = "UdrP0")]
        [DataRow(BankType.UdrP1, "CMP_P1", DisplayName = "UdrP1")]
        public void GenerateSyntaxInstanceRow_Ver2BankMapping_SetsExpectedCmpBankAndRvOnly(string inputBank, string expectedCmpBank)
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow { TestName = "Bank_SyntaxCheck_" + EFuseConst.Ver2, BankName = inputBank };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateSyntaxInstanceRow(bankItem);

            // Assert
            Assert.AreEqual(expectedCmpBank, row.GetArgument("bank"));
            Assert.AreEqual(string.Empty, row.GetArgument("compareWR"));
            Assert.AreEqual("TRUE", row.GetArgument("RvOnly"));
        }

        [DataTestMethod]
        [DataRow(EfuseExtraType.Early, DisplayName = "Early")]
        [DataRow(EfuseExtraType.Deid, DisplayName = "Deid")]
        public void GenerateSyntaxInstanceRow_EarlyOrDeidExtraType_SetsEarlyFuseNotCheckAll(EfuseExtraType extraType)
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow { TestName = "Bank_SyntaxCheck", BankName = "BankZ", ExtraType = extraType };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateSyntaxInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("TRUE", row.GetArgument("earlyfuse"));
            Assert.AreEqual(string.Empty, row.GetArgument("checkAll"));
        }

        [TestMethod]
        public void GenerateSyntaxInstanceRow_EcidBank_SetsEcidTrue()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow { TestName = "Bank_SyntaxCheck", BankName = BankType.Ecid };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateSyntaxInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("TRUE", row.GetArgument("ecid"));
        }

        [TestMethod]
        public void GenerateSyntaxInstanceRow_RvSuffixNonUdrBank_SetsRvOnly()
        {
            // Arrange
            var efuseGenerateInstance = new EfuseGenerateInstance();
            var bankItem = new EfuseFinalInstanceRow { TestName = "Bank_SyntaxCheck_RV", BankName = "BankZ" };

            // Act
            InstanceRow row = efuseGenerateInstance.GenerateSyntaxInstanceRow(bankItem);

            // Assert
            Assert.AreEqual("TRUE", row.GetArgument("RvOnly"));
        }

        [TestMethod]
        public void GenerateBkmFtInstanceRow_NonNetSetBkm_UsesGetFusedBkmData()
        {
            // Arrange - force resolution of the VBT "GetFusedBKMData" path by removing the
            // .NET "Set_BKM" override that this fixture library also defines.
            var efuseGenerateInstance = new EfuseGenerateInstance();
            List<Function> removed = RemoveNetFunction("Set_BKM");
            try
            {
                LocalSpecs.HasBkmProcess = true;
                var bankItem = new EfuseFinalInstanceRow { BankName = "FuseTypeX" };

                // Act
                InstanceRow row = efuseGenerateInstance.GenerateBkmFtInstanceRow(bankItem);

                // Assert
                Assert.AreEqual("VBT", row.VbtType);
                Assert.AreEqual("GetFusedBKMData", row.VbtName);
                Assert.AreEqual("FuseTypeX", row.GetArgument("FuseType"));
                Assert.AreEqual("bkm_process", row.GetArgument("cateName"));
            }
            finally
            {
                TestProgram.VbtFunctionLib.VbtLib.AddRange(removed);
                LocalSpecs.HasBkmProcess = false;
            }
        }

        private static List<Function> RemoveNetFunction(string functionName)
        {
            List<Function> removed = [.. TestProgram.VbtFunctionLib.VbtLib.Where(f => f.FunctionName.EqualsIgnoreCase(functionName) && f.Type == ".NET")];
            foreach (Function function in removed)
            {
                TestProgram.VbtFunctionLib.VbtLib.Remove(function);
            }
            return removed;
        }
    }
}
