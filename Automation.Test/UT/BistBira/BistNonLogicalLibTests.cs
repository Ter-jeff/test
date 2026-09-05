using Automation.GenerateIgxl.BistBira.Base;
using Automation.GenerateIgxl.BistBira.NewLogicData;
using Automation.GenerateIgxl.BistBira.NonLogicData;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Singleton;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using ScghLib.Base;
using ScghLib.Enums;
using ScghLib.Reader;

namespace Automation.Test.UT.BistBira
{
    [TestClass]
    public class BistNonLogicalLibTests : FunctionTestBase
    {
        [TestClass]
        public class BypassDcSpecValueTests
        {
            [TestMethod]
            public void GetBypassDcSpecValue_WhenSplitDc_Return_SC()
            {
                string result = BistNonLogicalLib.GetBypassDcSpecValue(true);

                Assert.AreEqual("_SC", result);
            }

            [TestMethod]
            public void GetBypassDcSpecValue_WhenNotSplitDc_ReturnEmpty()
            {
                string result = BistNonLogicalLib.GetBypassDcSpecValue(false);

                Assert.AreEqual(string.Empty, result);
            }
        }

        private BistNonLogicalLib _lib = null!;

        [TestInitialize]
        public void Setup()
        {
            var mbistConfig = new MbistConfig();
            _lib = new BistNonLogicalLib(mbistConfig, MultiTestSettingSheetsSingleton.Instance());
        }

        [TestMethod]
        public void GenInstanceRow_GetColumnA_ReturnsCorrectFormat()
        {
            var genInstanceRow = new BistNonLogicalLib.GenInstanceRow();
            var row = new BistProdFlowRow { SheetName = "Sheet1", RowNum = 5 };

            string result = genInstanceRow.GetColumnA(row);

            Assert.AreEqual("Sheet1,Row5", result);
        }

        [TestMethod]
        public void GenInstanceRow_GetTimeSetName_ReturnsEmpty_WhenPatternListIsNull()
        {
            var genInstanceRow = new BistNonLogicalLib.GenInstanceRow();

            string result = genInstanceRow.GetTimeSetName("PatternX");

            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GenInstanceRow_GetTimeSetVersion_ReturnsMissing_WhenDicIsNull()
        {
            var genInstanceRow = new BistNonLogicalLib.GenInstanceRow();

            string result = genInstanceRow.GetTimeSetVersion("TS1");

            Assert.AreEqual("missing timeset", result);
        }

        [TestMethod]
        public void WorkFlow_AddsInstanceSheetAndRows()
        {
            var bistProdFlowSheet = new BistProdFlowSheet
            {
                MbistSheet = new MbistSheet { SheetName = "TestSheet" },
                Rows =
                [
                    new()
                    {
                        RowNum = 101,
                        Label = "AA",
                        Pattern = "",
                        Note = "",
                        Voltage = "Bincut_X_X_X NV",
                        SheetName = "Test",
                        PassBranch = "B1",
                        FailBranch = "B2",
                        IsDsscRow = true,
                        OriPerformance = "MS001,MS002"
                    },
                    new()
                    {
                        RowNum = 102,
                        Label = "A",
                        Pattern = "Pattern1_2_3_4_BI_6_RPI",
                        Note = "Func:Mbist_RBOX_SEL_DigCap_info",
                        Voltage = "Bincut_X_X_X NV",
                        SheetName = "Test",
                        PassBranch = "B",
                        FailBranch = "B",
                        IsDsscRow = true,
                        OriPerformance = "MS001,MS002"
                    },
                    new()
                    {
                        RowNum = 1,
                        Label = "A",
                        Pattern = "Pattern1_BIR",
                        Note = "",
                        Voltage = "Bincut_X_X_X NV",
                        SheetName = "Test",
                        PassBranch = "B",
                        FailBranch = "B",
                        IsDsscRow = true,
                        OriPerformance = "MS001,MS002"
                    },
                    new()
                    {
                        RowNum = 2,
                        Label = "B",
                        Pattern = "Pattern2",
                        Note = "",
                        Voltage = "Bincut_X_X_X LV",
                        SheetName = "Test",
                        PassBranch = "C",
                        FailBranch = "C",
                        IsDsscRow = true,
                        OriPerformance = "MS001,MS002"
                    },
                    new()
                    {
                        RowNum = 3,
                        Label = "C",
                        Pattern = "Pattern3",
                        Note = "",
                        Voltage = "Bincut_X_X_X HV",
                        SheetName = "Test",
                        PassBranch = "D",
                        FailBranch = "D",
                        IsDsscRow = true,
                        OriPerformance = "MS001,MS002"
                    },

                    new()
                    {
                        RowNum = 4,
                        Label = "D",
                        Pattern = "Pattern4",
                        Note = "",
                        Voltage = "Bincut_X_X_X VMARGIN1",
                        SheetName = "Test"
                    },
                    new()
                    {
                        RowNum = 5,
                        Label = "E",
                        Pattern = "Pattern5",
                        Note = "",
                        Voltage = "Bincut_X_X_X VMARGIN3",
                        SheetName = "Test"
                    },
                    new()
                    {
                        RowNum = 6,
                        Label = "F",
                        Pattern = "Pattern6",
                        Note = "",
                        Voltage = "Bincut_X_X_X MHV",
                        SheetName = "Test"
                    },
                    new()
                    {
                        RowNum = 7,
                        Label = "G",
                        Pattern = "Pattern7",
                        Note = "",
                        Voltage = "Bincut_X_X_X MLV",
                        SheetName = "Test"
                    },
                    new()
                    {
                        RowNum = 8,
                        Label = "H",
                        Pattern = "Pattern8",
                        Note = "",
                        Voltage = "Bincut_X_X_X Vvrs",
                        SheetName = "Test"
                    },
                    new()
                    {
                        RowNum = 9,
                        Label = "I",
                        Pattern = "Pattern9",
                        Note = "",
                        Voltage = "Bincut_X_X_X VDST",
                        SheetName = "Test"
                    },
                    new()
                    {
                        RowNum = 10,
                        Label = "J",
                        Pattern = "Pattern10",
                        Note = "",
                        Voltage = "Bincut_X_X_X VRET",
                        SheetName = "Test"
                    },
                    new()
                    {
                        RowNum = 11,
                        Label = "K",
                        Pattern = "Pattern11",
                        Note = "",
                        Voltage = "Bincut_X_X_X VWNOM",
                        SheetName = "Test"
                    },
                    new()
                    {
                        RowNum = 12,
                        Label = "L",
                        Pattern = "Pattern12",
                        Note = "",
                        Voltage = "Bincut_X_X_X VNOM",
                        SheetName = "Test"
                    },
                    new()
                    {
                        RowNum = 13,
                        Label = "M",
                        Pattern = "Pattern13",
                        Note = "",
                        Voltage = "Bincut_X_X_X VMAX",
                        SheetName = "Test"
                    },
                    new()
                    {
                        RowNum = 14,
                        Label = "N",
                        Pattern = "Pattern14",
                        Note = "",
                        Voltage = "Bincut_X_X_X VMIN",
                        SheetName = "Test"
                    },
                    new()
                    {
                        RowNum = 15,
                        Label = "O",
                        Pattern = "Pattern15",
                        Note = "",
                        Voltage = "Bincut_X_X_X VRS",
                        SheetName = "Test"
                    },
                    new()
                    {
                        RowNum = 16,
                        Label = "O",
                        Pattern = "Pattern15",
                        Note = "",
                        Voltage = "Bincut_X_X_X VRS LV",
                        SheetName = "Test"
                    }
                ]
            };

            var bistIgxlResult = new BistIgxlResult { InstanceSheets = [] };
            var mbistDataStore = new MbistDataStore();

            _lib.WorkFlow(bistProdFlowSheet, bistIgxlResult, mbistDataStore, MbistPatSetType.BurstNo);

            Assert.AreEqual(1, bistIgxlResult.InstanceSheets.Count);
            Assert.AreEqual("TestInst_TestSheet", bistIgxlResult.InstanceSheets[0].Name);
            Assert.AreEqual(18, bistIgxlResult.InstanceSheets[0].Rows.Count);
        }

        [DataTestMethod]
        [DataRow("_ERT_BIR", "ERTBIRA")]
        [DataRow("Some_ERTBIRA_Pattern", "ERTBIRA")]
        [DataRow("_ERT_OTHER", "ERTBIST")]
        [DataRow("Pattern_ERTBIST", "ERTBIST")]
        [DataRow("_NRT_", "NRT")]
        [DataRow("_SRTBIRA_", "SRTBIRA")]
        [DataRow("_WUS_", "WUS")]
        [DataRow("_IRT_", "IRT")]
        [DataRow("UnknownPattern", "")]
        public void CheckRetentionNew_ShouldReturnExpectedValue(string input, string expected)
        {
            // Act
            string result = BistNonLogicalLib.CheckRetentionNew(input);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("NV", "Typ")]
        [DataRow("HV", "Max")]
        [DataRow("LV", "Min")]
        [DataRow("Other", "Typ")]
        public void GetSelector_MapsVoltageToDcSelector(string voltage, string expected)
        {
            // Act
            string result = _lib.GetSelector(voltage);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("HV", "Max")]
        [DataRow("MHV1", "Max")]
        [DataRow("LV", "Min")]
        [DataRow("MLV", "Min")]
        [DataRow("NV", "Typ")]
        [DataRow("Other", "")]
        public void GetDcSelector_MapsVoltageTypeToDcSelector(string voltageType, string expected)
        {
            // Act
            string result = _lib.GetDcSelector(voltageType);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("Bincut_X_X_X NV", "Typ")]
        [DataRow("Bincut_X_X_X HV", "Max")]
        [DataRow("Bincut_X_X_X LV", "Min")]
        public void FindDcSelector_TwoTokenVoltage_DelegatesToGetSelector(string voltage, string expected)
        {
            // Act - the two-token branch resolves via GetSelector without touching _multiTestSettings
            string result = _lib.FindDcSelector(voltage);

            // Assert
            Assert.AreEqual(expected, result);
        }

    }
}
