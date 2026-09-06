using System.Collections.Generic;

using Automation.GenerateIgxl.EFuse.Business;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Efuse.Input;

namespace Automation.Test.UT.Efuse
{
    [TestClass]
    public class EFusePreProcessTests : FunctionTestBase
    {
        private readonly string _patSetNameWithInfo = "EFuse_PatSetName_Test";
        private readonly string _patSetNameWithNull = "";
        private List<EfusePatternRow> _efusePatternRows = null!;

        [TestInitialize]
        public void Setup()
        {
            _efusePatternRows =
            [
                new()
                {
                    InitList = ["I1_2_3_4_5_6_7_8_9_10_11_12"],
                    PayloadList = ["1_2_3_4_5_6_7_DAA_SNS_10_11_CPE_FLDSSC1"],
                    SheetName = "Instance_EFuse",
                    RowNum = 1
                }
            ];
        }

        [TestMethod]
        public void GetEfuseInstanceRows_With_PatSetNameOrange()
        {
            // Arrange
            var instanceSheet = new BinCutInstanceSheet("")
            {
                Rows =
                [
                    new()
                    {
                        PatSetNameOrange = _patSetNameWithInfo,
                        InitList = ["I1_2_3_4_5_6_7_8_9_10_11_12"],
                        PayloadList = ["1_2_3_4_5_6_7_DAA_SNS_10_11_CPE_FLDSSC1"],
                        SheetName = "Instance_EFuse",
                        RowNum = 1
                    }
                ]
            };

            var efuseInstanceSheets = new List<BinCutInstanceSheet>
            {
                instanceSheet
            };

            EfuseFinalInstanceRows finalResult = new EFusePreProcess().GetEfuseInstanceRows(efuseInstanceSheets, _efusePatternRows);
            string result = finalResult[0].EfuseInstanceRow.PatSetNameOrange;
            Assert.AreEqual(_patSetNameWithInfo, result);
        }

        [TestMethod]
        public void GetEfuseInstanceRows_With_Null_PatSetNameOrange()
        {
            // Arrange
            var instanceSheet = new BinCutInstanceSheet("")
            {
                Rows =
                [
                    new()
                    {
                        PatSetNameOrange = _patSetNameWithNull,
                        InitList = ["I1_2_3_4_5_6_7_8_9_10_11_12"],
                        PayloadList = ["1_2_3_4_5_6_7_DAA_SNS_10_11_CPE_FLDSSC1"],
                        SheetName = "Instance_EFuse",
                        RowNum = 1
                    }
                ]
            };

            var efuseInstanceSheets = new List<BinCutInstanceSheet>
            {
                instanceSheet
            };

            EfuseFinalInstanceRows finalResult = new EFusePreProcess().GetEfuseInstanceRows(efuseInstanceSheets, _efusePatternRows);
            string result = finalResult[0].EfuseInstanceRow.PatSetNameOrange;
            Assert.AreEqual(_patSetNameWithNull, result);
        }

        [TestMethod]
        public void ExtraNameWithApb_PrependsApbPrefix()
        {
            // Act
            string result = EFusePreProcess.ExtraNameWithApb("Test");

            // Assert
            Assert.AreEqual("APB_Test", result);
        }

        [TestMethod]
        public void GetVoltageList_NullReadItem_ReturnsEmptyList()
        {
            // Arrange
            var target = new EFusePreProcess();

            // Act
            List<string> result = target.GetVoltageList(null, []);

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetVoltageList_CombinesDistinctJobValuesAndBankVoltages()
        {
            // Arrange
            var target = new EFusePreProcess();
            var readItem = new EfuseReadRow
            {
                JobDictionary = new Dictionary<string, string> { { "J1", "HV" }, { "J2", "LV" }, { "J3", "HV/LV" }, { "J4", "" } }
            };

            // Act
            List<string> result = target.GetVoltageList(readItem, ["NV"]);

            // Assert
            CollectionAssert.AreEquivalent(new List<string> { "HV", "LV", "NV" }, result);
        }

        [DataTestMethod]
        [DataRow("MarginRead", "MR1")]
        [DataRow("NormalRead", "MR0")]
        [DataRow("Other", "")]
        public void GetEfuseReadMrName_MapsWriteReadToMrCode(string writeRead, string expected)
        {
            // Arrange
            var target = new EFusePreProcess();
            var readItem = new EfuseReadRow { WriteRead = writeRead };

            // Act
            string result = target.GetEfuseReadMrName(readItem);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetEfuseReadMrName_NullReadItem_ReturnsEmpty()
        {
            // Arrange
            var target = new EFusePreProcess();

            // Act
            string result = target.GetEfuseReadMrName(null);

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void MergeSamePatternType_MultipleRowsSameGroup_MergesListsAndDeduplicatesReadWritePin()
        {
            // Arrange
            var target = new EFusePreProcess();
            var patType = new EfusePatternType { TestMode = EfuseTestMode.Crc, PatJob = "CP1" };
            var rows = new List<EfusePatternRow>
            {
                new() { BankName = "BankA", PatternType = patType, PayloadList = ["P1"], InitList = ["I1"], PatList = ["Pat1"], ReadWritePin = "PIN_A" },
                new() { BankName = "BankA", PatternType = patType, PayloadList = ["P2"], InitList = ["I1"], PatList = ["Pat2"], ReadWritePin = "PIN_A" }
            };

            // Act
            List<EfusePatternRow> result = target.MergeSamePatternType(rows);

            // Assert
            Assert.AreEqual(1, result.Count);
            CollectionAssert.AreEquivalent(new List<string> { "P1", "P2" }, result[0].PayloadList);
            CollectionAssert.AreEqual(new List<string> { "I1" }, result[0].InitList);
            Assert.AreEqual("PIN_A", result[0].ReadWritePin);
        }

        [TestMethod]
        public void MergeSamePatternType_SingleRowInGroup_CopiesListsDirectly()
        {
            // Arrange
            var target = new EFusePreProcess();
            var patType = new EfusePatternType { TestMode = EfuseTestMode.Ver1, PatJob = "CP1" };
            var rows = new List<EfusePatternRow>
            {
                new() { BankName = "BankB", PatternType = patType, PayloadList = ["P1"], InitList = ["I1"], PatList = ["Pat1"], ReadWritePin = "PIN_B" }
            };

            // Act
            List<EfusePatternRow> result = target.MergeSamePatternType(rows);

            // Assert
            Assert.AreEqual(1, result.Count);
            CollectionAssert.AreEqual(new List<string> { "P1" }, result[0].PayloadList);
            Assert.AreEqual("PIN_B", result[0].ReadWritePin);
        }

        [TestMethod]
        public void MergeSamePatternType_UnknownTestMode_IsExcluded()
        {
            // Arrange
            var target = new EFusePreProcess();
            var rows = new List<EfusePatternRow>
            {
                new() { BankName = "BankC", PatternType = new EfusePatternType() }
            };

            // Act
            List<EfusePatternRow> result = target.MergeSamePatternType(rows);

            // Assert
            Assert.AreEqual(0, result.Count);
        }

    }
}
