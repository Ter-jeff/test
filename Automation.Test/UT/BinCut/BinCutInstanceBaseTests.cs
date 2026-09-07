using System.IO;

using Automation.GenerateIgxl.BinCut.Base;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinCut.FlowNew;

namespace Automation.Test.UT.BinCut
{
    [TestClass]
    public class BinCutInstanceBaseTests : BinCutTestBase
    {
        [TestMethod]
        public void GetDcSelector_ShouldReturnCorrectSelector()
        {
            Assert.AreEqual("Min", Instance.GetDcSelector("LV"));
            Assert.AreEqual("Max", Instance.GetDcSelector("HV"));
            Assert.AreEqual("Typ", Instance.GetDcSelector("NV"));
        }

        [TestMethod]
        public void GetInstanceName_ShouldContainBinningDomain()
        {
            string name = Instance.GetInstanceName();
            StringAssert.Contains(name, SourceRow.GetBinType());
        }

        [TestMethod]
        public void GenerateFlowRow_ShouldReturnFlowRowWithExpectedValues()
        {
            FlowRow row = Instance.GenerateFlowRow(false, false, false);
            Assert.AreNotEqual(null, row);
            Assert.IsFalse(string.IsNullOrEmpty(row.Parameter));
            Assert.IsTrue(row.Opcode == OpCode.Test || row.Opcode == OpCode.Nop);
        }

        [TestMethod]
        public void GenerateFlowRow_ShouldReturnFlowRowWithExpectedValuesIsCs()
        {
            FlowRow row = Instance.GenerateFlowRow(false, false, true);
            Assert.AreNotEqual(null, row);
            Assert.IsFalse(string.IsNullOrEmpty(row.Parameter));
            Assert.IsTrue(row.Opcode == OpCode.Test || row.Opcode == OpCode.Nop);
        }

        [TestMethod]
        [DataRow(EnumColumnName.TD, "Bin_TD_Pmode_HBV")]
        [DataRow(EnumColumnName.Mbist, "Bin_Mbist_Pmode_HBV")]
        [DataRow(EnumColumnName.FUNC, "Bin_RTOS_Pmode_HBV")]
        [DataRow(EnumColumnName.ELB, "Bin_ELB_Pmode_HBV")]
        [DataRow(EnumColumnName.ILB, "Bin_ILB_Pmode_HBV")]
        public void GetBinTableRow_ShouldReturnFlowRowWithBinTableOpcode(EnumColumnName enumColumnName, string expected)
        {
            var binCutFlowSheetRow = new BinCutFlowSheetRow("sheetName", []);
            var newBinCutFlowSheetRow = new NewBinCutFlowSheetRow("sheetName", "job")
            {
                TableType = EnumBinCutTableType.Hv
            };
            SourceRow = new BinCutSourceItem(binCutFlowSheetRow, newBinCutFlowSheetRow, enumColumnName, "TEMP SENSOR MONITOR")
            {
                PerformanceMode = "Pmode"
            };
            Instance = new TestBinCutInstance(FinalRow, SourceRow, BinCutInputData);
            FlowRow row = Instance.GetBinTableRow();
            Assert.AreEqual(OpCode.BinTable, row.Opcode);
            Assert.AreEqual(expected, row.Parameter);
        }

        [TestMethod]
        [DataRow(EnumColumnName.TD, "Bin_Pmode_TD_BV")]
        [DataRow(EnumColumnName.Mbist, "Bin_Pmode_Mbist_BV")]
        [DataRow(EnumColumnName.FUNC, "Bin_Pmode_RTOS_BV")]
        [DataRow(EnumColumnName.ELB, "Bin_Pmode_ELB_BV")]
        [DataRow(EnumColumnName.ILB, "Bin_Pmode_ILB_BV")]
        public void GetBinTableRow_ShouldReturnFlowRowWithBinTableOpcode_1(EnumColumnName enumColumnName, string expected)
        {
            var binCutFlowSheetRow = new BinCutFlowSheetRow("sheetName", []);
            var newBinCutFlowSheetRow = new NewBinCutFlowSheetRow("sheetName", "job")
            {
                TableType = EnumBinCutTableType.Hv
            };
            SourceRow = new BinCutSourceItem(binCutFlowSheetRow, newBinCutFlowSheetRow, enumColumnName, "TEMP SENSOR MONITOR")
            {
                PerformanceMode = "Pmode",
            };
            Instance = new TestBinCutInstance(FinalRow, SourceRow, BinCutInputData);
            FlowRow row = Instance.GetBinTableRowBv();
            Assert.AreEqual(OpCode.BinTable, row.Opcode);
            Assert.AreEqual(expected, row.Parameter);
        }

        [TestMethod]
        public void SetVbtArg()
        {
            string subName = "BinCutInstanceBase_SetVbtArg";
            var instanceRow = new InstanceRow
            {
                TestName = "TestInstance",
                VbtName = "callinstance",
                DcSelector = "Min"
            };
            SetVbtArgTest(subName, instanceRow);
        }

        [TestMethod]
        public void SetCsArg()
        {
            string subName = "SetCsArg";
            string outputPath = Path.Combine(OutputPath, "BinCut", subName);
            string expectPath = Path.Combine(ExpectPath, "BinCut", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var instanceRow = new InstanceRow
            {
                TestName = "TestInstance",
                VbtName = "callinstance",
                DcSelector = "Min"
            };

            Instance.SetCsArg(ref instanceRow);

            // Assert
            var instanceSheet = new InstanceSheet("InstSheet");
            instanceSheet.Rows.Add(instanceRow);
            instanceSheet.Write(Path.Combine(outputPath, instanceSheet.Name + ".txt"));

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void DeleteModeInInstanceName_SegmentMatchesTargetPerformanceMode_RemovesIt()
        {
            // Arrange - fixture's SourceRow.PerformanceMode is "Pmode", so TargetPerformanceMode is "Pmode"
            string result = Instance.DeleteModeInInstanceName("X_Pmode_Y");

            // Assert
            Assert.AreEqual("X_Y", result);
        }

        [TestMethod]
        public void DeleteModeInInstanceName_SegmentDoesNotMatch_ReturnsUnchanged()
        {
            string result = Instance.DeleteModeInInstanceName("X_Other_Y");

            Assert.AreEqual("X_Other_Y", result);
        }

        [TestMethod]
        public void DeleteModeInInstanceName_SingleSegment_ReturnsUnchanged()
        {
            string result = Instance.DeleteModeInInstanceName("X");

            Assert.AreEqual("X", result);
        }

        [TestMethod]
        public void AdditionInfoInstanceName_NoTmpsSuffix_PrependsPerformanceMode()
        {
            // Arrange - fixture's SourceRow.PerformanceMode is "Pmode"
            string result = Instance.AdditionInfoInstanceName("Rest");

            // Assert
            Assert.AreEqual("Pmode_Rest", result);
        }

        [TestMethod]
        public void AdditionInfoInstanceName_TmpsSuffix_StripsTmpsFromKeyword()
        {
            // Arrange
            var binCutFlowSheetRow = new BinCutFlowSheetRow("sheetName", []);
            var newBinCutFlowSheetRow = new NewBinCutFlowSheetRow("sheetName", "job") { TableType = EnumBinCutTableType.Hv };
            SourceRow = new BinCutSourceItem(binCutFlowSheetRow, newBinCutFlowSheetRow, EnumColumnName.FUNC, "TEMP SENSOR MONITOR")
            {
                PerformanceMode = "Pmode_TMPS"
            };
            Instance = new TestBinCutInstance(FinalRow, SourceRow, BinCutInputData);

            // Act
            string result = Instance.AdditionInfoInstanceName("Rest");

            // Assert
            Assert.AreEqual("Pmode_Rest", result);
        }

        [TestMethod]
        public void GenerateAcSelector_ReturnsTyp()
        {
            string result = Instance.GenerateAcSelector();

            Assert.AreEqual("Typ", result);
        }

        [TestMethod]
        public void GenerateDcSelector_NoSelectorAndNonMatchingAllOther_ReturnsMin()
        {
            // Arrange - fixture's SourceRow.AllOther is "TEMP SENSOR MONITOR", which doesn't
            // match any HV/NV/LV pattern, so GetDcSelector falls through to its default of Min.
            string result = Instance.GenerateDcSelector(null);

            // Assert
            Assert.AreEqual("Min", result);
        }

        [TestMethod]
        public void GenerateDcSelector_WithSelectorName_DelegatesToGetDcSelector()
        {
            string result = Instance.GenerateDcSelector("HV");

            Assert.AreEqual("Max", result);
        }
    }
}
