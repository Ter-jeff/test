using System;
using System.Collections.Generic;
using System.Data;

using Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy;
using Automation.Reader;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.VbtLib;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class ContiBaseTests
    {
        private class TestableContiBase(DcTestContiRow dcTestContiRow) : ContiBase(dcTestContiRow)
        {
            public override List<FlowRow> GenerateFlowRows()
            {
                throw new NotImplementedException();
            }

            public override List<InstanceRow> GenerateInstanceRows()
            {
                throw new NotImplementedException();
            }
        }

        private DcTestContiRow _dcRow = null!;
        private DataTable _relayTable = null!;

        [TestInitialize]
        public void Setup()
        {
            _dcRow = new DcTestContiRow
            {
                InstanceName = "TestRow",
                PinGroup = "PIN_A;PIN_B",
                Category = "Power Short"
            };

            _relayTable = new DataTable();
            _relayTable.Columns.Add("Index");
            _relayTable.Columns.Add("PinName");
            _relayTable.Columns.Add("Relay1");
            _relayTable.Columns.Add("Relay2");
            _relayTable.Rows.Add("0", "PIN_A", "1", "0");
        }

        [TestMethod]
        public void CreateTestFlowRow_ShouldSetCorrectValues()
        {
            // Arrange
            var conti = new TestableContiBase(_dcRow);

            // Act
            FlowRow row = conti.CreateTestFlowRow("MyTest", "StopOnFail", "Comment");

            // Assert
            Assert.AreEqual("Comment", row.ColumnA);
            Assert.AreEqual(OpCode.Test, row.Opcode);
            Assert.AreEqual("MyTest", row.Parameter);
            Assert.AreEqual("StopOnFail", row.FailAction);
        }

        [TestMethod]
        public void CreateBinTableFlowRow_ShouldSetOpcodeAndParameter()
        {
            // Arrange
            var conti = new TestableContiBase(_dcRow);

            // Act
            FlowRow row = conti.CreateBinTableFlowRow("MyBin");

            // Assert
            Assert.AreEqual(OpCode.BinTable, row.Opcode);
            Assert.AreEqual("MyBin", row.Parameter);
        }

        [TestMethod]
        public void SetArgumentFromCondition_ShouldCombineInterposeValues()
        {
            // Arrange
            _dcRow.Condition = "VOL=0.3";
            var conti = new TestableContiBase(_dcRow);
            var func = new Function { FunctionName = "FakeFunc", Type = ".NET", Parameters = "interposePrePat" };
            // Act
            conti.SetArgumentFromCondition(ref func);

            // Assert
            string interposeValue = func.GetParamValue("interposePrePat");
            Assert.AreEqual("PIN_A;PIN_B:VOL:0.3", interposeValue);
        }

        [TestMethod]
        public void CreateWalkingZPatternNameOnly_ShouldReturnExpectedName()
        {
            // Arrange
            LocalSpecs.CurrentProject = "MyProj";
            var conti = new TestableContiBase(_dcRow);

            // Act
            string result = conti.CreateWalkingZPatternNameOnly(true);

            // Assert
            StringAssert.StartsWith(result, $"MyProj_Pattern_{_dcRow.PinGroup}");
            StringAssert.Contains(result, "WalkingZ");
        }

        [TestMethod]
        public void CreateName_ProjectNameEndsWithDashandIncludePattern()
        {
            // Arrange
            LocalSpecs.CurrentProject = "ABC-DEF";
            var conti = new TestableContiBase(_dcRow);

            // Act
            string result = conti.CreateWalkingZPatternNameOnly(true);

            // Assert
            StringAssert.StartsWith(result, $"ABC_DEF_Pattern_{_dcRow.PinGroup}");
            StringAssert.Contains(result, "WalkingZ");
        }

        [TestMethod]
        public void CreateName_ProjectNameEndsWithDash()
        {
            // Arrange
            LocalSpecs.CurrentProject = "ABC-DEF";
            var conti = new TestableContiBase(_dcRow);

            // Act
            string result = conti.CreateWalkingZPatternNameOnly(false);

            // Assert
            StringAssert.StartsWith(result, $"ABC_DEF_{_dcRow.PinGroup}");
            StringAssert.Contains(result, "WalkingZ");
        }

        [TestMethod]
        public void SetArgFromCondition_WithDcKey_ReturnsDcValue()
        {
            // Arrange
            var conti = new TestableContiBase(_dcRow);
            conti.DcTestContiRow.ConditionDict.Add("DC", "MyDcCategory");

            // Act
            string result = conti.SetCategoryFromCondition();

            // Assert
            Assert.AreEqual("MyDcCategory", result);

        }

        [TestMethod]
        public void SetArgFromCondition_WithoutDcKey_ReturnsDefaultValue()
        {
            // Arrange
            var conti = new TestableContiBase(_dcRow);
            conti.DcTestContiRow.ConditionDict.Add("AC", "MyAcCategory");

            string result = conti.SetCategoryFromCondition();

            // Assert
            Assert.AreEqual("Conti_X_X_X", result);
        }

    }
}
