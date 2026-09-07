using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlow;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class FlowSheetGeneratorBaseTests : FunctionTestBase
    {
        internal class DummyFlowSheetGenerator : FlowSheetGeneratorBase
        {
            public DummyFlowSheetGenerator(HardIpInputData hardIpInputData, string sheetName, List<HardIpPattern> hardIpPatterns) : base(hardIpInputData, sheetName, hardIpPatterns)
            {
                var hardIpParaData = new HardIpParaData(EnumBlock.HardIp);
                FlowRowGenerator = new DummyFlowRowGenerator(new HardIpInputData(hardIpParaData), "");
            }

            protected override List<HardIpPattern> DividePatterns()
            {
                return [];
            }

            protected override List<FlowRow> GenFlowBodyRows(bool shmooFlag = false, bool vtShmooFlag = false)
            {
                return [];
            }
        }

        internal class DummyFlowRowGenerator(HardIpInputData hardIpInputData, string sheetName) : FlowRowGeneratorBase(hardIpInputData, sheetName)
        {
            public override FlowRow GenBinTableRow(string voltage = "")
            {
                throw new NotImplementedException();
            }

            public override List<FlowRow> GenRunScenarioRows(string rtosBoostinst = "", bool isFirstScenario = false)
            {
                throw new NotImplementedException();
            }

            public override List<FlowRow> GenShmooRows(string labelVoltage = "")
            {
                throw new NotImplementedException();
            }

            public override List<FlowRow> GenTestRows(bool isCz2Only = false)
            {
                throw new NotImplementedException();
            }

            protected override void SetBasicInfoByPattern(HardIpPattern hardIpPattern)
            {
                throw new NotImplementedException();
            }

            public override string CreateTestFailActionByTestPlanDefine()
            {
                throw new NotImplementedException();
            }

            public override string CreateIfConditionByTestPanDefine()
            {
                throw new NotImplementedException();
            }

            public override List<FlowRow> GetIfFlowRows(string dataRow, FlowRow flowRow, List<FlowRow> flowRows)
            {
                throw new NotImplementedException();
            }
        }

        private DummyFlowSheetGenerator _generator = null!;
        private HardIpInputData _hardIpInputData = null!;
        private HardIpSheet _hardIpSheet = null!;

        [TestInitialize]
        public void Setup()
        {
            // Arrange
            var hardIpParaData = new HardIpParaData(EnumBlock.HardIp);
            _hardIpInputData = new HardIpInputData(hardIpParaData);
            _hardIpSheet = new HardIpSheet();
            _generator = new DummyFlowSheetGenerator(_hardIpInputData, "TEST_BLOCK", _hardIpSheet.Rows);
        }

        [TestMethod]
        public void GenerateFlowSheetForRfMain_ShouldReturn_OneFlowSheetWithExpectedRows()
        {
            // Arrange
            var inputRows = new List<FlowRow> { new() { Opcode = OpCode.Test, Parameter = "PATTERN_1" } };

            // Act
            List<SubFlowSheet> sheets = _generator.GenerateFlowSheetForRfMain(inputRows);

            // Assert
            Assert.AreEqual(1, sheets.Count);
            SubFlowSheet sheet = sheets.First();
            Assert.IsTrue(sheet.Name.StartsWith("Flow_TEST_BLOCK"));
            Assert.IsTrue(sheet.Rows.Any(r => r.Opcode == OpCode.Print));
            Assert.IsTrue(sheet.Rows.Any(r => r.Opcode == OpCode.Return));
        }

        [TestMethod]
        public void GenStartRows_ShouldContain_PrintStartRow()
        {
            List<FlowRow> rows = _generator.GenStartRows();

            Assert.IsTrue(rows.Any(r => r.Opcode == OpCode.Print));
        }

        [TestMethod]
        public void GenEndRows_ShouldContain_PrintStopRow_And_ReturnRow()
        {
            List<FlowRow> rows = _generator.GenEndRows();

            Assert.IsTrue(rows.Any(r => r.Opcode == OpCode.Print));
            Assert.IsTrue(rows.Any(r => r.Opcode == OpCode.Return));
        }

        [TestMethod]
        public void GenerateFlowSheet_ShouldInclude_Start_And_End_Rows()
        {
            List<SubFlowSheet> sheets = _generator.GenerateFlowSheet();

            Assert.AreEqual(1, sheets.Count);
            SubFlowSheet sheet = sheets.First();

            Assert.IsTrue(sheet.Rows.First().Opcode == OpCode.Print);
            Assert.IsTrue(sheet.Rows.Last().Opcode == OpCode.Return);
        }

    }
}
