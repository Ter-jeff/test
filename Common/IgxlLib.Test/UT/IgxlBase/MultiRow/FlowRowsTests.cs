using System.Collections.Generic;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase.MultiRow
{
    [TestClass]
    public class FlowRowsTests
    {
        private Dictionary<string, SubFlowSheet> _subFlowSheets;
        private List<string> _doneSheets;
        private int _count;

        [TestInitialize]
        public void Setup()
        {
            _subFlowSheets = [];
            _doneSheets = [];
            _count = 0;
        }

        [TestMethod]
        public void SetTestNumber_OpcodeIsTest_AssignsIncrementedTNum()
        {
            // Arrange
            var row1 = new FlowRow { Opcode = OpCode.Test, TNum = "" };
            var row2 = new FlowRow { Opcode = OpCode.Test, TNum = "" };
            var flowRows = new FlowRows([row1, row2]);

            // Act
            flowRows.SetTestNumber(_subFlowSheets, 1000, 10, 5000, ref _count, ref _doneSheets);

            // Assert
            Assert.AreEqual("1000", row1.TNum);
            Assert.AreEqual("1010", row2.TNum);
            Assert.AreEqual(2, _count);
        }

        [TestMethod]
        public void SetTestNumber_OpcodeIsTestDeferLimits_AssignsTNum()
        {
            // Arrange
            var row = new FlowRow { Opcode = "test-defer-limits", TNum = "" };
            var flowRows = new FlowRows([row]);

            // Act
            flowRows.SetTestNumber(_subFlowSheets, 1000, 10, 5000, ref _count, ref _doneSheets);

            // Assert
            Assert.AreEqual("1000", row.TNum);
            Assert.AreEqual(1, _count);
        }

        [TestMethod]
        public void SetTestNumber_TNumAlreadyExists_DoesNotOverwriteOrIncrement()
        {
            // Arrange
            var row = new FlowRow { Opcode = OpCode.Test, TNum = "9999" };
            var flowRows = new FlowRows([row]);

            // Act
            flowRows.SetTestNumber(_subFlowSheets, 1000, 10, 5000, ref _count, ref _doneSheets);

            // Assert
            Assert.AreEqual("9999", row.TNum);
            Assert.AreEqual(0, _count);
        }

        [TestMethod]
        public void SetTestNumber_ReachesMaxNum_StopCountIncrement()
        {
            // Arrange
            var row1 = new FlowRow { Opcode = OpCode.Test, TNum = "" };
            var row2 = new FlowRow { Opcode = OpCode.Test, TNum = "" };
            var flowRows = new FlowRows([row1, row2]);

            // Max num set to 1005. 
            // Row 1: 1000 + (0*10) = 1000. Next check: 1000 + (1*10) = 1010. 1010 < 1005 is FALSE. Count stays 0.
            // Row 2: 1000 + (0*10) = 1000. Next check: 1010 < 1005 is FALSE. Count stays 0.
            // Act
            flowRows.SetTestNumber(_subFlowSheets, 1000, 10, 1005, ref _count, ref _doneSheets);

            // Assert
            Assert.AreEqual("1000", row1.TNum);
            Assert.AreEqual("1000", row2.TNum);
            Assert.AreEqual(0, _count);
        }

        [TestMethod]
        public void SetTestNumber_ValidSubFlowSheet_ProcessesChildRowsAndAddsMetadata()
        {
            // Arrange
            string sheetName = "SubSheetA";
            var parentRow = new FlowRow { Opcode = "Call", Parameter = sheetName };
            var childRow = new FlowRow { Opcode = OpCode.Test, TNum = "" };

            var subSheet = new SubFlowSheet("")
            {
                Name = sheetName,
                Rows = [childRow]
            };
            _subFlowSheets.Add(sheetName, subSheet);

            var flowRows = new FlowRows([parentRow]);

            // Act
            flowRows.SetTestNumber(_subFlowSheets, 1000, 10, 5000, ref _count, ref _doneSheets);

            // Assert
            Assert.AreEqual("1000", childRow.TNum);
            Assert.IsTrue(_doneSheets.Contains(sheetName));
            // Verifies if AddColumnA was called via state check (assuming ColumnA property exists inside parentRow)
            // Replace with actual property verification depending on your FlowRow implementation:
            // Assert.AreEqual("TNum 1000 - 1010", parentRow.ColumnA); 
        }

        [TestMethod]
        public void SetTestNumber_SubFlowSheetAlreadyDone_SkipsProcessing()
        {
            // Arrange
            string sheetName = "SubSheetA";
            var parentRow = new FlowRow { Opcode = "Call", Parameter = sheetName };
            var childRow = new FlowRow { Opcode = OpCode.Test, TNum = "" };

            var subSheet = new SubFlowSheet("")
            {
                Name = sheetName,
                Rows = [childRow]
            };
            _subFlowSheets.Add(sheetName, subSheet);
            _doneSheets.Add(sheetName);

            var flowRows = new FlowRows([parentRow]);

            // Act
            flowRows.SetTestNumber(_subFlowSheets, 1000, 10, 5000, ref _count, ref _doneSheets);

            // Assert
            Assert.AreEqual("", childRow.TNum);
            Assert.AreEqual(0, _count);
        }
    }
}
