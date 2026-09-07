using System.Collections.Generic;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.Utility
{
    [TestClass]
    public class DivideFlowTests
    {
        [TestMethod]
        public void SplitFlowSheet_UnderThreshold_DoesNotSplit()
        {
            // Arrange
            var sheet = new SubFlowSheet("MainFlow");
            for (int i = 0; i < 10; i++)
            {
                sheet.Rows.Add(new FlowRow { Opcode = "Test" });
            }

            var source = new Dictionary<string, SubFlowSheet> { ["MainFlow"] = sheet };

            // Act
            (Dictionary<string, SubFlowSheet> dividedSheets, Dictionary<string, List<string>> expansionMap) = DivideFlow.SplitFlowSheet(source, maxDivideRow: 50);

            // Assert
            Assert.AreEqual(0, dividedSheets.Count);
            Assert.AreEqual(0, expansionMap.Count);
        }

        [TestMethod]
        public void SplitFlowSheet_OverThresholdOutsideBlock_SplitsCorrectly()
        {
            // Arrange
            var sheet = new SubFlowSheet("BigFlow");
            for (int i = 0; i < 15; i++)
            {
                sheet.Rows.Add(new FlowRow { Opcode = "Nop" });
            }

            sheet.Rows.Add(new FlowRow { Opcode = "Test" });
            for (int i = 0; i < 5; i++)
            {
                sheet.Rows.Add(new FlowRow { Opcode = "Nop" });
            }

            var source = new Dictionary<string, SubFlowSheet> { ["BigFlowKey"] = sheet };

            // Act
            (Dictionary<string, SubFlowSheet> dividedSheets, Dictionary<string, List<string>> expansionMap) = DivideFlow.SplitFlowSheet(source, maxDivideRow: 14);

            // Assert
            Assert.AreEqual(1, dividedSheets.Count);
            Assert.IsTrue(expansionMap.ContainsKey("BigFlow"));
            Assert.AreEqual("Return", sheet.Rows[15].Opcode);
        }

        [TestMethod]
        public void SplitFlowSheet_OverThresholdInsideIfBlock_PostponesSplitUntilBlockCloses()
        {
            // Arrange
            var sheet = new SubFlowSheet("BlockFlow");
            for (int i = 0; i < 5; i++)
            {
                sheet.Rows.Add(new FlowRow { Opcode = "Nop" });
            }

            sheet.Rows.Add(new FlowRow { Opcode = "If" });
            sheet.Rows.Add(new FlowRow { Opcode = "Nop" });
            sheet.Rows.Add(new FlowRow { Opcode = "Endif" });
            sheet.Rows.Add(new FlowRow { Opcode = "Test" });

            var source = new Dictionary<string, SubFlowSheet> { ["BlockKey"] = sheet };

            // Act
            (Dictionary<string, SubFlowSheet> dividedSheets, Dictionary<string, List<string>> _) = DivideFlow.SplitFlowSheet(source, maxDivideRow: 5);

            // Assert
            Assert.AreEqual(1, dividedSheets.Count);
            // Splitting must happen cleanly at index 8 after Endif evaluation block drops safely to 0
            Assert.AreEqual("Return", sheet.Rows[8].Opcode);
        }

        [TestMethod]
        public void SplitFlowSheet_NotRunFlowMatched_SkipsProcessing()
        {
            // Arrange
            var sheet = new SubFlowSheet("SkippedFlow");
            for (int i = 0; i < 100; i++)
            {
                sheet.Rows.Add(new FlowRow { Opcode = "Test" });
            }

            var source = new Dictionary<string, SubFlowSheet> { ["SkippedFlowKey"] = sheet };
            var blackList = new List<string> { "SkippedFlowKey" };

            // Act
            (Dictionary<string, SubFlowSheet> dividedSheets, Dictionary<string, List<string>> _) = DivideFlow.SplitFlowSheet(source, maxDivideRow: 10, executionProfileNotRunFlows: blackList);

            // Assert
            Assert.AreEqual(0, dividedSheets.Count);
        }

        [TestMethod]
        public void InsExpFlowSheet_ValidCallMaps_InjectsSequentially()
        {
            // Arrange
            var sheet = new SubFlowSheet("ParentSheet");
            sheet.Rows.Add(new FlowRow { Opcode = "Call", Parameter = "TargetFlow" });

            var subFlowSheets = new Dictionary<string, SubFlowSheet> { ["Parent"] = sheet };
            var expSheets = new Dictionary<string, List<string>>
            {
                ["TargetFlow"] = ["TargetFlow_1", "TargetFlow_2"]
            };

            // Act
            DivideFlow.InsExpFlowSheet(subFlowSheets, expSheets);

            // Assert
            Assert.AreEqual(3, sheet.Rows.Count);
            Assert.AreEqual("TargetFlow_1", sheet.Rows[1].Parameter);
            Assert.AreEqual("TargetFlow_2", sheet.Rows[2].Parameter);
        }
    }
}
