using System.Collections.Generic;

using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.InputManager;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using ScghLib.Base;
using ScghLib.Enums;

namespace Automation.Test.UT.BistBira
{
    [TestClass]
    public class BistBiraInputManagerTests
    {
        [TestMethod]
        public void IsMultiChipLet_EmptyArray_ReturnsFalse()
        {
            // Arrange
            string[] sheets = [];

            // Act
            bool result = new BistBiraInputManager(null, null).IsMultiChipLet(sheets);

            // Assert
            Assert.IsFalse(result, "Should be false for an empty input.");
        }

        [DataTestMethod]
        [DataRow(new string[] { "A", "B", "A" }, true)]
        [DataRow(new string[] { "CPU_A1", "CPU_B1", "CPU_C1" }, false)]
        public void IsMultiChipLet_TheoryTests(string[] sheets, bool expected)
        {
            // Act
            bool result = new BistBiraInputManager(null, null).IsMultiChipLet(sheets);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("PatSets:SINGLE", MbistPatSetType.Single, MbistBinTableType.Single, false)]
        [DataRow("PatSets:BURSTNO", MbistPatSetType.BurstNo, MbistBinTableType.Burst, false)]
        [DataRow("PatSets:BURSTYES", MbistPatSetType.BurstYes, MbistBinTableType.Burst, false)]
        [DataRow("BinTable:BURST", MbistPatSetType.Single, MbistBinTableType.Burst, false)]
        [DataRow("BinTable:SINGLE", MbistPatSetType.Single, MbistBinTableType.Single, false)]
        [DataRow("MbistLoop", MbistPatSetType.Single, MbistBinTableType.Single, true)]
        [DataRow("BinTable:BURST;MbistLoop", MbistPatSetType.Single, MbistBinTableType.Burst, true)]
        [DataRow("PatSets:BURSTNO;BinTable:SINGLE;MbistLoop", MbistPatSetType.BurstNo, MbistBinTableType.Single, true)]
        [DataRow("PatSets:BURSTYES;BinTable:SINGLE;MbistLoop", MbistPatSetType.BurstYes, MbistBinTableType.Single, true)]
        [DataRow("Invalid:DATA", MbistPatSetType.Single, MbistBinTableType.Single, false)]
        public void AddSourceSheetFromFlowMain_ShouldHandleAllOptionCases(string optionInput, MbistPatSetType expectedPatSet, MbistBinTableType expectedBinTable, bool expectedMbistLoop)
        {
            // Arrange
            var flows = new List<FlowSequenceNew>
            {
                new()
                {
                    SheetName = "TestSheet",
                    SubFlowName = "",
                    Option = optionInput
                }
            };

            var service = new BistBiraInputManager(null, null);

            // Act
            List<MbistSheet> result = service.AddSourceSheetFromFlowMain(flows);

            // Assert
            Assert.AreEqual(1, result.Count, "Should process one flow");
            Assert.AreEqual(expectedPatSet, result[0].MbistPatSetType, $"Failed for input: {optionInput}");
            Assert.AreEqual(expectedBinTable, result[0].MbistBinTableType, $"Failed for input: {optionInput}");
            Assert.AreEqual(expectedMbistLoop, result[0].MbistLoop, $"Failed for input: {optionInput}");
        }

        [TestMethod]
        public void AddSourceSheetFromFlowMain_CaseInsensitivity_ReturnsCorrectType()
        {
            // Arrange
            var flows = new List<FlowSequenceNew>
            {
                new() { SheetName = "S1", Option = "patsets:burstno" ,SubFlowName = ""}
            };
            var service = new BistBiraInputManager(null, null);

            // Act
            List<MbistSheet> result = service.AddSourceSheetFromFlowMain(flows);

            // Assert
            Assert.AreEqual(MbistPatSetType.BurstNo, result[0].MbistPatSetType, "Should handle lowercase inputs via .ToUpper()");
        }
    }
}
