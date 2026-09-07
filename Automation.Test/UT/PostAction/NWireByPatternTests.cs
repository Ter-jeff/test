using System.Collections.Generic;

using Automation.GenerateIgxl.PostAction.GenWireByPattern;
using Automation.Singleton;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class NWireByPatternTests
    {
        [TestMethod]
        public void WorkFlow_WhenPatternMatches_ShouldInsertNWireCall()
        {
            // 1. Arrange: Setup mock data (Assumes you've made dependencies injectable or mockable)
            var processor = new NWireByPatternMain();
            var subFlowSheet = new SubFlowSheet("");
            NwireSingleton.Instance().SettingInfo.PatternDic.Add("P1", "V1");

            // Add a dummy "Test" row that should trigger an insertion
            var flowRows = new FlowRows()
            {
                new FlowRow() { Opcode = "Test", Parameter = "MY_MATCHING_PATTERN" },
                new FlowRow() { Opcode = "Test", Parameter = "P1" }
            };
            subFlowSheet.Rows = flowRows;

            // 2. Act: Run the logic
            processor.WorkFlow(new Dictionary<string, SubFlowSheet> { { "Sheet1", subFlowSheet } });

            // 3. Assert: Verify the row was inserted at the correct index
            // The "P1" row is the one that matches the pattern, so the call is inserted right before it, at index 1
            Assert.AreEqual(3, subFlowSheet.Rows.Count, "Row count should increase by 1.");
            Assert.AreEqual("Call", subFlowSheet.Rows[1].Opcode, "An NWire call row should be inserted at index 1.");
        }
    }

}
