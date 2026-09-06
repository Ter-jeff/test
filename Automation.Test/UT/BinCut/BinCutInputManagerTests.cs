using System.Collections.Generic;
using System.IO;

using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.Static;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.Test.UT.BinCut
{
    [TestClass]
    public class BinCutInputManagerTests : FunctionTestBase
    {
        [TestMethod]
        public void AddBinCutInstanceFromTtrTableTest()
        {
            var hipSyntaxList = new List<string> { "RNGSTD VSOCDC0 NV", "RNGSTD VSOCDC0 PBC" };
            var instanceSheets = new List<BinCutInstanceSheet>();
            var postInstanceSheets = new List<BinCutInstanceSheet>();
            var binCutInputData = new BinCutInputData();
            BlockStatus.GetAutomationBlockStatus(BlockStatus.HardIp).Down = true;

            // Arrange
            LocalSpecs.TtrSummaryFileName = Path.Combine(InputPath, "borneo_documents", "A0_V04A", "Borneo_A0_TTR_V04A_X_DigHardIP.xlsx");

            // Arrange
            var manager = new BinCutInputManager(null);
            manager.AddBinCutInstanceFromTtrTable(hipSyntaxList, instanceSheets, postInstanceSheets, ref binCutInputData);

            // Assert
            Assert.AreEqual(0, instanceSheets.Count);
        }
    }
}
