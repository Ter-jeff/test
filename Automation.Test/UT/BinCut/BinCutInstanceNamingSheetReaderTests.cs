using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.Test.UT.BinCut
{
    [TestClass]
    public class BinCutInstanceNamingSheetReaderTests
    {
        [TestMethod]
        public void GetInitNameByOrder_ShouldReturnCorrectInitNames()
        {
            // Arrange
            var row = new BinCutInstanceRow
            {
                InitList = ["A_B_C", "D_E_F"]
            };

            var patList = new List<string> { "0,2", "1:5" };

            // Act
            List<string> result = BinCutInstanceNamingSheet.GetInitNameByOrder(row, patList);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.IsTrue(result.Contains("A"));
            Assert.IsTrue(result.Contains("C"));
            Assert.IsTrue(result.Contains("E"));
        }

        [TestMethod]
        public void GetInitNameByOrder_ShouldHandleEmptyInitList()
        {
            var row = new BinCutInstanceRow { InitList = [] };
            var patList = new List<string>();

            List<string> result = BinCutInstanceNamingSheet.GetInitNameByOrder(row, patList);

            Assert.AreNotEqual(null, result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetInitNameByOrder_ShouldIgnoreInvalidIndexes()
        {
            var row = new BinCutInstanceRow
            {
                InitList = ["X_Y"]
            };
            var patList = new List<string> { "0,5" };

            List<string> result = BinCutInstanceNamingSheet.GetInitNameByOrder(row, patList);

            Assert.AreNotEqual(null, result);
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("X", result[0]);
        }
    }
}
