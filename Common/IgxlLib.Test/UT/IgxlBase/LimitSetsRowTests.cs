using System.Collections.Generic;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class LimitSetsRowTests
    {
        [TestMethod]
        public void LimitSetsRow_DefaultConstructor_InitializesWithDefaults()
        {
            // Arrange & Act
            var limitSetsRow = new LimitSetsRow();

            // Assert
            Assert.AreEqual("", limitSetsRow.Fields);
            Assert.AreEqual(">=", limitSetsRow.LowCompSign);
            Assert.AreEqual("<=", limitSetsRow.HighCompSign);
            Assert.AreEqual("none", limitSetsRow.Result);
        }

        [TestMethod]
        public void LimitSetsRow_Constructor_WithParameters_InitializesAllProperties()
        {
            // Arrange
            var categories = new List<LimitSetsItem>();

            // Act
            var limitSetsRow = new LimitSetsRow(
                "Row1",
                "FlowTable1",
                "Instance1",
                "TestName1",
                "TestNum1",
                categories,
                "1",
                "V",
                "0.00",
                "PassSort1",
                "FailSort1",
                "100",
                "200",
                "PassAction1",
                "FailAction1",
                "Assume1",
                "Site1",
                "Test comment"
            );

            // Assert
            Assert.AreEqual("Row1", limitSetsRow.Row);
            Assert.AreEqual("FlowTable1", limitSetsRow.FlowTable);
            Assert.AreEqual("Instance1", limitSetsRow.Instance);
            Assert.AreEqual("TestName1", limitSetsRow.TestName);
            Assert.AreEqual("TestNum1", limitSetsRow.TestNumber);
            Assert.AreEqual("1", limitSetsRow.Scale);
            Assert.AreEqual("V", limitSetsRow.Units);
            Assert.AreEqual("0.00", limitSetsRow.Format);
            Assert.AreEqual("PassSort1", limitSetsRow.PassSort);
            Assert.AreEqual("FailSort1", limitSetsRow.FailSort);
            Assert.AreEqual("100", limitSetsRow.PassBin);
            Assert.AreEqual("200", limitSetsRow.FailBin);
            Assert.AreEqual("Test comment", limitSetsRow.Comment);
        }

        [TestMethod]
        public void LimitSetsRow_SetComparisonSigns_UpdatesValues()
        {
            // Arrange
            var limitSetsRow = new LimitSetsRow
            {
                // Act
                LowCompSign = ">",
                HighCompSign = "<"
            };

            // Assert
            Assert.AreEqual(">", limitSetsRow.LowCompSign);
            Assert.AreEqual("<", limitSetsRow.HighCompSign);
        }

        [TestMethod]
        public void LimitSetsRow_SetSortAndBin_UpdatesValues()
        {
            // Arrange
            var limitSetsRow = new LimitSetsRow
            {
                // Act
                SortName = "Sort1",
                BinName = "Bin1",
                PassSort = "PSort",
                FailSort = "FSort",
                PassBin = "100",
                FailBin = "200"
            };

            // Assert
            Assert.AreEqual("Sort1", limitSetsRow.SortName);
            Assert.AreEqual("Bin1", limitSetsRow.BinName);
            Assert.AreEqual("PSort", limitSetsRow.PassSort);
            Assert.AreEqual("FSort", limitSetsRow.FailSort);
        }

        [TestMethod]
        public void LimitSetsRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var limitSetsRow = new LimitSetsRow();

            // Assert
            Assert.IsInstanceOfType(limitSetsRow, typeof(IgxlRow));
        }

        [TestMethod]
        public void LimitSetsRow_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var row1 = new LimitSetsRow { TestName = "Test1", Scale = "1" };
            var row2 = new LimitSetsRow { TestName = "Test2", Scale = "10" };

            // Assert
            Assert.AreEqual("Test1", row1.TestName);
            Assert.AreEqual("Test2", row2.TestName);
            Assert.AreNotEqual(row1.Scale, row2.Scale);
        }
    }
}
