using System.Collections.Generic;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class PatSetTests
    {
        [TestMethod]
        public void PatSet_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var patSet = new PatSet();

            // Assert
            Assert.IsInstanceOfType(patSet, typeof(IgxlRow));
        }

        [TestMethod]
        public void PatSet_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var patSet = new PatSet();

            // Assert
            Assert.IsInstanceOfType(patSet, typeof(IgxlRow));
        }

        [TestMethod]
        public void PatSet_PatSetName_CanBeSet()
        {
            // Arrange
            var patSet = new PatSet
            {
                // Act
                PatSetName = "TestPatSet"
            };

            // Assert
            Assert.AreEqual("TestPatSet", patSet.PatSetName);
        }

        [TestMethod]
        public void PatSet_Domain_CanBeSet()
        {
            // Arrange
            var patSet = new PatSet
            {
                // Act
                Domain = "TestDomain"
            };

            // Assert
            Assert.AreEqual("TestDomain", patSet.Domain);
        }

        [TestMethod]
        public void PatSet_PatSetRows_InitializedAsEmptyList()
        {
            // Arrange & Act
            var patSet = new PatSet();

            // Assert
            Assert.AreEqual(0, patSet.PatSetRows.Count);
        }

        [TestMethod]
        public void PatSet_AddRow_AddsPatSetRowToCollection()
        {
            // Arrange
            var patSet = new PatSet();
            var row = new PatSetRow { PatternSet = "Pat1" };

            // Act
            patSet.AddRow(row);

            // Assert
            Assert.AreEqual(1, patSet.PatSetRows.Count);
            Assert.AreEqual("Pat1", patSet.PatSetRows[0].PatternSet);
        }

        [TestMethod]
        public void PatSet_AddRow_MultipleRows()
        {
            // Arrange
            var patSet = new PatSet();
            var row1 = new PatSetRow { PatternSet = "Pat1" };
            var row2 = new PatSetRow { PatternSet = "Pat2" };
            var row3 = new PatSetRow { PatternSet = "Pat3" };

            // Act
            patSet.AddRow(row1);
            patSet.AddRow(row2);
            patSet.AddRow(row3);

            // Assert
            Assert.AreEqual(3, patSet.PatSetRows.Count);
            Assert.AreEqual("Pat1", patSet.PatSetRows[0].PatternSet);
            Assert.AreEqual("Pat2", patSet.PatSetRows[1].PatternSet);
            Assert.AreEqual("Pat3", patSet.PatSetRows[2].PatternSet);
        }

        [TestMethod]
        public void PatSet_GetNewPatSetNameWithX_EmptyPatterns_ReturnsEmpty()
        {
            // Arrange
            var patterns = new List<string>();

            // Act
            string result = PatSet.GetNewPatSetNameWithX(patterns);

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void PatSet_GetNewPatSetNameWithX_SinglePattern()
        {
            // Arrange
            var patterns = new List<string> { "A_B_C" };

            // Act
            string result = PatSet.GetNewPatSetNameWithX(patterns);

            // Assert
            Assert.AreEqual("A_B_C", result);
        }

        [TestMethod]
        public void PatSet_GetNewPatSetNameWithX_MultipleIdenticalPatterns()
        {
            // Arrange
            var patterns = new List<string> { "ABC", "ABC", "ABC" };

            // Act
            string result = PatSet.GetNewPatSetNameWithX(patterns);

            // Assert
            Assert.AreEqual("ABC", result);
        }

        [TestMethod]
        public void PatSet_GetNewPatSetNameWithX_DifferentPatterns()
        {
            // Arrange
            var patterns = new List<string> { "ABC", "ABX", "AYC" };

            // Act
            string result = PatSet.GetNewPatSetNameWithX(patterns);

            // Assert
            Assert.AreEqual("AXX", result);
        }

        [TestMethod]
        public void PatSet_GetPatSetName_EmptyPatterns_ReturnsEmpty()
        {
            // Arrange
            var patterns = new List<string>();

            // Act
            string result = PatSet.GetPatSetName(patterns);

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void PatSet_GetPatSetName_SinglePattern()
        {
            // Arrange
            var patterns = new List<string> { "Pattern1" };

            // Act
            string result = PatSet.GetPatSetName(patterns);

            // Assert
            Assert.AreEqual("Pattern1", result);
        }

        [TestMethod]
        public void PatSet_GetPatSetName_MultiplePatterns_DefaultOnlyPl()
        {
            // Arrange
            var patterns = new List<string>
            {
                "PL_Test_SL_1",
                "PL_Test_SL_1",
                "PL_Test_SL_1"
            };

            // Act
            string result = PatSet.GetPatSetName(patterns);

            // Assert
            Assert.AreEqual("PL_Test_SL_1", result);
        }

        [TestMethod]
        public void PatSet_GetNewPatSetName_SinglePattern()
        {
            // Arrange
            var patterns = new List<string> { "A_B_C" };

            // Act
            string result = PatSet.GetNewPatSetName(patterns);

            // Assert
            Assert.AreEqual("A_B_C", result);
        }

        [TestMethod]
        public void PatSet_GetNewPatSetName_MultipleIdentical()
        {
            // Arrange
            var patterns = new List<string> { "A_B_C", "A_B_C", "A_B_C" };

            // Act
            string result = PatSet.GetNewPatSetName(patterns);

            // Assert
            Assert.AreEqual("A_B_C", result);
        }

        [TestMethod]
        public void PatSet_GetNewPatSetName_WithDifferences()
        {
            // Arrange
            var patterns = new List<string>
            {
                "A_B_C_D",
                "A_X_C_D",
                "A_B_Y_D"
            };

            // Act
            string result = PatSet.GetNewPatSetName(patterns);

            // Assert
            Assert.IsNotNull(result);
            Assert.IsFalse(string.IsNullOrEmpty(result));
        }

        [TestMethod]
        public void PatSet_Properties_AllCanBeSet()
        {
            // Arrange
            var patSet = new PatSet
            {
                // Act
                PatSetName = "TestSet",
                Domain = "TestDomain",
                SheetName = "Sheet1",
                RowNum = 5
            };

            // Assert
            Assert.AreEqual("TestSet", patSet.PatSetName);
            Assert.AreEqual("TestDomain", patSet.Domain);
            Assert.AreEqual("Sheet1", patSet.SheetName);
            Assert.AreEqual(5, patSet.RowNum);
        }
    }
}
