using System;
using System.Collections.Generic;
using System.IO;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class PatSetSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void PatSetSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "PatSet";

            // Act
            var patSetSheet = new PatSetSheet(sheetName);

            // Assert
            Assert.IsNotNull(patSetSheet);
            Assert.AreEqual(sheetName, patSetSheet.Name);
            Assert.AreEqual("DTPatternSetSheet", patSetSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.PatternSet, patSetSheet.IgxlSheetName);
        }

        [TestMethod]
        public void PatSetSheet_GetIgxlSheetsVersion()
        {
            // Arrange & Act
            Dictionary<string, Dictionary<string, IGDataXML.IGXLSheetsVersion.SheetInfo>> versionDict = PatSetSheet.GetIgxlSheetsVersion();

            // Assert
            Assert.IsNotNull(versionDict);
            Assert.IsTrue(versionDict.Count > 0);
        }

        [TestMethod]
        public void PatSetSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var patSetSheet = new PatSetSheet("PatSet");

            // Assert
            Assert.AreEqual("DTPatternSetSheet", patSetSheet.SheetType);
        }

        [TestMethod]
        public void PatSetSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var patSetSheet = new PatSetSheet("PatSet");

            // Assert
            Assert.IsNotNull(patSetSheet.GetErrors());
            Assert.AreEqual(0, patSetSheet.GetErrors().Count);
        }

        [TestMethod]
        public void PatSetSheet_Name_CanBeSet()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet")
            {
                // Act
                Name = "NewPatSetName"
            };

            // Assert
            Assert.AreEqual("NewPatSetName", patSetSheet.Name);
        }

        [TestMethod]
        public void PatSetSheet_AddRow()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");
            var patSet = new PatSet { PatSetName = "PatSet1" };

            // Act
            patSetSheet.AddRow(patSet);

            // Assert
            Assert.AreEqual(1, patSetSheet.Rows.Count);
            Assert.AreEqual("PatSet1", patSetSheet.Rows[0].PatSetName);
        }

        [TestMethod]
        public void PatSetSheet_AddMultipleRows()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");

            // Act
            for (int i = 0; i < 5; i++)
            {
                patSetSheet.AddRow(new PatSet { PatSetName = $"PatSet{i}" });
            }

            // Assert
            Assert.AreEqual(5, patSetSheet.Rows.Count);
        }

        [TestMethod]
        public void PatSetSheet_RemoveRow()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");
            var row1 = new PatSet { PatSetName = "PatSet1" };
            var row2 = new PatSet { PatSetName = "PatSet2" };
            patSetSheet.AddRow(row1);
            patSetSheet.AddRow(row2);

            // Act
            patSetSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, patSetSheet.Rows.Count);
            Assert.AreEqual("PatSet2", patSetSheet.Rows[0].PatSetName);
        }

        [TestMethod]
        public void PatSetSheet_Rows_AreInitialized()
        {
            // Arrange & Act
            var patSetSheet = new PatSetSheet("PatSet");

            // Assert
            Assert.IsInstanceOfType(patSetSheet.Rows, typeof(List<PatSet>));
            Assert.AreEqual(0, patSetSheet.Rows.Count);
        }

        [TestMethod]
        public void PatSetSheet_IgxlSheetName_IsCorrect()
        {
            // Arrange & Act
            var patSetSheet = new PatSetSheet("PatSet");

            // Assert
            Assert.AreEqual(IgxlSheetNames.PatternSet, patSetSheet.IgxlSheetName);
        }

        [TestMethod]
        public void PatSetSheet_Write_WithEmptyRows_DoesNotThrow()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");
            string tempFileName = Path.Combine(Path.GetTempPath(), $"PatSet_{Guid.NewGuid()}.txt");

            try
            {
                // Act
                patSetSheet.Write(tempFileName);

                // Assert (no exception should be thrown)
                Assert.IsFalse(File.Exists(tempFileName));
            }
            finally
            {
                if (File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
            }
        }

        [TestMethod]
        public void PatSetSheet_Write_WithVersion()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");
            string tempFileName = Path.Combine(Path.GetTempPath(), $"PatSet_{Guid.NewGuid()}.txt");

            try
            {
                // Act
                patSetSheet.Write(tempFileName, "2.0");

                // Assert
                Assert.IsFalse(File.Exists(tempFileName));
            }
            finally
            {
                if (File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
            }
        }

        [TestMethod]
        public void PatSetSheet_ClearRows()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");
            patSetSheet.AddRow(new PatSet { PatSetName = "PatSet1" });
            patSetSheet.AddRow(new PatSet { PatSetName = "PatSet2" });
            Assert.AreEqual(2, patSetSheet.Rows.Count);

            // Act
            patSetSheet.Rows.Clear();

            // Assert
            Assert.AreEqual(0, patSetSheet.Rows.Count);
        }

        [TestMethod]
        public void PatSetSheet_GetRow_RetrievesCorrectly()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");
            var patSet1 = new PatSet { PatSetName = "PatSet1" };
            var patSet2 = new PatSet { PatSetName = "PatSet2" };
            patSetSheet.AddRow(patSet1);
            patSetSheet.AddRow(patSet2);

            // Act
            PatSet retrieved = patSetSheet.Rows[1];

            // Assert
            Assert.AreEqual("PatSet2", retrieved.PatSetName);
        }

        [TestMethod]
        public void PatSetSheet_Contains_ChecksForRow()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");
            var patSet = new PatSet { PatSetName = "PatSet1" };
            patSetSheet.AddRow(patSet);

            // Act
            bool contains = patSetSheet.Rows.Contains(patSet);

            // Assert
            Assert.IsTrue(contains);
        }

        [TestMethod]
        public void PatSetSheet_MultipleOperations_InSequence()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");

            // Act
            patSetSheet.AddRow(new PatSet { PatSetName = "PatSet1" });
            patSetSheet.AddRow(new PatSet { PatSetName = "PatSet2" });
            patSetSheet.AddRow(new PatSet { PatSetName = "PatSet3" });
            patSetSheet.RemoveRow(patSetSheet.Rows[1]);

            // Assert
            Assert.AreEqual(2, patSetSheet.Rows.Count);
            Assert.AreEqual("PatSet1", patSetSheet.Rows[0].PatSetName);
            Assert.AreEqual("PatSet3", patSetSheet.Rows[1].PatSetName);
        }

        [TestMethod]
        public void PatSetSheet_Name_CanBeChanged()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("Original")
            {
                // Act
                Name = "Changed"
            };

            // Assert
            Assert.AreEqual("Changed", patSetSheet.Name);
        }

        [TestMethod]
        public void PatSetSheet_ExistPatSet_ReturnsCaseInsensitiveHashSetOfKeys()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");
            patSetSheet.AddRow(new PatSet { PatSetName = "PatSet_Alpha" });
            patSetSheet.AddRow(new PatSet { PatSetName = "patset_beta" });

            // Act
            HashSet<string> keys = patSetSheet.ExistPatSet;

            // Assert
            Assert.IsNotNull(keys);
            Assert.AreEqual(2, keys.Count);
            Assert.IsTrue(keys.Contains("PATSET_ALPHA"));
            Assert.IsTrue(keys.Contains("PatSet_Beta"));
        }

        [TestMethod]
        public void PatSetSheet_IsExist_ReturnsTrueIfKeyExistsRegardlessOfCase()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");
            patSetSheet.AddRow(new PatSet { PatSetName = "Target_PatSet" });

            // Act & Assert
            Assert.IsTrue(patSetSheet.IsExist("target_patset"));
            Assert.IsTrue(patSetSheet.IsExist("TARGET_PATSET"));
            Assert.IsFalse(patSetSheet.IsExist("NonExistent_PatSet"));
        }

        [TestMethod]
        public void PatSetSheet_PatSetRowDic_ExposesDictionaryCorrectly()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");
            var row = new PatSet { PatSetName = "DicKey1" };
            patSetSheet.AddRow(row);

            // Act
            Dictionary<string, PatSet> dic = patSetSheet.PatSetRowDic;

            // Assert
            Assert.IsNotNull(dic);
            Assert.AreEqual(1, dic.Count);
            Assert.AreSame(row, dic["dickey1"]);
        }

        [TestMethod]
        public void PatSetSheet_IsExistTheSamePatSet_ReturnsTrueAndOutputsKey_OnExactMatch()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");

            // Build out mock baseline data for underlying row comparison validation
            var existingRow = new PatSet { PatSetName = "Pattern_Group_A" };
            var nestedRowMock = new PatSetRow();
            existingRow.PatSetRows.Add(nestedRowMock);
            patSetSheet.AddRow(existingRow);

            // Prepare candidate search criteria that matches structural configuration
            var matchingCandidate = new PatSet { PatSetName = "pattern_group_a" };
            matchingCandidate.PatSetRows.Add(nestedRowMock);

            // Act
            bool result = patSetSheet.IsExistTheSamePatSet(matchingCandidate, out string outputName);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual("Pattern_Group_A", outputName);
        }

        [TestMethod]
        public void PatSetSheet_IsExistTheSamePatSet_ReturnsFalse_WhenNestedRowCountsDiffer()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");
            var existingRow = new PatSet { PatSetName = "Pattern_Group_B" };
            existingRow.PatSetRows.Add(new PatSetRow());
            patSetSheet.AddRow(existingRow);

            var candidateWithNoNestedRows = new PatSet { PatSetName = "Pattern_Group_B" };

            // Act
            bool result = patSetSheet.IsExistTheSamePatSet(candidateWithNoNestedRows, out string outputName);

            // Assert
            Assert.IsFalse(result);
            Assert.AreEqual(string.Empty, outputName);
        }

        [TestMethod]
        public void PatSetSheet_IsExistTheSamePatSet_ReturnsFalse_WhenCompareRowFails()
        {
            // Arrange
            var patSetSheet = new PatSetSheet("PatSet");
            var existingRow = new PatSet { PatSetName = "Pattern_Group_C" };
            existingRow.PatSetRows.Add(new PatSetRow { PatternSet = "P1", File = "P1" });
            patSetSheet.AddRow(existingRow);

            var breakingCandidate = new PatSet { PatSetName = "Pattern_Group_C" };
            breakingCandidate.PatSetRows.Add(new PatSetRow { PatternSet = "P2", File = "P2" });

            // Act
            bool result = patSetSheet.IsExistTheSamePatSet(breakingCandidate, out string outputName);

            // Assert
            Assert.IsFalse(result);
            Assert.AreEqual(string.Empty, outputName);
        }
    }
}
