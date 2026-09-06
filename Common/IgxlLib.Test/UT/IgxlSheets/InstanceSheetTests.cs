using System.Collections.Generic;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class InstanceSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void InstanceSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "Instance";

            // Act
            var instanceSheet = new InstanceSheet(sheetName);

            // Assert
            Assert.IsNotNull(instanceSheet);
            Assert.AreEqual(sheetName, instanceSheet.Name);
            Assert.AreEqual("DTTestInstancesSheet", instanceSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.TestInstances, instanceSheet.IgxlSheetName);
        }

        [TestMethod]
        public void InstanceSheet_AddRow()
        {
            // Arrange
            var instanceSheet = new InstanceSheet("Instance");
            var instanceRow = new InstanceRow { TestName = "Test1" };

            // Act
            instanceSheet.AddRow(instanceRow);

            // Assert
            Assert.AreEqual(1, instanceSheet.Rows.Count);
        }

        [TestMethod]
        public void InstanceSheet_AddRows()
        {
            // Arrange
            var instanceSheet = new InstanceSheet("Instance");
            var rows = new List<InstanceRow>
            {
                new() { TestName = "Test1"     },
                new() { TestName = "Test2" },
                new() { TestName = "Test3" }
            };

            // Act
            instanceSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, instanceSheet.Rows.Count);
        }

        [TestMethod]
        public void InstanceSheet_RemoveRow()
        {
            // Arrange
            var instanceSheet = new InstanceSheet("Instance");
            var row1 = new InstanceRow { TestName = "Test1" };
            var row2 = new InstanceRow { TestName = "Test2" };
            instanceSheet.AddRow(row1);
            instanceSheet.AddRow(row2);

            // Act
            instanceSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, instanceSheet.Rows.Count);
            Assert.AreEqual("Test2", instanceSheet.Rows[0].TestName);
        }

        [TestMethod]
        public void InstanceSheet_InsertRow()
        {
            // Arrange
            var instanceSheet = new InstanceSheet("Instance");
            var row1 = new InstanceRow { TestName = "Test1" };
            var row2 = new InstanceRow { TestName = "Test2" };
            var rowToInsert = new InstanceRow { TestName = "Test3" };
            instanceSheet.AddRow(row1);
            instanceSheet.AddRow(row2);

            // Act
            int index = instanceSheet.InsertRow(1, rowToInsert);

            // Assert
            Assert.AreEqual(2, index);
            Assert.AreEqual(3, instanceSheet.Rows.Count);
            Assert.AreEqual("Test3", instanceSheet.Rows[1].TestName);
        }

        [TestMethod]
        public void InstanceSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var instanceSheet = new InstanceSheet("Instance");

            // Assert
            Assert.AreEqual("DTTestInstancesSheet", instanceSheet.SheetType);
        }

        [TestMethod]
        public void InstanceSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var instanceSheet = new InstanceSheet("Instance");

            // Assert
            Assert.IsNotNull(instanceSheet.GetErrors());
            Assert.AreEqual(0, instanceSheet.GetErrors().Count);
        }

        [TestMethod]
        public void InstanceSheet_Name_CanBeSet()
        {
            // Arrange
            var instanceSheet = new InstanceSheet("Instance")
            {
                // Act
                Name = "NewInstanceName"
            };

            // Assert
            Assert.AreEqual("NewInstanceName", instanceSheet.Name);
        }

        [TestMethod]
        public void InstanceSheet_GetUsedDcSpecs_ReturnsDistinctValidCategoryAndSelectorPairs()
        {
            // Arrange
            var instanceSheet = new InstanceSheet("Instance");

            // Add combinations to verify filtering, grouping, and duplication removal
            instanceSheet.AddRow(new InstanceRow { TestName = "T1", DcCategory = "DC_CatA", DcSelector = "Sel1" });
            instanceSheet.AddRow(new InstanceRow { TestName = "T2", DcCategory = "DC_CatA", DcSelector = "Sel1" });
            instanceSheet.AddRow(new InstanceRow { TestName = "T3", DcCategory = "DC_CatB", DcSelector = "Sel2" });
            instanceSheet.AddRow(new InstanceRow { TestName = "T4", DcCategory = "", DcSelector = "Sel3" });
            instanceSheet.AddRow(new InstanceRow { TestName = "T5", DcCategory = "DC_CatC", DcSelector = null });

            // Act
            List<System.Tuple<string, string>> result = instanceSheet.GetUsedDcSpecs();

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(2, result.Count, "Should filter out empty/null fields and group identical pairs.");

            // Validate first unique pair
            Assert.AreEqual("DC_CatA", result[0].Item1);
            Assert.AreEqual("Sel1", result[0].Item2);

            // Validate second unique pair
            Assert.AreEqual("DC_CatB", result[1].Item1);
            Assert.AreEqual("Sel2", result[1].Item2);
        }

        [TestMethod]
        public void InstanceSheet_AddFooter_AppendsCorrectlyConfiguredFooterRow()
        {
            // Arrange
            var instanceSheet = new InstanceSheet("Instance");
            string blockName = "Functional";

            // Act
            instanceSheet.AddFooter(blockName);

            // Assert
            Assert.AreEqual(1, instanceSheet.Rows.Count);

            InstanceRow footerRow = instanceSheet.Rows[0];
            Assert.AreEqual("Functional_Footer_1", footerRow.TestName);
            Assert.AreEqual("VBT", footerRow.VbtType);
            Assert.AreEqual("PrintInfo", footerRow.ArgList);
            Assert.AreEqual("Print_Footer", footerRow.VbtName);

            // Verify structural arguments collection assignment
            Assert.AreEqual(1, footerRow.Args.Count);
            Assert.AreEqual("Functional", footerRow.Args[0]);
        }

        [TestMethod]
        public void InstanceSheet_AddHeader_AppendsCorrectlyConfiguredHeaderRow()
        {
            // Arrange
            var instanceSheet = new InstanceSheet("Instance");
            string blockName = "Scan";

            // Act
            instanceSheet.AddHeader(blockName);

            // Assert
            Assert.AreEqual(1, instanceSheet.Rows.Count);

            InstanceRow headerRow = instanceSheet.Rows[0];
            Assert.AreEqual("Scan_Header_1", headerRow.TestName);
            Assert.AreEqual("VBT", headerRow.VbtType);
            Assert.AreEqual("PrintInfo", headerRow.ArgList);
            Assert.AreEqual("Print_Header", headerRow.VbtName);

            // Verify structural arguments collection assignment
            Assert.AreEqual(1, headerRow.Args.Count);
            Assert.AreEqual("Scan", headerRow.Args[0]);
        }

        [TestMethod]
        public void InstanceSheet_AddHeaderFooter_RemovesPrefixAndAddsBothRows()
        {
            // Arrange
            var instanceSheet = new InstanceSheet("Instance");
            string inputSheetName = "Flow_PatternBurst";

            // Act
            instanceSheet.AddHeaderFooter(inputSheetName);

            // Assert
            Assert.AreEqual(2, instanceSheet.Rows.Count, "Should append exactly one header and one footer.");

            // Verify Header is added first with prefix removed
            InstanceRow headerRow = instanceSheet.Rows[0];
            Assert.AreEqual("PatternBurst_Header_1", headerRow.TestName);
            Assert.AreEqual("PatternBurst", headerRow.Args[0]);

            // Verify Footer is added second with prefix removed
            InstanceRow footerRow = instanceSheet.Rows[1];
            Assert.AreEqual("PatternBurst_Footer_1", footerRow.TestName);
            Assert.AreEqual("PatternBurst", footerRow.Args[0]);
        }
    }
}
