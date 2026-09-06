using System.Linq;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using TestPlanLib;

namespace Automation.Test.UT.Basic
{

    [TestClass]
    public class VreTestCaseTableReaderTests
    {
        private ExcelPackage _package = new();

        [TestInitialize]
        public void Setup()
        {
            _package = new ExcelPackage();
        }

        [TestCleanup]
        public void TestCleanup()
        {
            _package?.Dispose();
        }

        private ExcelWorksheet CreateSheet()
        {
            ExcelWorksheet ws = _package.Workbook.Worksheets.Add("VRE_Test_Scenarios");

            // Header row
            ws.Cells[1, 1].Value = "Case ID";
            ws.Cells[1, 2].Value = "SubProgram";
            ws.Cells[1, 3].Value = "ProcesureName";
            ws.Cells[1, 4].Value = "InstanceName";
            ws.Cells[1, 5].Value = "Pattern1";
            ws.Cells[1, 6].Value = "Pattern2";
            ws.Cells[1, 7].Value = "User_Def";
            ws.Cells[1, 8].Value = "Expected Hard Bin";
            ws.Cells[1, 9].Value = "Expected Soft Bin";
            ws.Cells[1, 10].Value = "Level check";
            ws.Cells[1, 11].Value = "Comment";

            // Data row
            ws.Cells[2, 1].Value = 1001;
            ws.Cells[2, 2].Value = "SUB1";
            ws.Cells[2, 3].Value = "PROC_A";
            ws.Cells[2, 4].Value = "INST_1";
            ws.Cells[2, 5].Value = "PAT_A";
            ws.Cells[2, 6].Value = "PAT_B";
            ws.Cells[2, 7].Value = "UDEF";
            ws.Cells[2, 8].Value = "10";
            ws.Cells[2, 9].Value = "20";
            ws.Cells[2, 10].Value = "Y";
            ws.Cells[2, 11].Value = "Sample Comment";

            return ws;
        }

        [TestMethod]
        public void ReadAllColumn()
        {
            // Arrange
            ExcelWorksheet worksheet = CreateSheet();

            // Act
            VreTestCaseTable? table = new VreTestCaseTableReader().ReadSheet(worksheet);

            // Assert
            Assert.AreEqual("VRE_Test_Scenarios", table?.SheetName);
            Assert.AreEqual(1, table?.Rows.Count);

            VreTestCaseRow? row = table?.Rows.First();
            Assert.AreEqual(1001, row?.CaseId);
            Assert.AreEqual("SUB1", row?.SubProgram);
            Assert.AreEqual("PROC_A", row?.ProcesureName);
            Assert.AreEqual("INST_1", row?.InstanceName);
            Assert.AreEqual("UDEF", row?.UserDef);
            Assert.AreEqual("10", row?.HardBin);
            Assert.AreEqual("20", row?.SorfdBin);
            Assert.AreEqual("Y", row?.LevelCheck);
            Assert.AreEqual("Sample Comment", row?.Comment);

            CollectionAssert.AreEqual(
                new[] { "PAT_A", "PAT_B" },
                row?.Pattern.ToArray()
            );
        }

        [TestMethod]
        public void FindHeaderIndex()
        {
            // Arrange
            ExcelWorksheet worksheet = CreateSheet();

            // Act
            VreTestCaseTable? table = new VreTestCaseTableReader().ReadSheet(worksheet);

            // Assert
            Assert.IsTrue(table?.HeaderIndex.ContainsKey("Case ID"));
            Assert.IsTrue(table?.HeaderIndex.ContainsKey("Pattern1"));
            Assert.IsTrue(table?.HeaderIndex.ContainsKey("Pattern2"));
            Assert.AreEqual(1, table?.HeaderIndex["Case ID"]);
        }

        [TestMethod]
        public void MissingOtherInformation()
        {
            // Arrange
            ExcelWorksheet ws = _package.Workbook.Worksheets.Add("VRE_Test_Scenarios");

            ws.Cells[1, 1].Value = "Case ID";
            ws.Cells[2, 1].Value = 2001;

            // Act
            VreTestCaseTable? table = new VreTestCaseTableReader().ReadSheet(ws);

            // Assert
            Assert.IsNotNull(table);
            Assert.AreEqual(1, table.Rows.Count);
            Assert.AreEqual(2001, table.Rows[0].CaseId);
        }

        [TestMethod]
        public void MissingFirstHeader()
        {
            // Arrange
            ExcelWorksheet ws = _package.Workbook.Worksheets.Add("VRE_Test_Scenarios");
            ws.Cells[1, 1].Value = "WrongHeader";

            // Act
            VreTestCaseTable? result = new VreTestCaseTableReader().ReadSheet(ws);

            // Assert
            Assert.IsNull(result);
        }
    }

}
