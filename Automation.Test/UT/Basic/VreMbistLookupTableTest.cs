using System.Linq;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using TestPlanLib;

namespace Automation.Test.UT.Basic
{

    [TestClass]
    public class VreMbistLookupTableReaderTests
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

        private ExcelWorksheet CreateTable()
        {
            ExcelWorksheet? ws = _package.Workbook.Worksheets.Add("VRE_Mbist_Lookup");
            // Header
            ws.Cells[1, 1].Value = "Harvesting";
            ws.Cells[1, 2].Value = "Server";
            ws.Cells[1, 3].Value = "Memory Group";
            ws.Cells[1, 4].Value = "P mode";
            ws.Cells[1, 5].Value = "Exclude Pattern";

            // Data
            ws.Cells[2, 1].Value = "HARV_A";
            ws.Cells[2, 2].Value = "SRV1";
            ws.Cells[2, 3].Value = "MEM_GRP1";
            ws.Cells[2, 4].Value = "PMODE_X";
            ws.Cells[2, 5].Value = "PAT_EX";

            return ws;
        }

        [TestMethod]
        public void ReadAllColumn()
        {
            // Arrange
            ExcelWorksheet worksheet = CreateTable();
            // Act
            VreMbistLookupTable? table = new VreMbistLookupTableReader().ReadSheet(worksheet);

            // Assert
            Assert.IsNotNull(table);
            Assert.AreEqual("VRE_Mbist_Lookup", table.SheetName);
            Assert.AreEqual(1, table.Rows.Count);

            OreMbistLookupRow row = table.Rows.First();
            Assert.AreEqual("HARV_A", row.Harvesting);
            Assert.AreEqual("SRV1", row.Server);
            Assert.AreEqual("MEM_GRP1", row.MemoryGroup);
            Assert.AreEqual("PMODE_X", row.Pmode);
            Assert.AreEqual("PAT_EX", row.ExcludePattern);
        }

        [TestMethod]
        public void FindHeaderIndex()
        {
            // Arrange
            ExcelWorksheet worksheet = CreateTable();
            // Act
            VreMbistLookupTable? table = new VreMbistLookupTableReader().ReadSheet(worksheet);

            // Assert
            Assert.IsNotNull(table);
            Assert.IsTrue(table.HeaderIndex.ContainsKey("Harvesting"));
            Assert.IsTrue(table.HeaderIndex.ContainsKey("Server"));
            Assert.IsTrue(table.HeaderIndex.ContainsKey("Memory Group"));
            Assert.IsTrue(table.HeaderIndex.ContainsKey("P mode"));
            Assert.IsTrue(table.HeaderIndex.ContainsKey("Exclude Pattern"));

            Assert.AreEqual(1, table.HeaderIndex["Harvesting"]);
        }

        [TestMethod]
        public void MissingOtherInformation()
        {
            // Arrange
            ExcelWorksheet ws = _package.Workbook.Worksheets.Add("VRE_Mbist_Lookup");
            ws.Cells[1, 1].Value = "Harvesting";
            ws.Cells[2, 1].Value = "HARV_ONLY";

            // Act
            VreMbistLookupTable? table = new VreMbistLookupTableReader().ReadSheet(ws);

            // Assert
            Assert.IsNotNull(table);
            Assert.AreEqual(1, table.Rows.Count);
            Assert.AreEqual("HARV_ONLY", table.Rows[0].Harvesting);
        }

        [TestMethod]
        public void MissingFirstHeader()
        {
            // Arrange
            ExcelWorksheet? ws = _package.Workbook.Worksheets.Add("VRE_Mbist_Lookup");
            ws.Cells[1, 1].Value = "WrongHeader";


            // Act
            VreMbistLookupTable? result = new VreMbistLookupTableReader().ReadSheet(ws);

            // Assert
            Assert.IsNull(result);
        }
    }

}
