using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.Basic.Business.GenMappingTable;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Extension;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using TestPlanLib.BinCut;
using TestPlanLib.DataStruct;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class NewDsscMappingTableTests : FunctionTestBase
    {
        [TestMethod]
        public void NewDsscMappingTableTest()
        {
            string subName = "NewDsscMappingTable";
            string outputPath = Path.Combine(OutputPath, "Basic", subName);
            string expectPath = Path.Combine(ExpectPath, "Basic", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;

            // Act
            string csvPath = Path.Combine(InputPath, "borneo_documents", "A0_V04A", "TestPlan", "SELSRAM", "SELSRM_Mapping_Table.csv");
            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("SELSRM_Mapping_Table");
                _ = sheet.Cells[1, 1].PrintExcelRowByList(csvPath.CsvConvertToLists());
                var dsscMappingTable = new NewDsscMappingTable();
                dsscMappingTable.Workflow(sheet, outputPath, MultiTestSettingSheetsSingleton.Instance());
            }

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        private static SelsrmMappingTableRow NewRow(string block, string stage, string logicPins = "P1", string sramPins = "P2")
        {
            return new SelsrmMappingTableRow { Block = block, Stage = stage, LogicPins = logicPins, SramPins = sramPins };
        }

        [TestMethod]
        public void GetSelsrmMappingTableRows_RtosBlockWithMatchingStage_IsIncluded()
        {
            // Arrange
            var sheet = new SelsrmMappingSheet("Sheet1") { Rows = [NewRow("RTOS", "CP1,FT1")] };

            // Act
            List<SelsrmMappingTableRow> result = NewDsscMappingTable.GetSelsrmMappingTableRows(sheet, "CP1");

            // Assert
            Assert.AreEqual(1, result.Count);
        }

        [TestMethod]
        public void GetSelsrmMappingTableRows_NonRtosNonWildcardBlock_IsExcluded()
        {
            // Arrange
            var sheet = new SelsrmMappingSheet("Sheet1") { Rows = [NewRow("ATE", "CP1")] };

            // Act
            List<SelsrmMappingTableRow> result = NewDsscMappingTable.GetSelsrmMappingTableRows(sheet, "CP1");

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetSelsrmMappingTableRows_MissingLogicOrSramPins_IsExcluded()
        {
            // Arrange
            var sheet = new SelsrmMappingSheet("Sheet1") { Rows = [NewRow("RTOS", "CP1", logicPins: "")] };

            // Act
            List<SelsrmMappingTableRow> result = NewDsscMappingTable.GetSelsrmMappingTableRows(sheet, "CP1");

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetSelsrmMappingTableRows_WildcardStage_MatchesPrefixedJob()
        {
            // Arrange
            var sheet = new SelsrmMappingSheet("Sheet1") { Rows = [NewRow("*", "CP*")] };

            // Act
            List<SelsrmMappingTableRow> result = NewDsscMappingTable.GetSelsrmMappingTableRows(sheet, "CP2");

            // Assert
            Assert.AreEqual(1, result.Count);
        }

        [TestMethod]
        public void GetSelsrmMappingTableRows_JobDoesNotMatchStage_IsExcluded()
        {
            // Arrange
            var sheet = new SelsrmMappingSheet("Sheet1") { Rows = [NewRow("RTOS", "FT1")] };

            // Act
            List<SelsrmMappingTableRow> result = NewDsscMappingTable.GetSelsrmMappingTableRows(sheet, "CP1");

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [DataTestMethod]
        [DataRow("BI", "C", "Mbist_Cpu_Test", DisplayName = "BiCpu")]
        [DataRow("BI", "L", "Mbist_Gfx_Test", DisplayName = "BiGfx")]
        [DataRow("BI", "S", "Mbist_Soc_Test", DisplayName = "BiSoc")]
        [DataRow("SC", "C", "Sa_Cpu_Test", DisplayName = "ScSaCpu")]
        [DataRow("SC", "L", "Sa_Gpu_Test", DisplayName = "ScSaGpu")]
        [DataRow("SC", "S", "Td_Soc_Test", DisplayName = "ScTdSoc")]
        public void GetDcCategoryNames_MatchingBlockAndDomain_ReturnsFilteredCategories(string block, string domain, string categoryName)
        {
            // Arrange
            var testSettingSheet = new TestSettingData
            {
                DcCategorys = [new DcCategoryName(categoryName), new DcCategoryName("Unrelated_Category")]
            };

            // Act
            List<DcCategoryName> result = NewDsscMappingTable.GetDcCategoryNames(testSettingSheet, block, domain);

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(categoryName, result[0].CategoryName);
        }

        [TestMethod]
        public void GetDcCategoryNames_UnknownBlock_ReturnsNull()
        {
            // Arrange
            var testSettingSheet = new TestSettingData
            {
                DcCategorys = [new DcCategoryName("Mbist_Cpu_Test")]
            };

            // Act
            List<DcCategoryName>? result = NewDsscMappingTable.GetDcCategoryNames(testSettingSheet, "XX", "C");

            // Assert
            Assert.IsNull(result);
        }
    }
}
