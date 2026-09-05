using System;
using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.PostAction.TempMon;
using Automation.GenerateIgxl.PostAction.TempMon.Data;
using Automation.GenerateIgxl.PostAction.TempMon.Enums;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using TestPlanLib.NonIgxlSheets;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class TempMonMainTests : FunctionTestBase
    {
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void WorkFlow_Should_Generate_Output_Correctly()
        {
            SetUpTempMonConfigSheet((tempMonConfigSheet) =>
            {
                // Arrange
                string subName = "TempMon";
                string outputPath = Path.Combine(OutputPath, "PostAction", subName);
                string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

                if (Directory.Exists(outputPath))
                {
                    Directory.Delete(outputPath, true);
                }
                _ = Directory.CreateDirectory(outputPath);

                var datas = new HashSet<TempMonData>()
                {
                new(){ Condition = EnumCondition.Include, Item = "123", Mode = "TEST", Type = EnumType.Instance},
                new(){ Condition = EnumCondition.Include, Item = "456", Mode = "TEST", Type = EnumType.Flow}
                };
                var nonIgxlSheets = new NonIgxlSheets();

                var target = new TempMonMain(tempMonConfigSheet, datas, nonIgxlSheets, outputPath);

                // Act
                target.WorkFlow();

                // Assert
                bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
                if (fail)
                {
                    Assert.Fail("Unit Test Fail!!!");
                }
            });
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void WorkFlow_Should_Do_Nothing()
        {
            // Arrange
            string subName = "TempMonNull";
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var datas = new HashSet<TempMonData>();
            var nonIgxlSheets = new NonIgxlSheets();

            var target = new TempMonMain(null, datas, nonIgxlSheets, outputPath);

            // Act
            target.WorkFlow();

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        // Helper methods

        // Provides a controlled test environment:
        // 1. Creates ExcelPackage with proper disposal
        // 2. Initializes a default ConfigSheet
        private static void SetUpTempMonConfigSheet(Action<ExcelWorksheet> excelWorksheets)
        {
            using var package = new ExcelPackage();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("TempMon_configuration");
            BuildDefaultConfigSheet(ws);
            excelWorksheets(ws);
        }

        // Builds a shared ConfigSheet used across tests
        private static void BuildDefaultConfigSheet(ExcelWorksheet excelWorksheet)
        {
            excelWorksheet.Cells[1, 1].Value = "Item";
            excelWorksheet.Cells[1, 2].Value = "Value";
            excelWorksheet.Cells[1, 2].Value = "Column";

            excelWorksheet.Cells[2, 1].Value = "CmnIDLength";
            excelWorksheet.Cells[2, 2].Value = "7";

            excelWorksheet.Cells[3, 1].Value = "TmpsA0";
            excelWorksheet.Cells[3, 2].Value = "65.5";
        }
    }
}
