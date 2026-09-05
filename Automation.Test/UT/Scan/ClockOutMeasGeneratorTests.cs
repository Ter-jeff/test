using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.Reader.TestPlan.ClockOut;
using Automation.Static;

using CommonLib.ErrorReport;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Scan;

namespace Automation.Test.UT.Scan
{
    [TestClass]
    public class ClockOutMeasGeneratorTests : FunctionTestBase
    {
        [TestMethod]
        public void ClockOutMeasGeneratorTest()
        {
            string subName = "ClockOutMeasGenerator";
            string outputPath = Path.Combine(OutputPath, "Scan", subName);
            string expectPath = Path.Combine(ExpectPath, "Scan", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            string sheetName = "sheetName";
            BinCutFinalInstanceRows binCutInstanceRows =
            [
                new BinCutFinalInstanceRow
                {
                    BinCutInstanceRow = new BinCutInstanceRow
                    {
                        FlowName = "FlowName"
                    }
                }
            ];
            List<ClockMeasRow> clockRows =
            [
                new()
            ];
            new ClockOutMeasGenerator(sheetName, clockRows, binCutInstanceRows).WorkFlow();

            ErrorReportManager.PrintErrors(Path.Combine(outputPath, "ErrorReport.txt"));

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
