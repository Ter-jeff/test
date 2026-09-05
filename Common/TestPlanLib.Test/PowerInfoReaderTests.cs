using System.IO;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

using Newtonsoft.Json;

using OfficeOpenXml;

using TestPlanLib.DataStruct;


namespace TestPlanLib.Test
{
    [TestClass]
    public class PowerInfoReaderTests
    {
        private static readonly string _inputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        private static readonly string _outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        private static readonly string _expectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void PowerInfoReaderTest()
        {
            string subName = "PowerInfoReader";
            string outputPath = Path.Combine(_outputPath, subName);
            string expectPath = Path.Combine(_expectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string filePath = Path.Combine(_inputPath, "PowerInfo.csv");
            using var excelPackage = new ExcelPackage();
            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add(subName);
            sheet.Cells["A1"].LoadFromText(new FileInfo(filePath), new ExcelTextFormat { Delimiter = ',' });
            PowerInfoSheet result = new PowerInfoReader().ReadSheet(sheet);

            // Assert
            string json = JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
