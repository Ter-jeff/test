using System.IO;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

using Newtonsoft.Json;

using OfficeOpenXml;

using ScghLib.Base;
using ScghLib.Reader;

namespace ScghLib.Test
{
    [TestClass]
    public class ScghLibTests
    {
        public static readonly string InputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        public static readonly string OutputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        public static readonly string ExpectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void BistProdFlowReaderTest()
        {
            string subName = "BistProdFlowReader";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string file = Path.Combine(InputPath, "borneo_A0_SCGH_X_X_X#208.xlsx");
            using (var package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["C_BI_PP_CP1"];
                BistProdFlowSheet result = new BistProdFlowReader(new MbistSheet { SheetName = worksheet.Name }).ReadSheet(worksheet);
                string json = JsonConvert.SerializeObject(result, Formatting.Indented);
                File.WriteAllText(Path.Combine(outputPath, "result.json"), json);
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void HardIpProdCharReaderTest()
        {
            string subName = "HardIpProdCharReader";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            _ = Directory.CreateDirectory(outputPath);

            string file = Path.Combine(InputPath, "borneo_A0_SCGH_X_X_X#208.xlsx");
            using (var package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["HardIP_PC"];
                HardIpScghSheet result = new HardIpProdCharReader().ReadSheet(worksheet);
                string json = JsonConvert.SerializeObject(result, Formatting.Indented);
                File.WriteAllText(Path.Combine(outputPath, "result.json"), json);
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void ProdCharSheetReaderTest()
        {
            string subName = "ProdCharSheetReader";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            _ = Directory.CreateDirectory(outputPath);

            string file = Path.Combine(InputPath, "borneo_A0_SCGH_X_X_X#208.xlsx");
            using (var package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["C_BI_PC"];
                ProdCharSheet result = new ProdCharSheetReader().ReadSheet(worksheet);
                string json = JsonConvert.SerializeObject(result, Formatting.Indented);
                File.WriteAllText(Path.Combine(outputPath, "result.json"), json);
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void RtosProdCharReaderTest()
        {
            string subName = "RtosProdCharReader";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            _ = Directory.CreateDirectory(outputPath);

            string file = Path.Combine(InputPath, "borneo_A0_SCGH_X_X_X#208.xlsx");
            using (var package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["RTOS_PC"];
                RtosProdCharSheet result = new RtosProdCharReader(null).ReadSheet(worksheet);
                string json = JsonConvert.SerializeObject(result, Formatting.Indented);
                File.WriteAllText(Path.Combine(outputPath, "result.json"), json);
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
