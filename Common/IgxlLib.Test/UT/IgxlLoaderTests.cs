using System.Collections.Generic;
using System.IO;

using FileDiffLib;

using IgxlLib.Enums;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

using OfficeOpenXml;

namespace IgxlLib.Test.UT
{
    [TestClass]
    public class IgxlLoaderTests
    {
        public string InputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        public string OutputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        public string ExpectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void IgxlDataLoader_ExcelWorkbook_Test()
        {
            string subName = "IgxlDataLoader_ExcelWorkbook";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string file = Path.Combine(InputPath, "Golden_Spirom.xlsx");
            using var package = new ExcelPackage(new FileInfo(file));
            var types = new List<EnumSheetType> { EnumSheetType.DTFlowtableSheet, EnumSheetType.DTBintablesSheet, EnumSheetType.DTTestInstancesSheet, EnumSheetType.DTTestInstancesSheet, EnumSheetType.DTLevelSheet, EnumSheetType.DTPatternSetSheet, EnumSheetType.DTTimesetBasicSheet };
            var igxlDataLoader = new IgxlLoader(package.Workbook, types);
            foreach (IIgxlSheet sheet in igxlDataLoader.GetAllSheets())
            {
                sheet.Write(Path.Combine(outputPath, sheet.Name + ".txt"));
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void IgxlDataLoader_Igxl_Test()
        {
            string subName = "IgxlDataLoader_Igxl";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string file = Path.Combine(InputPath, "Sample.igxl");
            var types = new List<EnumSheetType> { EnumSheetType.DTFlowtableSheet, EnumSheetType.DTBintablesSheet, EnumSheetType.DTTestInstancesSheet, EnumSheetType.DTTestInstancesSheet, EnumSheetType.DTLevelSheet, EnumSheetType.DTPatternSetSheet, EnumSheetType.DTTimesetBasicSheet };
            var igxlDataLoader = new IgxlLoader(file, types);
            foreach (IIgxlSheet sheet in igxlDataLoader.GetAllSheets())
            {
                sheet.Write(Path.Combine(outputPath, sheet.Name + ".txt"));
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
