using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.PostAction.GenVreTestCase;
using Automation.Static;
using Automation.Test.Static;

using CommonLib.Extension;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using TestPlanLib;
using TestPlanLib.Harvest;
using TestPlanLib.Singleton;
using TestPlanLib.VbtLib;

namespace Automation.Test.UT.VRE
{
    [TestClass]
    public class VreMainTests : FunctionTestBase
    {
        private static List<Function> _functions = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            _functions = TestSuiteInitialize.Functions;
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void VreMainTest()
        {
            string subName = "VreTestCaseMain";
            string outputPath = Path.Combine(OutputPath, "Vre", subName);
            string expectPath = Path.Combine(ExpectPath, "Vre", subName);
            string binNumSheetPath = Path.Combine(InputPath, "borneo_documents", "A0_V04A", "TestPlan", "SOC_Xlsx", "BinNum.xlsx");
            string vreTestCaseTablePath = Path.Combine(InputPath, "borneo_documents", "A0_V04A", "TestPlan", "VRE", "VRE_Test_Scenarios.csv");
            string mbistLookupTablePath = Path.Combine(InputPath, "borneo_documents", "A0_V04A", "TestPlan", "VRE", "VRE_Mbist_Lookup.csv");
            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            if (Directory.Exists(FolderStructure.DirIgLink))
            {
                Directory.Delete(FolderStructure.DirIgLink, true);
            }
            _ = Directory.CreateDirectory(FolderStructure.DirIgLink);
            TestProgram.Clear();
            TestProgram.VbtFunctionLib.AddVbtFunctionRange([.. _functions.Where(x => x.Type.EqualsIgnoreCase(".NET"))]);
            BlockStatus.GetAutomationBlockStatus(BlockConst.BinCut).Down = true;

            ExcelWorkbook binNumSheet = new ExcelPackage(new FileInfo(binNumSheetPath)).Workbook;
            LocalSpecs.Options.VreEnable = true;
            BinNumberSingleton.Instance.Initialize(binNumSheet);
            VreTestCaseTable? vreTestCaseTable = new VreTestCaseTableReader().ReadSheet(ConvertCsvToExcelSheet(vreTestCaseTablePath));
            VreMbistLookupTable? mbistLookupTable = new VreMbistLookupTableReader().ReadSheet(ConvertCsvToExcelSheet(mbistLookupTablePath));
            HarvestingTruthTableSheet? harvestingTruthTable = TestPlanStatic.HarvestingTruthTableSheets.FirstOrDefault();
            MappingCoreTable? mappingDigCoreTable = new MappingCoreTable("MappingDigitalCore");
            mappingDigCoreTable.Rows.Add(new MappingCoreRow { InitPattern = "", Pattern = "*SC_E0ED_SAA_COM_HAR_MEXXXX_DM*", CoreName = "pg_ECPU_Main", HarvestFlag = "F_ECPU_SA_CORE0" });
            mappingDigCoreTable.Rows.Add(new MappingCoreRow { InitPattern = "", Pattern = "*SC_E0ED_TDF_COM_HAR_MEXXXX_DM_DIV166*", CoreName = "pg_ECPU_Main", HarvestFlag = "F_ECPU_SA_CORE0" });
            mappingDigCoreTable.Rows.Add(new MappingCoreRow { InitPattern = "", Pattern = "*SC_CCD0_SAA_SSC_HAR_MSXXXX_DM*", CoreName = "AISBES_inst_HARV_ISP", HarvestFlag = "F_SOC_SA_IS" });
            if (LocalSpecs.Options.VreEnable && vreTestCaseTable != null && mbistLookupTable != null && harvestingTruthTable != null)
            {
                new GenVreTestCaseMain(vreTestCaseTable,
                                                mappingDigCoreTable,
                                                harvestingTruthTable,
                                                mbistLookupTable,
                                                TestPlanStatic.FlagOperationSheets).Workflow();
            }
            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
        public static ExcelWorksheet ConvertCsvToExcelSheet(string fileName)
        {
            var excelPackage = new ExcelPackage();
            string sheetName = Path.GetFileNameWithoutExtension(fileName);
            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
            int index = 0;
            using (var sr = new StreamReader(fileName))
            {
                if (sr != null)
                {
                    while (!sr.EndOfStream)
                    {
                        index++;
                        string? line = sr.ReadLine();
                        if (line != null)
                        {
                            string[] arr = line.Split(new[] { ',' }, StringSplitOptions.None);
                            int cnt = 0;
                            foreach (string item in arr)
                            {
                                sheet.Cells[index, 1 + cnt].Value = item;
                                cnt++;
                            }
                        }
                    }
                }
            }
            return sheet;
        }
    }
}
