using System.Collections.Generic;
using System.IO;

using Automation.Utility.Pattern;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.BistBira
{
    [TestClass]
    public class PatMbistInfoManagerTests : FunctionTestBase
    {
        private string _testCsvPath = null!;
        private string _testMbistInfoPath = null!;
        private PatMbistInfoManager _manager = null!;

        [TestInitialize]
        public void Setup()
        {
            _testCsvPath = Path.Combine(Path.GetTempPath(), "test_pat_list.csv");
            _testMbistInfoPath = Path.Combine(Path.GetTempPath(), "test_mbist_info.txt");
            var lines = new List<string>()
            {
                "Pattern,File Versions,USE/No Use",
                "PP_BRNA0_C_IN00_BI_E0ED_XXX_XXX_XXX_MEXXXX_SI_SRMDSSC,PP_KMDA0_L_PLLP_BI_MGP0_ERT_JTG_1RB_MGXXXX_SI_BIRA_FSTP0_1_A0_2502220533,use",
                "DD_BRNA0_C_IN00_BI_E0ED_XXX_XXX_XXX_MEXXXX_SI_SRMDSSC,DD_KMDA0_L_PLLP_BI_MGP0_ERT_JTG_1RB_MGXXXX_SI_BIRA_FSTP0_1_A0_2502220533,false",
                "AA_BRNA0_C_IN00_BI_E0ED_XXX_XXX_XXX_MEXXXX_SI_SRMDSSC,AA_KMDA0_L_PLLP_BI_MGP0_ERT_JTG_1RB_MGXXXX_SI_BIRA_FSTP0_1_A0_2502220533,use",
            };
            File.WriteAllLines(_testCsvPath, lines);
            File.WriteAllText(_testMbistInfoPath, "dummy_mbist_info");

            _manager = new PatMbistInfoManager(_testCsvPath)
            {
                FileName = "output.txt",
                OutputPath = Path.GetTempPath()
            };
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (File.Exists(_testCsvPath))
            {
                File.Delete(_testCsvPath);
            }

            if (File.Exists(_testMbistInfoPath))
            {
                File.Delete(_testMbistInfoPath);
            }
        }

        [TestMethod]
        public void GetMbistInfoFromServer_ShouldReturnTrue_WhenMbistInfoFileMissing()
        {
            // Arrange
            List<string> prodFlowSheetList = ["C_BI_PP_CP1"];
            var manager = new PatMbistInfoManager(_testCsvPath, prodFlowSheetList)
            {
                SocOffset = 1
            };
            string info = Path.Combine(KPath, "borneo", "A0_V04A", "HARD_IP", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All_borneo.txt");
            string mbistInfo = Path.Combine(KPath, "borneo", "A0_V04A", "CSV_Bist_Info", "Borneo_Bist_Info_All.csv");

            // Act
            bool result = manager.GetMbistInfoFromServer(info, mbistInfo);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void GetMbistInfoFromServer_2()
        {
            // Arrange
            List<string> prodFlowSheetList = [];
            var manager = new PatMbistInfoManager(_testCsvPath, prodFlowSheetList);
            string info = Path.Combine(KPath, "borneo", "A0_V04A", "HARD_IP", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All_borneo.txt");
            string mbistInfo = Path.Combine(KPath, "borneo", "A0_V04A", "CSV_Bist_Info", "Borneo_Bist_Info_All.csv");

            // Act
            bool result = manager.GetMbistInfoFromServer(info, mbistInfo);

            // Assert
            Assert.IsTrue(result);

        }

        [TestMethod]
        public void SaveMbistInfo_ShouldWriteFiles_WhenLMbistInfoNotEmpty()
        {
            // Arrange
            string subName = "SaveFingerPrintMaxDepth";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            _manager.LMbistInfo =
            [
                "S_BLOCK\tVector1\tCycle1\t1\tCompare\tType",
                "C_BLOCK\tVector2\tCycle2\t2\tCompare\tType"
            ];

            // Act
            _manager.SaveMbistInfo();
            _manager.SaveFingerPrintMaxDepth(Path.Combine(outputPath, "FingerPrintMaxDepth.txt"));

            // Assert
            string expectedFile = Path.Combine(_manager.OutputPath, _manager.FileName);
            string expectedModuleFile = Path.Combine(_manager.OutputPath, Path.GetFileNameWithoutExtension(_manager.FileName) + "_ModuleNameOnly" + Path.GetExtension(_manager.FileName));
            Assert.IsTrue(File.Exists(expectedFile), "Main file should exist");
            Assert.IsTrue(File.Exists(expectedModuleFile), "ModuleNameOnly file should exist");

            string[] content = File.ReadAllLines(expectedFile);
            Assert.IsTrue(content.Length > 1);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void FilterByConfig_ShouldReturnFilteredList_BasedOnConfig()
        {
            // Arrange
            var inputList = new List<string>
            {
                "PAT_S_X\tVector1\tCycle1\t1\tX\tType1",
                "PAT_C_Y\tVector2\tCycle2\t2\tO\tType2"
            };

            // Act
            List<string> result = _manager.FilterByConfig(inputList);

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.IsTrue(result[0].Contains("PAT_S_X\tVector1\tCycle1\t1\tX\tType1"));
        }
    }
}
