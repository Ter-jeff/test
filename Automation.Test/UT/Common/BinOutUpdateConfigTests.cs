using System.IO;

using Automation.Utility.TpUpdate.HardIPBinoutTPUpdate;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class BinOutUpdateConfigTests : FunctionTestBase
    {
        private const string BinOutUpdateConfigFileName = "BinOutUpdate_Config";
        private const string DocumentFolder = "andros_documents";

        [TestMethod]
        [DataRow("Andros_A0_BinOut_V04A_X_HardIP_X_Setting_True.xlsx", true, DisplayName = "BinOutConfig_OverWriteTestPlanLimits_TRUE_Case")]
        [DataRow("Andros_A0_BinOut_V04A_X_HardIP_X_Setting_False.xlsx", false, DisplayName = "BinOutConfig_OverWriteTestPlanLimits_FALSE_Case")]
        [DataRow("Andros_A0_BinOut_V04A_X_HardIP_X_Setting_Other.xlsx", false, DisplayName = "BinOutConfig_OverWriteTestPlanLimits_OTHER_Case")]
        public void BinOutConfigTest(string reportName, bool expectedSetting)
        {
            string reportPath = Path.Combine(InputPath, DocumentFolder, reportName);
            using var excelPackage = new ExcelPackage(new FileInfo(reportPath));
            ExcelWorksheet configSheet = excelPackage.Workbook.Worksheets[BinOutUpdateConfigFileName];
            BinOutUpdateConfigSheet binoutUpdateConfig = new BinOutUpdateConfigReader().ReadSheet(configSheet);
            Assert.AreNotEqual(null, binoutUpdateConfig);
            Assert.AreEqual(1, binoutUpdateConfig.IndexOption);
            Assert.AreEqual(2, binoutUpdateConfig.IndexValue);
            Assert.AreEqual(expectedSetting, binoutUpdateConfig.EnableOverWriteTestLimits);
        }

        [TestMethod]
        [DataRow("Andros_A0_BinOut_V04A_X_HardIP_X_Setting_Sheet_Not_Exists.xlsx", DisplayName = "BinOutConfig_SheetNotExists_Case")]
        public void BinOutConfigTestNotExists(string reportName)
        {
            string reportPath = Path.Combine(InputPath, DocumentFolder, reportName);
            using var excelPackage = new ExcelPackage(new FileInfo(reportPath));
            ExcelWorksheet configSheet = excelPackage.Workbook.Worksheets[BinOutUpdateConfigFileName];
            Assert.AreEqual(null, configSheet);
        }
    }
}
