using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.Static;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

using OfficeOpenXml;

using ScghLib.Reader;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class ScghDataTests : FunctionTestBase
    {
        private static ScghData _scghData = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            _scghData = new ScghData();
            List<HardIpScghSheet> hardIpScghSheets = new ScghData().GetEfuseFromHardIpScghSheets(EpWorkbook.ScghWorkbook);
            foreach (HardIpScghSheet hardIpScghSheet in hardIpScghSheets)
            {
                _scghData.AddHardipSheet(hardIpScghSheet);
            }
        }

        [TestMethod]
        public void GetProdCharSheetRowTest()
        {
            string subName = "GetProdCharSheetRow";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var results = new List<ProdCharSheetRow>();
            List<string> patterns = ["pp_kmda0_s_fulp_ef_ecc0_cfg_daa_sns_allfrv_si_all_fldssc"];
            results.Add(_scghData.GetProdCharSheetRow(patterns));
            string pattern = "pp_kmda0_s_fulp_ef_ecc0_cfg_daa_prg_allfrv_si_ecc_init;pp_kmda0_s_fulp_ef_ecc0_cfg_daa_prg_allfrv_si_all_eccdssc";
            results.Add(_scghData.GetProdCharSheetRow(pattern));

            // Assert
            string json = JsonConvert.SerializeObject(results, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void GetPerformanceMode_ShouldReturnEmpty_WhenNoPatterns()
        {
            // Arrange
            using var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("EmptySheet");

            // Act
            List<string> modes = ScghData.GetPerformanceMode(package.Workbook, true);

            // Assert
            Assert.AreNotEqual(null, modes);
            Assert.AreEqual(0, modes.Count);
        }

        [TestMethod]
        public void GetPerformanceMode_ShouldBeCaseInsensitive()
        {
            // Act
            List<string> modes = ScghData.GetPerformanceMode(EpWorkbook.ScghWorkbook, true);

            // Assert
            Assert.AreNotEqual(null, modes);
            Assert.AreEqual(76, modes.Count);
        }

        [TestMethod]
        public void GetOtpScghRows_ValidRow_ReturnsRow()
        {
            // Arrange
            var list = new List<ProdCharSheetRow>
            {
                new()
                {
                    PayloadList = ["A_B_C_D_E_MYOTP"],
                    Usage = "1",
                    Block = "OTP_BLOCK"
                }
            };

            // Act
            List<ProdCharSheetRow> result = ScghData.GetOtpScghRows(list);

            // Assert
            Assert.AreEqual(1, result.Count);
        }
    }
}
