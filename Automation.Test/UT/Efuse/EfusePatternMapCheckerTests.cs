using System.Collections.Generic;

using Automation.GenerateIgxl.EFuse.Business;
using Automation.GenerateIgxl.EFuse.InputChecker;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using ScghLib.Reader;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Efuse.Input;

namespace Automation.Test.UT.Efuse
{
    [TestClass]
    public class EfusePatternMapCheckerTests
    {
        private EfusePatternMapChecker _checker = null!;

        [TestInitialize]
        public void Setup()
        {
            ErrorReportManager.ClearErrors();

            var patternDic = new Dictionary<string, PatternData>
            {
                ["valid_pattern"] = new PatternData { Use = "use", FileVersion = "1.0" },
                ["dontusepattern"] = new PatternData { Use = "dont_use", FileVersion = "1.0" },
                ["naversionpattern"] = new PatternData { Use = "use", FileVersion = "n/a" },
                ["goodpattern"] = new PatternData { Use = "use", FileVersion = "1.0" }
            };

            var patternRow = new EfusePatternRow
            {
                PatternType = new EfusePatternType { TestMode = EfuseTestMode.WriteDssc },
                ReadOrWrite = EfusePatternWorR.Write,
                ReadWritePin = "PIN_A",
                BankName = "BANK1",
                SendStoreCount = 8
            };

            var bankRanges = new List<BitDefBankRange>
            {
                new() { BankName = "BANK1", Width = 8 }
            };

            _checker = new EfusePatternMapChecker(patternDic, patternRow, bankRanges);
        }

        [TestMethod]
        public void CheckPatternInScgh_ShouldAddError_WhenPatternNotInCsv()
        {
            var hardIpRow = new ProdCharSheetRow("Sheet1")
            {
                PayloadList = ["not_exist_pattern"],
                RowNum = 10
            };

            _checker.CheckScghPatterns(hardIpRow, "AccessMode", "Mode1");

            Assert.AreEqual("The pattern not_exist_pattern can't be found in the CSV !!!", ErrorReportManager.GetErrorList()[0].Message);
        }

        [TestMethod]
        public void CheckPatternInScgh_ShouldAddWarning_WhenPatternIsDontUse()
        {
            var hardIpRow = new ProdCharSheetRow("Sheet1")
            {
                PayloadList = ["dont_use_pattern"],
                RowNum = 12
            };

            _checker.CheckScghPatterns(hardIpRow, "AccessMode", "Mode1");

            Assert.AreEqual("The pattern dont_use_pattern can't be found in the CSV !!!", ErrorReportManager.GetErrorList()[0].Message);
        }

        [TestMethod]
        public void CheckPatternInScgh_ShouldAddError_WhenPatternVersionNA()
        {
            var hardIpRow = new ProdCharSheetRow("Sheet2")
            {
                PayloadList = ["na_pattern"],
                RowNum = 15
            };

            _checker.CheckScghPatterns(hardIpRow, "AccessMode", "Mode1");

            Assert.AreEqual("The pattern na_pattern can't be found in the CSV !!!", ErrorReportManager.GetErrorList()[0].Message);
        }

        [TestMethod]
        public void CheckWritePatternSendCount_ShouldAddError_WhenCountNotMatch()
        {
            var patternDic = new Dictionary<string, PatternData>
            {
                ["valid_pattern"] = new PatternData { Use = "use", FileVersion = "1.0" }
            };

            var patternRow = new EfusePatternRow
            {
                PatternType = new EfusePatternType { TestMode = EfuseTestMode.WriteDssc },
                ReadOrWrite = EfusePatternWorR.Write,
                BankName = "BANK1",
                SendStoreCount = 5
            };

            var bankRanges = new List<BitDefBankRange>
            {
                new() { BankName = "BANK1", Width = 8 }
            };

            var checker = new EfusePatternMapChecker(patternDic, patternRow, bankRanges);
            checker.CheckScghPatterns(
                new ProdCharSheetRow("Sheet1")
                {
                    PayloadList = ["valid_pattern"],
                    RowNum = 1
                },
                "AccessMode",
                "Mode1"
            );

            Assert.AreEqual("The pattern valid_pattern send count 5 5 is not equal to 8 in the bank BANK1, 8*1 !!!!!", ErrorReportManager.GetErrorList()[0].Message);
        }

        [TestMethod]
        public void CheckReadPatternStoreCount_ShouldHandleDAAMismatch()
        {
            var patternDic = new Dictionary<string, PatternData>
            {
                ["valid_pattern"] = new PatternData { Use = "use", FileVersion = "1.0" }
            };

            var patternRow = new EfusePatternRow
            {
                PatternType = new EfusePatternType { TestMode = EfuseTestMode.JtagRead },
                ReadOrWrite = EfusePatternWorR.Read,
                BankName = "BANK1",
                SendStoreCount = 1
            };

            var bankRanges = new List<BitDefBankRange>
            {
                new() { BankName = "BANK1", Width = 8 }
            };

            var checker = new EfusePatternMapChecker(patternDic, patternRow, bankRanges);
            checker.CheckScghPatterns(
                new ProdCharSheetRow("Sheet3")
                {
                    PayloadList = ["aaa_bbb_ccc_ddd_eee_fff_ggg_daa"],
                    RowNum = 5
                },
                "JTAG-Access Mode",
                "Mode2"
            );

            Assert.AreEqual("The pattern aaa_bbb_ccc_ddd_eee_fff_ggg_daa (DAA) store count 1 is not equal to 16 in the bank BANK1 (JTAG-Access Mode), 8*2 !!!", ErrorReportManager.GetErrorList()[0].Message);
        }

        [DataTestMethod]
        [DataRow("UnknownPattern", null, "The pattern UnknownPattern can't be found in the CSV !!!", DisplayName = "01_Payload unknown")]
        [DataRow("goodpattern", "dontusepattern", "This pattern dontusepattern is \"Dont_useInCsv\" !!!", DisplayName = "02_Init dont_use")]
        [DataRow("goodpattern", "naversionpattern", "This pattern naversionpattern of the FileVersion column is \"n/a\" !!!", DisplayName = "03_Init FileVersion n/a")]
        public void CheckPatternInInstance_DataRow_ShouldReportErrors(string payload, string initList, string expectedMessage)
        {
            // Arrange
            var instanceRow = new BinCutInstanceRow
            {
                SheetName = "Sheet1",
                RowNum = 1,
                PayloadList = [payload],
                InitList = initList != null ? [initList] : []
            };

            // Act
            _checker.CheckPatternInInstance(instanceRow);

            List<Error> errors = ErrorReportManager.GetErrorList();

            // Assert
            Assert.IsTrue(errors.Count > 0, "Expected at least one error");
            Assert.AreEqual(expectedMessage, errors[0].Message);

            // Cleanup
            ErrorReportManager.ClearErrors();
        }
    }
}
