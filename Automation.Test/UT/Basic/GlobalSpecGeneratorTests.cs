using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.Basic.Business.GenGlobalSpec;
using Automation.Singleton;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.DataStruct;

using PinMapSheet = IgxlLib.IgxlSheets.PinMapSheet;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class GlobalSpecGeneratorTests : FunctionTestBase
    {
        private static PinMapSheet _pinMapSheet = null!;
        private static IoContiSheet _ioContiSheet = null!;
        private static IfoldPowerTableSheet _ifoldPowerTable = null!;
        private static MultiTestSettingSheetsSingleton _settingData = null!;
        private static GlobalSpecGenerator _generator = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            _settingData = MultiTestSettingSheetsSingleton.Instance();
            _pinMapSheet = new PinMapSheet("");
            _ioContiSheet = new IoContiSheet();
            _ifoldPowerTable = new IfoldPowerTableSheet("");

            _generator = new GlobalSpecGenerator(_pinMapSheet, _ioContiSheet);
        }

        [TestMethod]
        public void GetPowerGlobalSpecs_ShouldGenerateExpectedGlobals_ForValidPowerPin()
        {
            // Arrange

            var powerSheet = new PowerInfoSheet();
            powerSheet.Rows.Add(new PowerInfoRow
            {
                PinName = "VDD1",
                SheetName = "UnitTestSheet",
                RowNum = 10,
                PowerSequence = "t1:ON;t2:OFF",
                PowerDownSequence = "t3:DOWN",
                Ifold = "IFold_1:5"
            });

            // Act
            List<GlobalSpec> specs = _generator.GetPowerGlobalSpecs(powerSheet, _settingData, _ifoldPowerTable);

            // Assert
            Assert.AreNotEqual(null, specs, "Specs list should not be null");
            Assert.IsTrue(specs.Count > 0, "At least one GlobalSpec should be created");

            Assert.IsTrue(specs.Exists(s => s.Symbol.Contains("VDD1_PowerSequence_GLB")), "Missing PowerSequence_GLB symbol");
            Assert.IsTrue(specs.Exists(s => s.Symbol.Contains("VDD1_PowerDownSequence_GLB")), "Missing PowerDownSequence_GLB symbol");
        }

        [TestMethod]
        public void GetPowerGlobalSpecs_ShouldSkipNwirePins()
        {
            // Arrange
            var powerSheet = new PowerInfoSheet();
            powerSheet.Rows.Add(new PowerInfoRow
            {
                PinName = "NWIRE_VDD",
                SheetName = "UnitTestSheet",
                RowNum = 12,
                PowerSequence = "t1:ON",
                Ifold = "IFold_1:5",
                PowerDownSequence = ""
            });

            // Act
            List<GlobalSpec> specs = _generator.GetPowerGlobalSpecs(powerSheet, _settingData, _ifoldPowerTable);

            // Assert
            Assert.AreEqual(2, specs.Count, "Nwire pins should not generate any global specs");
        }

        [TestMethod]
        public void GetPowerGlobalSpecs_ShouldHandleDcviAndDcvsPins()
        {
            // Arrange
            var powerSheet = new PowerInfoSheet();
            powerSheet.Rows.Add(new PowerInfoRow
            {
                PinName = "VDDX",
                SheetName = "UnitTestSheet",
                RowNum = 20,
                PowerSequence = "t1:ON",
                PowerDownSequence = "t2:OFF",
                Ifold = "IFold_1:5"
            });

            // Act
            List<GlobalSpec> specs = _generator.GetPowerGlobalSpecs(powerSheet, _settingData, _ifoldPowerTable);

            // Assert
            Assert.IsTrue(specs.Exists(s => s.Symbol.Contains("VDDX_PowerSequence_GLB")), "Expected DCVI sequence global spec");
            Assert.IsTrue(specs.Exists(s => s.Symbol.Contains("VDDX_PowerDownSequence_GLB")), "Expected DCVS sequence global spec");
        }

        [TestMethod]
        public void GetIoGlobalSpecsTest()
        {
            string subName = "GetIoGlobalSpecs";
            string outputPath = Path.Combine(OutputPath, "Htol", subName);
            string expectPath = Path.Combine(ExpectPath, "Htol", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            (IoInfoSheet printSheet, IoInfoSheet printConcurrentSheet) = GenIoInfoSheets();
            List<GlobalSpec> rows = _generator.GetIoGlobalSpecs(printSheet, printConcurrentSheet);

            var sheet = new GlobalSpecSheet("GlobalSpecSheet");
            sheet.Rows.AddRange(rows);
            sheet.Write(Path.Combine(outputPath, sheet.Name + ".txt"));

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void GetPowerSeqValue_LettersFollowedByDigits_ReturnsDigitsOnly()
        {
            // Act
            string result = _generator.GetPowerSeqValue("t123");

            // Assert
            Assert.AreEqual("123", result);
        }

        [TestMethod]
        public void GetPowerSeqValue_EmptyString_ReturnsEmptyString()
        {
            // Act
            string result = _generator.GetPowerSeqValue("");

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetPowerSeqValue_NoDigits_ReturnsNinetyNine()
        {
            // Act
            string result = _generator.GetPowerSeqValue("###");

            // Assert
            Assert.AreEqual("99", result);
        }

        [TestMethod]
        public void GetIfold_MatchingIfoldPowerTableRow_ReturnsTableCurrentValue()
        {
            // Arrange
            var powerInfoRow = new PowerInfoRow { PinName = "VDD1", Ifold = "5" };
            var ifoldTable = new IfoldPowerTableSheet("");
            ifoldTable.Rows.Add(new IfoldPowerTableRow { PinName = "VDD1", Current = "42" });

            // Act
            string result = _generator.GetIfold(powerInfoRow, ifoldTable);

            // Assert
            Assert.AreEqual("42", result);
        }

        [TestMethod]
        public void GetIfold_NoTableMatchAndPlainValue_ReturnsValueUnchanged()
        {
            // Arrange
            var powerInfoRow = new PowerInfoRow { PinName = "VDD1", Ifold = "5" };
            var ifoldTable = new IfoldPowerTableSheet("");

            // Act
            string result = _generator.GetIfold(powerInfoRow, ifoldTable);

            // Assert
            Assert.AreEqual("5", result);
        }

        [TestMethod]
        public void GetIfold_ColonWithEmptyJobPrefix_PrependsZero()
        {
            // Arrange
            var powerInfoRow = new PowerInfoRow { PinName = "VDD1", Ifold = ":5" };
            var ifoldTable = new IfoldPowerTableSheet("");

            // Act
            string result = _generator.GetIfold(powerInfoRow, ifoldTable);

            // Assert
            Assert.AreEqual("0:5", result);
        }

        [TestMethod]
        public void GetIfold_ColonWithNonEmptyJobPrefix_ReturnsUnchanged()
        {
            // Arrange
            var powerInfoRow = new PowerInfoRow { PinName = "VDD1", Ifold = "IFold_1:5" };
            var ifoldTable = new IfoldPowerTableSheet("");

            // Act
            string result = _generator.GetIfold(powerInfoRow, ifoldTable);

            // Assert
            Assert.AreEqual("IFold_1:5", result);
        }

        [TestMethod]
        public void GetIfold_EmptySegment_DefaultsToZero()
        {
            // Arrange
            var powerInfoRow = new PowerInfoRow { PinName = "VDD1", Ifold = "" };
            var ifoldTable = new IfoldPowerTableSheet("");

            // Act
            string result = _generator.GetIfold(powerInfoRow, ifoldTable);

            // Assert
            Assert.AreEqual("0", result);
        }
    }
}
