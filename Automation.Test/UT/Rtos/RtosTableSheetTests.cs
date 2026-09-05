using Automation.Reader.ConfigFile.RtosTable;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Rtos
{
    [TestClass]
    public class RtosTableSheetTests
    {
        [TestMethod]
        public void LoadConfig_FullSheet_ParsesAllFiveSections()
        {
            // Arrange - one header row across non-overlapping columns, one shared data row,
            // and a blank row-2 col-1 terminator (all sections use column 1 as their stop check)
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Rtos");
            sheet.Cells[1, 1].Value = "SheetName";
            sheet.Cells[1, 2].Value = "FuncName";
            sheet.Cells[1, 3].Value = "ArgName";
            sheet.Cells[1, 4].Value = "ArgInfo";
            sheet.Cells[1, 5].Value = "PinName";
            sheet.Cells[1, 6].Value = "PinValue";
            sheet.Cells[1, 7].Value = "BinaryFileName";
            sheet.Cells[1, 8].Value = "RtosLevel";
            sheet.Cells[1, 9].Value = "Parameter";
            sheet.Cells[1, 10].Value = "Value";
            sheet.Cells[1, 11].Value = "MeasMode";
            sheet.Cells[1, 12].Value = "MeasPattern";

            sheet.Cells[2, 1].Value = "Sheet1";
            sheet.Cells[2, 2].Value = "Func1";
            sheet.Cells[2, 3].Value = "Arg1";
            sheet.Cells[2, 4].Value = "Info1";
            sheet.Cells[2, 5].Value = "PinA";
            sheet.Cells[2, 6].Value = "1.0";
            sheet.Cells[2, 7].Value = "binary.bin";
            sheet.Cells[2, 8].Value = "PinX";
            sheet.Cells[2, 9].Value = "ParamA";
            sheet.Cells[2, 10].Value = "ValueA";
            sheet.Cells[2, 11].Value = "ModeA";
            sheet.Cells[2, 12].Value = "PatternA";

            // Act
            RtosTableSheet result = RtosTableSheet.LoadConfig(sheet);

            // Assert
            Assert.AreEqual(1, result.ArgRows.Count);
            Assert.AreEqual("Sheet1", result.ArgRows[0].SheetName);
            Assert.AreEqual("Func1", result.ArgRows[0].FuncName);
            Assert.AreEqual("Arg1", result.ArgRows[0].ArgName);
            Assert.AreEqual("Info1", result.ArgRows[0].ArgInfo);

            Assert.AreEqual("1.0", result.PinRows["PinA"]);
            Assert.AreEqual("binary.bin", result.BinFileName);

            Assert.AreEqual(1, result.LevelRows.Count);
            Assert.AreEqual("PinX", result.LevelRows[0].PinName);
            Assert.AreEqual("ValueA", result.LevelRows[0].Parameters["ParamA"]);

            Assert.AreEqual("PatternA", result.MeasRows["ModeA"]);
        }

        [TestMethod]
        public void LoadConfig_NoRecognizedHeaders_ReturnsAllEmptyWithoutThrowing()
        {
            // Arrange - a non-empty sheet (so Dimension is not null) with no matching headers
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Rtos");
            sheet.Cells[1, 1].Value = "Unrelated";

            // Act
            RtosTableSheet result = RtosTableSheet.LoadConfig(sheet);

            // Assert
            Assert.AreEqual(0, result.ArgRows.Count);
            Assert.AreEqual(0, result.PinRows.Count);
            Assert.AreEqual("", result.BinFileName);
            Assert.AreEqual(0, result.LevelRows.Count);
            Assert.AreEqual(0, result.MeasRows.Count);
        }

        [TestMethod]
        public void LoadConfig_EmptyWorksheetWithNoDimension_ThrowsWrappedException()
        {
            // Arrange - a worksheet with zero cells set has a null Dimension, which the
            // header-scanning loop dereferences directly; LoadConfig wraps the failure
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Rtos");

            // Act & Assert
            System.Exception ex = Assert.ThrowsException<System.Exception>(() => RtosTableSheet.LoadConfig(sheet));
            StringAssert.Contains(ex.Message, "Error occurs during load Rtos Table Config");
        }

        [TestMethod]
        public void LoadConfig_MultipleArgRows_ReadsUntilBlankRow()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Rtos");
            sheet.Cells[1, 1].Value = "SheetName";
            sheet.Cells[1, 2].Value = "FuncName";
            sheet.Cells[1, 3].Value = "ArgName";
            sheet.Cells[1, 4].Value = "ArgInfo";

            sheet.Cells[2, 1].Value = "Sheet1";
            sheet.Cells[2, 2].Value = "Func1";
            sheet.Cells[2, 3].Value = "Arg1";
            sheet.Cells[2, 4].Value = "Info1";

            sheet.Cells[3, 1].Value = "Sheet2";
            sheet.Cells[3, 2].Value = "Func2";
            sheet.Cells[3, 3].Value = "Arg2";
            sheet.Cells[3, 4].Value = "Info2";

            // Act
            RtosTableSheet result = RtosTableSheet.LoadConfig(sheet);

            // Assert
            Assert.AreEqual(2, result.ArgRows.Count);
            Assert.AreEqual("Sheet2", result.ArgRows[1].SheetName);
        }

        [TestMethod]
        public void LoadConfig_LevelRowsWithSamePinName_MergeParametersIntoSingleRow()
        {
            // Arrange - two data rows share the same RtosLevel pin name; the second row's
            // parameter/value should be merged into the existing RtosLevelRow, not create a new one
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Rtos");
            sheet.Cells[1, 1].Value = "RtosLevel";
            sheet.Cells[1, 2].Value = "Parameter";
            sheet.Cells[1, 3].Value = "Value";

            sheet.Cells[2, 1].Value = "PinX";
            sheet.Cells[2, 2].Value = "ParamA";
            sheet.Cells[2, 3].Value = "ValueA";

            sheet.Cells[3, 1].Value = "PinX";
            sheet.Cells[3, 2].Value = "ParamB";
            sheet.Cells[3, 3].Value = "ValueB";

            // Act
            RtosTableSheet result = RtosTableSheet.LoadConfig(sheet);

            // Assert
            Assert.AreEqual(1, result.LevelRows.Count);
            Assert.AreEqual("ValueA", result.LevelRows[0].Parameters["ParamA"]);
            Assert.AreEqual("ValueB", result.LevelRows[0].Parameters["ParamB"]);
        }

        [TestMethod]
        public void LoadConfig_LevelRowWithBlankParameterOrValue_SkipsParameterEntry()
        {
            // Arrange
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Rtos");
            sheet.Cells[1, 1].Value = "RtosLevel";
            sheet.Cells[1, 2].Value = "Parameter";
            sheet.Cells[1, 3].Value = "Value";

            sheet.Cells[2, 1].Value = "PinX";

            // Act
            RtosTableSheet result = RtosTableSheet.LoadConfig(sheet);

            // Assert
            Assert.AreEqual(1, result.LevelRows.Count);
            Assert.AreEqual("PinX", result.LevelRows[0].PinName);
            Assert.AreEqual(0, result.LevelRows[0].Parameters.Count);
        }
    }
}
