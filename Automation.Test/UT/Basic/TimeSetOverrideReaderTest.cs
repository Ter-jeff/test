using System.IO;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Reader;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Results;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Basic;

[TestClass]
public class TimeSetOverrideReaderTest
{
    private static readonly string _inputFolder = Path.Combine(
        Directory.GetCurrentDirectory(), "Input", "Basic", "TimeSetOverride", "Reader"
    );

    [TestMethod]
    public void ReadTimeSetOverrides_NoError()
    {
        string excelPath = Path.Combine(_inputFolder, "sheet_main_reader_integration_no_error.xlsx");
        FileInfo excelFile = new(excelPath);
        using ExcelPackage package = new(excelFile);
        ExcelWorksheet worksheet = package.Workbook.Worksheets[TimeSetOverrideSchema.SheetName];
        TimeSetOverrideReader reader = new();
        Result<TimeSetOverrideReaderOutput, SheetError> result = reader.Read(worksheet);
        if (!result.Success)
        {
            Assert.Fail(result.Error.GetMessage());
        }
        Assert.IsTrue(result.Value.RowErrors.Count == 0);
        Assert.IsTrue(result.Value.CellErrors.Count == 0);
    }

    [TestMethod]
    public void ReadTimeSetOverrides_InvalidHeaders()
    {
        string excelPath = Path.Combine(_inputFolder, "sheet_invalid_header.xlsx");
        FileInfo excelFile = new(excelPath);
        using ExcelPackage package = new(excelFile);
        ExcelWorksheet worksheet = package.Workbook.Worksheets[TimeSetOverrideSchema.SheetName];
        TimeSetOverrideReader reader = new();
        Result<TimeSetOverrideReaderOutput, SheetError> result = reader.Read(worksheet);
        Assert.IsFalse(result.Success);
        Assert.IsTrue(result.Error.Error == BasicErrorType.E_TimeSetOverride_01);
    }
}
