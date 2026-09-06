using System.IO;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Dto;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Results;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Basic;

[TestClass]
public class TimeSetOverrideHeaderReaderTest
{
    private static readonly string _inputFolder = Path.Combine(
        Directory.GetCurrentDirectory(), "Input", "Basic", "TimeSetOverride", "Reader"
    );

    [TestMethod]
    public void ReadTimeSetOverrideHeader_ReturnSuccess()
    {
        string excelPath = Path.Combine(_inputFolder, "sheet_valid_header.xlsx");
        FileInfo excelFile = new(excelPath);
        using ExcelPackage package = new(excelFile);
        ExcelWorksheet worksheet = package.Workbook.Worksheets[TimeSetOverrideSchema.SheetName];
        TimeSetOverrideHeaderReader reader = new();
        Result<TimeSetOverrideMetadata, SheetError> result = reader.Read(worksheet);
        Assert.IsTrue(result.Success);
    }

    [TestMethod]
    public void ReadTimeSetOverrideHeader_ReturnFail()
    {
        string excelPath = Path.Combine(_inputFolder, "sheet_invalid_header.xlsx");
        FileInfo excelFile = new(excelPath);
        using ExcelPackage package = new(excelFile);
        ExcelWorksheet worksheet = package.Workbook.Worksheets[TimeSetOverrideSchema.SheetName];
        TimeSetOverrideHeaderReader reader = new();
        Result<TimeSetOverrideMetadata, SheetError> result = reader.Read(worksheet);
        Assert.IsFalse(result.Success);
        Assert.AreEqual(result.Error.Error, BasicErrorType.E_TimeSetOverride_01);
    }

    [TestMethod]
    public void ReadTimeSetOverrideHeader_GetNonExistHeaderIndex()
    {
        string excelPath = Path.Combine(_inputFolder, "sheet_valid_header.xlsx");
        FileInfo excelFile = new(excelPath);
        using ExcelPackage package = new(excelFile);
        ExcelWorksheet worksheet = package.Workbook.Worksheets[TimeSetOverrideSchema.SheetName];
        TimeSetOverrideHeaderReader reader = new();
        Result<TimeSetOverrideMetadata, SheetError> result = reader.Read(worksheet);
        Assert.IsTrue(result.Success);
        int index = result.Value.GetColumnIndex("FOO");
        Assert.AreEqual(index, -1);
    }
}
