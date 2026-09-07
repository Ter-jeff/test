using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Dto;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Reader;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Results;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Basic;

[TestClass]
public class TimeSetOverrideRowReaderTest
{
    private static readonly string _inputFolder = Path.Combine(
        Directory.GetCurrentDirectory(), "Input", "Basic", "TimeSetOverride", "Reader"
    );

    public static IEnumerable<object[]> NoErrorCases =>
    [
        ["sheet_valid_rows.xlsx"],
        ["sheet_valid_rows_with_empty_row.xlsx"],
    ];

    [TestMethod]
    [DynamicData(nameof(NoErrorCases), DynamicDataSourceType.Property)]
    public void ReadTimeSetOverrideRows_NoError(string fileName)
    {
        string excelPath = Path.Combine(_inputFolder, fileName);
        FileInfo excelFile = new(excelPath);
        using ExcelPackage package = new(excelFile);
        ExcelWorksheet worksheet = package.Workbook.Worksheets[TimeSetOverrideSchema.SheetName];
        TimeSetOverrideHeaderReader reader = new();
        Result<TimeSetOverrideMetadata, SheetError> headerResult = reader.Read(worksheet);
        if (!headerResult.Success)
        {
            Assert.Fail("Header reading failed while testing TimeSetOverrideRowReader!");
        }
        TimeSetOverrideRowReader rowReader = new();
        TimeSetOverrideRowReaderOutPut result = rowReader.Read(worksheet, headerResult.Value);
        Assert.IsTrue(result.RowErrors.Count == 0);
    }

    [TestMethod]
    public void ReadTimeSetOverrideRows_MissingRequiredCellValues()
    {
        string excelPath = Path.Combine(
            _inputFolder,
            "sheet_invalid_rows_missing_required_cell_values.xlsx"
        );
        FileInfo excelFile = new(excelPath);
        using ExcelPackage package = new(excelFile);
        ExcelWorksheet worksheet = package.Workbook.Worksheets[TimeSetOverrideSchema.SheetName];
        TimeSetOverrideHeaderReader reader = new();
        Result<TimeSetOverrideMetadata, SheetError> headerResult = reader.Read(worksheet);
        if (!headerResult.Success)
        {
            Assert.Fail("Header reading failed while testing TimeSetOverrideRowReader!");
        }
        TimeSetOverrideRowReader rowReader = new();
        TimeSetOverrideRowReaderOutPut result = rowReader.Read(worksheet, headerResult.Value);
        Assert.IsTrue(result.RowErrors.Count == 4);
        foreach (RowError rowError in result.RowErrors)
        {
            Assert.AreEqual(rowError.Error, BasicErrorType.E_TimeSetOverride_02);
        }
    }
}
