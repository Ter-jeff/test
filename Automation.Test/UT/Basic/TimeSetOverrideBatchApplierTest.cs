using System.IO;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Applier;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Basic;

[TestClass]
public class TimeSetOverrideBatchApplierTest
{
    private static readonly string _inputFolder = Path.Combine(
        Directory.GetCurrentDirectory(), "Input", "Basic", "TimeSetOverride", "Reader"
    );

    [TestMethod]
    public void BatchOverrideTimeSet_NoError()
    {
        ComTimeSetBasicSheet tsSheet = new("TIMESET_TLTA0_S_AN_SI_1")
        {
            Rows =
            [
                new ComTimeSetBasic()
                {
                    Name = "ts1",
                    TimingRows =
                    [
                        new()
                        {
                            PinGrpName = "PADIO_CKRX_FSYS_CLK",
                            PinGrpClockPeriod = "=(1/_PADIO_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PADIO_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PADIO_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                        new()
                        {
                            PinGrpName = "PCIE_REFCLK_P",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                        new()
                        {
                            PinGrpName = "PCIE_REFCLK_N",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                        new()
                        {
                            PinGrpName = "XIO_CLK_48",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                    ]
                }
            ],
        };
        TimeSetSheets tsSheets = [tsSheet];
        string excelPath = Path.Combine(_inputFolder, "sheet_main_reader_integration_no_error.xlsx");
        FileInfo excelFile = new(excelPath);
        using ExcelPackage package = new(excelFile);
        EpWorkbook.TestPlanWorkbook = package.Workbook;
        ExcelWorksheet worksheet = package.Workbook.Worksheets[TimeSetOverrideSchema.SheetName];
        TimeSetOverrideBatchApplier applier = new();
        TimeSetOverrideBatchApplierOutput output = applier
            .Apply(tsSheets, ["TIMESET_TLTA0_S_AN_SI_1.txt"]);
        Assert.IsTrue(output.SheetErrors.Count == 0);
        Assert.IsTrue(output.RowErrors.Count == 0);
        Assert.IsTrue(output.CellErrors.Count == 0);
    }

    [TestMethod]
    public void BatchOverrideTimeSet_NotInPatternDashboard()
    {
        ComTimeSetBasicSheet tsSheet = new("TIMESET_TLTA0_S_AN_SI_1")
        {
            Rows =
            [
                new ComTimeSetBasic()
                {
                    Name = "ts1",
                    TimingRows =
                    [
                        new()
                        {
                            PinGrpName = "PADIO_CKRX_FSYS_CLK",
                            PinGrpClockPeriod = "=(1/_PADIO_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PADIO_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PADIO_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                        new()
                        {
                            PinGrpName = "PCIE_REFCLK_P",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                        new()
                        {
                            PinGrpName = "PCIE_REFCLK_N",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                        new()
                        {
                            PinGrpName = "XIO_CLK_48",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                    ]
                }
            ],
        };
        TimeSetSheets tsSheets = [tsSheet];
        string excelPath = Path.Combine(_inputFolder, "sheet_main_reader_integration_no_error.xlsx");
        FileInfo excelFile = new(excelPath);
        using ExcelPackage package = new(excelFile);
        EpWorkbook.TestPlanWorkbook = package.Workbook;
        ExcelWorksheet worksheet = package.Workbook.Worksheets[TimeSetOverrideSchema.SheetName];
        TimeSetOverrideBatchApplier applier = new();
        TimeSetOverrideBatchApplierOutput output = applier.Apply(tsSheets, [""]);
        Assert.IsTrue(output.SheetErrors.Count == 1);
        Assert.AreEqual(output.SheetErrors[0].Error, BasicErrorType.E_TimeSetOverride_13);
        Assert.IsTrue(output.RowErrors.Count == 0);
        Assert.IsTrue(output.CellErrors.Count == 0);
    }

    [TestMethod]
    public void BatchOverrideTimeSet_HeaderError()
    {
        ComTimeSetBasicSheet tsSheet = new("TIMESET_TLTA0_S_AN_SI_1")
        {
            Rows =
            [
                new ComTimeSetBasic()
                {
                    Name = "ts1",
                    TimingRows =
                    [
                        new()
                        {
                            PinGrpName = "PADIO_CKRX_FSYS_CLK",
                            PinGrpClockPeriod = "=(1/_PADIO_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PADIO_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PADIO_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                        new()
                        {
                            PinGrpName = "PCIE_REFCLK_P",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                        new()
                        {
                            PinGrpName = "PCIE_REFCLK_N",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                        new()
                        {
                            PinGrpName = "XIO_CLK_48",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                    ]
                }
            ],
        };
        TimeSetSheets tsSheets = [tsSheet];
        string excelPath = Path.Combine(_inputFolder, "sheet_invalid_header.xlsx");
        FileInfo excelFile = new(excelPath);
        using ExcelPackage package = new(excelFile);
        EpWorkbook.TestPlanWorkbook = package.Workbook;
        ExcelWorksheet worksheet = package.Workbook.Worksheets[TimeSetOverrideSchema.SheetName];
        TimeSetOverrideBatchApplier applier = new();
        TimeSetOverrideBatchApplierOutput output = applier.Apply(tsSheets, [""]);
        Assert.IsTrue(output.SheetErrors.Count == 1);
        Assert.AreEqual(output.SheetErrors[0].Error, BasicErrorType.E_TimeSetOverride_01);
        Assert.IsTrue(output.RowErrors.Count == 0);
        Assert.IsTrue(output.CellErrors.Count == 0);
    }

    [TestMethod]
    public void BatchOverrideTimeSet_ApplyTimeSetNotFound()
    {
        ComTimeSetBasicSheet tsSheet = new("TIMESET_TLTA0_S_AN_SI_1")
        {
            Rows =
            [
                new ComTimeSetBasic()
                {
                    Name = "ts2"
                }
            ],
        };
        TimeSetSheets tsSheets = [tsSheet];
        string excelPath = Path.Combine(_inputFolder, "sheet_main_reader_integration_no_error.xlsx");
        FileInfo excelFile = new(excelPath);
        using ExcelPackage package = new(excelFile);
        EpWorkbook.TestPlanWorkbook = package.Workbook;
        ExcelWorksheet worksheet = package.Workbook.Worksheets[TimeSetOverrideSchema.SheetName];
        TimeSetOverrideBatchApplier applier = new();
        TimeSetOverrideBatchApplierOutput output = applier
            .Apply(tsSheets, ["TIMESET_TLTA0_S_AN_SI_1.txt"]);
        Assert.IsTrue(output.SheetErrors.Count == 0);
        Assert.IsTrue(output.RowErrors.Count == 1);
        Assert.AreEqual(output.RowErrors[0].Error, BasicErrorType.E_TimeSetOverride_21);
        Assert.IsTrue(output.CellErrors.Count == 0);
    }

    [TestMethod]
    public void BatchOverrideTimeSet_NewFrequencyVariableNoValue()
    {
        ComTimeSetBasicSheet tsSheet = new("TIMESET_TLTA0_S_AN_SI_1")
        {
            Rows =
            [
                new ComTimeSetBasic()
                {
                    Name = "ts1",
                    TimingRows =
                    [
                        new()
                        {
                            PinGrpName = "PADIO_CKRX_FSYS_CLK",
                            PinGrpClockPeriod = "=(1/_PADIO_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PADIO_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PADIO_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                        new()
                        {
                            PinGrpName = "PCIE_REFCLK_P",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                        new()
                        {
                            PinGrpName = "PCIE_REFCLK_N",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                        new()
                        {
                            PinGrpName = "XIO_CLK_48",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                    ]
                }
            ],
        };
        TimeSetSheets tsSheets = [tsSheet];
        string excelPath = Path.Combine(_inputFolder, "sheet_invalid_frequency_new_var_no_value.xlsx");
        FileInfo excelFile = new(excelPath);
        using ExcelPackage package = new(excelFile);
        EpWorkbook.TestPlanWorkbook = package.Workbook;
        ExcelWorksheet worksheet = package.Workbook.Worksheets[TimeSetOverrideSchema.SheetName];
        TimeSetOverrideBatchApplier applier = new();
        TimeSetOverrideBatchApplierOutput output = applier
            .Apply(tsSheets, ["TIMESET_TLTA0_S_AN_SI_1.txt"]);
        Assert.IsTrue(output.SheetErrors.Count == 0);
        Assert.IsTrue(output.RowErrors.Count == 1);
        Assert.AreEqual(output.RowErrors[0].Error, BasicErrorType.E_TimeSetOverride_22);
        Assert.IsTrue(output.CellErrors.Count == 0);
    }

    [TestMethod]
    public void BatchOverrideTimeSet_ReportSuccess()
    {
        SheetError sheetError = new(
            BasicErrorType.E_TimeSetOverride_13,
            TimeSetOverrideSchema.SheetName,
            ["123"]
        );
        RowError rowError = new(
            BasicErrorType.E_TimeSetOverride_21,
            TimeSetOverrideSchema.SheetName,
            1,
            ["123", "456"]
        );
        CellError cellError = new(
            BasicErrorType.E_TimeSetOverride_06,
            TimeSetOverrideSchema.SheetName,
            1,
            1,
            ["123", "456"]
        );
        TimeSetOverrideBatchApplierOutput output = new([sheetError], [rowError], [cellError]);
        output.Report();

        ErrorCode? targetSheetError = ErrorReportManager.GetErrorList()
            .Select(e => e.ErrorCode)
            .SingleOrDefault(e => e == BasicErrorType.E_TimeSetOverride_13);
        Assert.IsNotNull(targetSheetError);
        Assert.AreEqual(targetSheetError, sheetError.Error);

        ErrorCode? targetRowError = ErrorReportManager.GetErrorList()
            .Select(e => e.ErrorCode)
            .SingleOrDefault(e => e == BasicErrorType.E_TimeSetOverride_21);
        Assert.IsNotNull(targetRowError);
        Assert.AreEqual(targetRowError, rowError.Error);

        ErrorCode? targetCellError = ErrorReportManager.GetErrorList()
            .Select(e => e.ErrorCode)
            .SingleOrDefault(e => e == BasicErrorType.E_TimeSetOverride_06);
        Assert.IsNotNull(targetCellError);
        Assert.AreEqual(targetCellError, cellError.Error);
    }
}
