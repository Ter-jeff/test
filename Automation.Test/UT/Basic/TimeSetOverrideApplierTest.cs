using System.Collections.Generic;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Applier;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Results;

using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Basic;

[TestClass]
public class TimeSetOverrideApplierTest
{
    [TestMethod]
    public void ApplyTimeSetOverrides_NoError()
    {
        string filePath = @"C:\test\TIMESET_FILE_1.txt";
        // Variable reassign will only cause error in parser, applier should only override old value
        List<TimeSetFrequencyOverride> frequencyOverrides =
        [
            new("PCIE_REFCLK_Diff_Freq_VAR", "1000000", 1),
            new("PCIE_REFCLK_Diff_Freq_VAR", "2000000", 1),
        ];
        List<TimeSetPinValueOverride> pinValueOverrides =
        [
            new(2, "ts1", "PADIO_CKRX_FSYS_CLK", "i/o", "PAT", "NR", "PCIE_REFCLK_Diff_Freq_VAR"),
            new(3, "ts1", "PCIE_REFCLK_P", "clock_2x", "ALLHI", "RL", "PCIE_REFCLK_Diff_Freq_VAR"),
            new(4, "ts1", "PCIE_REFCLK_N", "clock", "ALLHI", "RL", "PCIE_REFCLK_Diff_Freq_VAR"),
        ];
        TimeSetOverrideBlock block = new("TIMESET_FILE_1", frequencyOverrides, pinValueOverrides);
        ComTimeSetBasicSheet tsSheet = new("TIMESET_FILE_1")
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
                    ]
                }
            ],
        };
        TimeSetOverrideApplier applier = new();
        Result<Unit, IReadOnlyList<RowError>> result = applier.Apply(tsSheet, block, filePath);
        Assert.IsTrue(result.Success);
    }

    [TestMethod]
    public void ApplyTimeSetOverrides_NoTimeSet()
    {
        string filePath = @"C:\test\TIMESET_FILE_1.txt";
        List<TimeSetFrequencyOverride> frequencyOverrides =
        [
            new("PCIE_REFCLK_Diff_Freq_VAR", "2000000", 1)
        ];
        List<TimeSetPinValueOverride> pinValueOverrides =
        [
            new(2, "ts1", "PADIO_CKRX_FSYS_CLK", "i/o", "PAT", "NR", "PCIE_REFCLK_Diff_Freq_VAR"),
        ];
        TimeSetOverrideBlock block = new("TIMESET_FILE_1", frequencyOverrides, pinValueOverrides);
        ComTimeSetBasicSheet tsSheet = new("TIMESET_FILE_1")
        {
            Rows = []
        };
        TimeSetOverrideApplier applier = new();
        Result<Unit, IReadOnlyList<RowError>> result = applier.Apply(tsSheet, block, filePath);
        Assert.IsFalse(result.Success);
        Assert.IsTrue(result.Error.Count == 1);
        Assert.IsTrue(result.Error[0].Error == BasicErrorType.E_TimeSetOverride_21);
    }

    [TestMethod]
    public void ApplyTimeSetOverrides_NoTargetTimingRow()
    {
        string filePath = @"C:\test\TIMESET_FILE_1.txt";
        List<TimeSetFrequencyOverride> frequencyOverrides =
        [
            new("PCIE_REFCLK_Diff_Freq_VAR", "2000000", 1)
        ];
        List<TimeSetPinValueOverride> pinValueOverrides =
        [
            new(2, "ts1", "PADIO_CKRX_FSYS_CLK", "i/o", "PAT", "NR", "PCIE_REFCLK_Diff_Freq_VAR"),
            new(3, "ts1", "PCIE_REFCLK_P", "clock_2x", "ALLHI", "RL", "PCIE_REFCLK_Diff_Freq_VAR"),
        ];
        TimeSetOverrideBlock block = new("TIMESET_FILE_1", frequencyOverrides, pinValueOverrides);
        ComTimeSetBasicSheet tsSheet = new("TIMESET_FILE_1")
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
                    ]
                }
            ],
        };
        TimeSetOverrideApplier applier = new();
        Result<Unit, IReadOnlyList<RowError>> result = applier.Apply(tsSheet, block, filePath);
        Assert.IsFalse(result.Success);
        Assert.IsTrue(result.Error.Count == 1);
        Assert.IsTrue(result.Error[0].Error == BasicErrorType.E_TimeSetOverride_15);
    }

    [TestMethod]
    public void ApplyTimeSetOverrides_MultipleVariablesInTimeSetFileClockPeriod()
    {
        string filePath = @"C:\test\TIMESET_FILE_1.txt";
        List<TimeSetFrequencyOverride> frequencyOverrides =
        [
            new("PCIE_REFCLK_Diff_Freq_VAR", "2000000", 1)
        ];
        List<TimeSetPinValueOverride> pinValueOverrides =
        [
            new(2, "ts1", "PADIO_CKRX_FSYS_CLK", "i/o", "PAT", "NR", ""),
        ];
        TimeSetOverrideBlock block = new("TIMESET_FILE_1", frequencyOverrides, pinValueOverrides);
        ComTimeSetBasicSheet tsSheet = new("TIMESET_FILE_1")
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
                            PinGrpClockPeriod = "=(1/_PADIO_Diff_Freq_VAR)*_PADIO_Diff_Freq2_VAR",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PADIO_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PADIO_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                    ]
                }
            ],
        };
        TimeSetOverrideApplier applier = new();
        Result<Unit, IReadOnlyList<RowError>> result = applier.Apply(tsSheet, block, filePath);
        Assert.IsFalse(result.Success);
        Assert.IsTrue(result.Error.Count == 1);
        Assert.IsTrue(result.Error[0].Error == BasicErrorType.E_TimeSetOverride_17);
    }

    [TestMethod]
    public void ApplyTimeSetOverrides_NoVariablesInTimeSetFileClockPeriod()
    {
        string filePath = @"C:\test\TIMESET_FILE_1.txt";
        List<TimeSetFrequencyOverride> frequencyOverrides =
        [
            new("PCIE_REFCLK_Diff_Freq_VAR", "2000000", 1)
        ];
        List<TimeSetPinValueOverride> pinValueOverrides =
        [
            new(2, "ts1", "PADIO_CKRX_FSYS_CLK", "i/o", "PAT", "NR", ""),
        ];
        TimeSetOverrideBlock block = new("TIMESET_FILE_1", frequencyOverrides, pinValueOverrides);
        ComTimeSetBasicSheet tsSheet = new("TIMESET_FILE_1")
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
                            PinGrpClockPeriod = "200",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RL",
                            DriveOn = "0",
                            DriveData = "=(1/_PADIO_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PADIO_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                    ]
                }
            ],
        };
        TimeSetOverrideApplier applier = new();
        Result<Unit, IReadOnlyList<RowError>> result = applier.Apply(tsSheet, block, filePath);
        Assert.IsFalse(result.Success);
        Assert.IsTrue(result.Error.Count == 1);
        Assert.IsTrue(result.Error[0].Error == BasicErrorType.E_TimeSetOverride_16);
    }

    [TestMethod]
    public void ApplyTimeSetOverrides_InvalidSetup()
    {
        string filePath = @"C:\test\TIMESET_FILE_1.txt";
        List<TimeSetFrequencyOverride> frequencyOverrides =
        [
            new("PCIE_REFCLK_Diff_Freq_VAR", "2000000", 1)
        ];
        List<TimeSetPinValueOverride> pinValueOverrides =
        [
            new(2, "ts1", "PADIO_CKRX_FSYS_CLK", "FOO", "PAT", "NR", "PCIE_REFCLK_Diff_Freq_VAR"),
        ];
        TimeSetOverrideBlock block = new("TIMESET_FILE_1", frequencyOverrides, pinValueOverrides);
        ComTimeSetBasicSheet tsSheet = new("TIMESET_FILE_1")
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
                    ]
                }
            ],
        };
        TimeSetOverrideApplier applier = new();
        Result<Unit, IReadOnlyList<RowError>> result = applier.Apply(tsSheet, block, filePath);
        Assert.IsFalse(result.Success);
        Assert.IsTrue(result.Error.Count == 1);
        Assert.IsTrue(result.Error[0].Error == BasicErrorType.E_TimeSetOverride_18);
    }

    [TestMethod]
    public void ApplyTimeSetOverrides_AllHiButNoRl()
    {
        string filePath = @"C:\test\TIMESET_FILE_1.txt";
        List<TimeSetFrequencyOverride> frequencyOverrides =
        [
            new("PCIE_REFCLK_Diff_Freq_VAR", "2000000", 1)
        ];
        List<TimeSetPinValueOverride> pinValueOverrides =
        [
            new(3, "ts1", "PCIE_REFCLK_P", "clock_2x", "ALLHI", "", "PCIE_REFCLK_Diff_Freq_VAR"),
        ];
        TimeSetOverrideBlock block = new("TIMESET_FILE_1", frequencyOverrides, pinValueOverrides);
        ComTimeSetBasicSheet tsSheet = new("TIMESET_FILE_1")
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
                            PinGrpName = "PCIE_REFCLK_P",
                            PinGrpClockPeriod = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            PinGrpSetup = "clock",
                            DataSrc = "ALLHI",
                            DataFmt = "RH",
                            DriveOn = "0",
                            DriveData = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)*0.5",
                            DriveReturn = "=(1/_PCIE_REFCLK_Diff_Freq_VAR)",
                            DriveOff = "0",
                        },
                    ]
                }
            ],
        };
        TimeSetOverrideApplier applier = new();
        Result<Unit, IReadOnlyList<RowError>> result = applier.Apply(tsSheet, block, filePath);
        Assert.IsFalse(result.Success);
        Assert.IsTrue(result.Error.Count == 1);
        Assert.IsTrue(result.Error[0].Error == BasicErrorType.E_TimeSetOverride_19);
    }

    [TestMethod]
    public void ApplyTimeSetOverrides_AllLoButNoRh()
    {
        string filePath = @"C:\test\TIMESET_FILE_1.txt";
        List<TimeSetFrequencyOverride> frequencyOverrides =
        [
            new("PCIE_REFCLK_Diff_Freq_VAR", "2000000", 1)
        ];
        List<TimeSetPinValueOverride> pinValueOverrides =
        [
            new(3, "ts1", "PCIE_REFCLK_P", "clock_2x", "ALLLO", "", "PCIE_REFCLK_Diff_Freq_VAR"),
        ];
        TimeSetOverrideBlock block = new("TIMESET_FILE_1", frequencyOverrides, pinValueOverrides);
        ComTimeSetBasicSheet tsSheet = new("TIMESET_FILE_1")
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
                    ]
                }
            ],
        };
        TimeSetOverrideApplier applier = new();
        Result<Unit, IReadOnlyList<RowError>> result = applier.Apply(tsSheet, block, filePath);
        Assert.IsFalse(result.Success);
        Assert.IsTrue(result.Error.Count == 1);
        Assert.IsTrue(result.Error[0].Error == BasicErrorType.E_TimeSetOverride_20);
    }
}
