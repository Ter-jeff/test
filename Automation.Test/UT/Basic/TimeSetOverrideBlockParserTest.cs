using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Dto;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Parser;

using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Tables;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Basic;

[TestClass]
public class TimeSetOverrideBlockParserTest
{
    private static readonly TimeSetOverrideMetadata _metadata = new(
        TimeSetOverrideSchema.SheetName,
        TimeSetOverrideSchema.ColumnConfigs.Select((c, i) => new ColumnMetadata(c.Name, i + 1)),
        1
    );

    [TestMethod]
    public void ParseTimeSetOverrideBlocks_ValidFrequency()
    {
        List<TimeSetOverrideRow> rows = [
            new()
            {
                RowIndex = 2,
                TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                Frequency = "PCIE_REFCLK_Diff_Freq_VAR",
                TimeSet = "ts1",
                PinGroupName = "PADIO_CKRX_FSYS_CLK",
                Setup = "io",
                DataSrc = "PAT",
                DataFmt = "NR",
            },
            new()
            {
                RowIndex = 3,
                TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                Frequency = "PCIE_REFCLK_Diff_Freq_VAR = 200MHz",
                TimeSet = "ts1",
                PinGroupName = "PCIE_REFCLK_P",
                Setup = "clock_2x",
                DataSrc = "ALLHI",
                DataFmt = "RL",
            },
        ];
        TimeSetOverrideBlockParser parser = new();
        TimeSetOverrideBlockParserOutput blockOutPut = parser.Parse(rows, _metadata);
        Assert.IsTrue(blockOutPut.RowErrors.Count == 0);
    }

    public static IEnumerable<object[]> InvalidVariableNameCases =>
    [
        ["@FOO"],
        ["_BAR"],
        ["FOO_!_BAR"],
        ["123_FOO_BAR"],
        ["@FOO = 100MHz"],
        ["_BAR = 100MHz"],
        ["FOO_!_BAR = 100MHz"],
        ["123_FOO_BAR = 100MHz"],
    ];

    [TestMethod]
    [DynamicData(nameof(InvalidVariableNameCases), DynamicDataSourceType.Property)]
    public void ParseTimeSetOverrideBlocks_InvalidVariableName(string frequency)
    {
        List<TimeSetOverrideRow> rows = [
            new()
            {
                RowIndex = 2,
                TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                Frequency = frequency,
                TimeSet = "ts1",
                PinGroupName = "PADIO_CKRX_FSYS_CLK",
                Setup = "io",
                DataSrc = "PAT",
                DataFmt = "NR",
            },
        ];
        TimeSetOverrideBlockParser parser = new();
        TimeSetOverrideBlockParserOutput blockOutPut = parser.Parse(rows, _metadata);
        Assert.IsTrue(blockOutPut.CellErrors.Count == 1);
        Assert.AreEqual(blockOutPut.CellErrors[0].Error, BasicErrorType.E_TimeSetOverride_03);
    }

    public static IEnumerable<object[]> InvalidFrequencyFormatCases =>
    [
        ["FOO ="],
        ["BAR="],
        ["FOO_BAR = "],
        ["FOO_BAR= "],
    ];

    [TestMethod]
    [DynamicData(nameof(InvalidFrequencyFormatCases), DynamicDataSourceType.Property)]
    public void ParseTimeSetOverrideBlocks_InvalidFrequencyFormat(string frequency)
    {
        List<TimeSetOverrideRow> rows = [
            new()
            {
                RowIndex = 2,
                TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                Frequency = frequency,
                TimeSet = "ts1",
                PinGroupName = "PADIO_CKRX_FSYS_CLK",
                Setup = "io",
                DataSrc = "PAT",
                DataFmt = "NR",
            },
        ];
        TimeSetOverrideBlockParser parser = new();
        TimeSetOverrideBlockParserOutput blockOutPut = parser.Parse(rows, _metadata);
        Assert.IsTrue(blockOutPut.CellErrors.Count == 1);
        Assert.AreEqual(blockOutPut.CellErrors[0].Error, BasicErrorType.E_TimeSetOverride_04);
    }

    public static IEnumerable<object[]> InvalidFrequencyValueCases =>
    [
        ["FOO = FOO"],
        ["BAR= @BAR"],
        ["FOO_BAR = 100-MHz"],
        ["FOO_BAR= 100KM"],
    ];

    [TestMethod]
    [DynamicData(nameof(InvalidFrequencyValueCases), DynamicDataSourceType.Property)]
    public void ParseTimeSetOverrideBlocks_InvalidFrequencyValue(string frequency)
    {
        List<TimeSetOverrideRow> rows = [
            new()
            {
                RowIndex = 2,
                TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                Frequency = frequency,
                TimeSet = "ts1",
                PinGroupName = "PADIO_CKRX_FSYS_CLK",
                Setup = "io",
                DataSrc = "PAT",
                DataFmt = "NR",
            },
        ];
        TimeSetOverrideBlockParser parser = new();
        TimeSetOverrideBlockParserOutput blockOutPut = parser.Parse(rows, _metadata);
        Assert.IsTrue(blockOutPut.CellErrors.Count == 1);
        Assert.AreEqual(blockOutPut.CellErrors[0].Error, BasicErrorType.E_TimeSetOverride_05);
    }

    [TestMethod]
    public void ParseTimeSetOverrideBlocks_FrequencyVariableReassignConflict()
    {
        List<TimeSetOverrideRow> rows = [
            new()
            {
                RowIndex = 2,
                TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                Frequency = "PCIE_REFCLK_Diff_Freq_VAR = 100MHz",
                TimeSet = "ts1",
                PinGroupName = "PADIO_CKRX_FSYS_CLK",
                Setup = "io",
                DataSrc = "PAT",
                DataFmt = "NR",
            },
            new()
            {
                RowIndex = 3,
                TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                Frequency = "PCIE_REFCLK_Diff_Freq_VAR = 200MHz",
                TimeSet = "ts1",
                PinGroupName = "PCIE_REFCLK_P",
                Setup = "clock_2x",
                DataSrc = "ALLHI",
                DataFmt = "RL",
            },
        ];
        TimeSetOverrideBlockParser parser = new();
        TimeSetOverrideBlockParserOutput blockOutPut = parser.Parse(rows, _metadata);
        Assert.IsTrue(blockOutPut.CellErrors.Count == 1);
        Assert.AreEqual(blockOutPut.CellErrors[0].Error, BasicErrorType.E_TimeSetOverride_06);
    }

    public static IEnumerable<object[]> InvalidPinValuesCases =>
    [
        // Empty Setup
        [
            new List<TimeSetOverrideRow>
            {
                new()
                {
                    RowIndex = 2,
                    TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                    Frequency = "PCIE_REFCLK_Diff_Freq_VAR = 100MHz",
                    TimeSet = "ts1",
                    PinGroupName = "PADIO_CKRX_FSYS_CLK",
                    Setup = "",
                    DataSrc = "PAT",
                    DataFmt = "NR",
                }
            },
            new List<ErrorCode>
            {
                BasicErrorType.E_TimeSetOverride_07,
                BasicErrorType.E_TimeSetOverride_08,
            }
        ],
        // Invalid Setup
        [
            new List<TimeSetOverrideRow>
            {
                new()
                {
                    RowIndex = 2,
                    TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                    Frequency = "PCIE_REFCLK_Diff_Freq_VAR = 100MHz",
                    TimeSet = "ts1",
                    PinGroupName = "PADIO_CKRX_FSYS_CLK",
                    Setup = "FOO",
                    DataSrc = "PAT",
                    DataFmt = "NR",
                }
            },
            new List<ErrorCode>
            {
                BasicErrorType.E_TimeSetOverride_08,
            }
        ],
        // Invalid DataSrc
        [
            new List<TimeSetOverrideRow>
            {
                new()
                {
                    RowIndex = 2,
                    TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                    Frequency = "PCIE_REFCLK_Diff_Freq_VAR = 100MHz",
                    TimeSet = "ts1",
                    PinGroupName = "PADIO_CKRX_FSYS_CLK",
                    Setup = "io",
                    DataSrc = "FOO",
                    DataFmt = "NR",
                }
            },
            new List<ErrorCode>
            {
                BasicErrorType.E_TimeSetOverride_09,
            }
        ],
        // Invalid DataFmt
        [
            new List<TimeSetOverrideRow>
            {
                new()
                {
                    RowIndex = 2,
                    TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                    Frequency = "PCIE_REFCLK_Diff_Freq_VAR = 100MHz",
                    TimeSet = "ts1",
                    PinGroupName = "PADIO_CKRX_FSYS_CLK",
                    Setup = "io",
                    DataSrc = "PAT",
                    DataFmt = "FOO",
                }
            },
            new List<ErrorCode>
            {
                BasicErrorType.E_TimeSetOverride_10,
            }
        ],
    ];

    [TestMethod]
    [DynamicData(nameof(InvalidPinValuesCases), DynamicDataSourceType.Property)]
    public void ParseTimeSetOverrideBlocks_InvalidPinValues(
        List<TimeSetOverrideRow> rows,
        List<ErrorCode> errorCodes
    )
    {
        TimeSetOverrideBlockParser parser = new();
        TimeSetOverrideBlockParserOutput blockOutPut = parser.Parse(rows, _metadata);
        Assert.IsTrue(blockOutPut.CellErrors.Count == errorCodes.Count);
        for (int i = 0; i < errorCodes.Count; i++)
        {
            Assert.AreEqual(blockOutPut.CellErrors[i].Error, errorCodes[i]);
        }
    }

    [TestMethod]
    public void ParseTimeSetOverrideBlocks_AllHiButNoRl()
    {
        List<TimeSetOverrideRow> rows = [
            new()
            {
                RowIndex = 2,
                TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                Frequency = "PCIE_REFCLK_Diff_Freq_VAR = 200MHz",
                TimeSet = "ts1",
                PinGroupName = "PCIE_REFCLK_P",
                Setup = "clock_2x",
                DataSrc = "ALLHI",
                DataFmt = "NR",
            },
        ];
        TimeSetOverrideBlockParser parser = new();
        TimeSetOverrideBlockParserOutput blockOutPut = parser.Parse(rows, _metadata);
        Assert.IsTrue(blockOutPut.CellErrors.Count == 1);
        Assert.AreEqual(blockOutPut.CellErrors[0].Error, BasicErrorType.E_TimeSetOverride_11);
    }

    [TestMethod]
    public void ParseTimeSetOverrideBlocks_AllLoButNoRh()
    {
        List<TimeSetOverrideRow> rows = [
            new()
            {
                RowIndex = 2,
                TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                Frequency = "PCIE_REFCLK_Diff_Freq_VAR = 200MHz",
                TimeSet = "ts1",
                PinGroupName = "PCIE_REFCLK_N",
                Setup = "clock_2x",
                DataSrc = "ALLLO",
                DataFmt = "NR",
            },
        ];
        TimeSetOverrideBlockParser parser = new();
        TimeSetOverrideBlockParserOutput blockOutPut = parser.Parse(rows, _metadata);
        Assert.IsTrue(blockOutPut.CellErrors.Count == 1);
        Assert.AreEqual(blockOutPut.CellErrors[0].Error, BasicErrorType.E_TimeSetOverride_12);
    }

    [TestMethod]
    public void ParseTimeSetOverrideBlocks_DuplicateRows()
    {
        List<TimeSetOverrideRow> rows = [
            new()
            {
                RowIndex = 2,
                TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                Frequency = "TEST_VAR = 400MHz",
                TimeSet = "ts1",
                PinGroupName = "XIO_CLK_48",
                Setup = "clock_2x",
                DataSrc = "PAT",
                DataFmt = "NR",
            },
            new()
            {
                RowIndex = 3,
                TimeSetFile = "TIMESET_TLTA0_S_AN_SI",
                Frequency = "TEST_VAR = 400MHz",
                TimeSet = "ts1",
                PinGroupName = "XIO_CLK_48",
                Setup = "clock_2x",
                DataSrc = "",
                DataFmt = "",
            },
        ];
        TimeSetOverrideBlockParser parser = new();
        TimeSetOverrideBlockParserOutput blockOutPut = parser.Parse(rows, _metadata);
        Assert.IsTrue(blockOutPut.RowErrors.Count == 1);
        Assert.AreEqual(blockOutPut.RowErrors[0].Error, BasicErrorType.E_TimeSetOverride_14);
    }
}
