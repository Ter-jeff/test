using System.Collections.Generic;
using System.Linq;

using CommonLib.Tables;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override;

public static class TimeSetOverrideSchema
{
    public const string SheetName = "TimeSetClockOverride";
    public const string TimeSetFileHeader = "TimeSet File";
    public const string FrequencyHeader = "Frequency";
    public const string TimeSetHeader = "TimeSet";
    public const string PinGroupHeader = "Pin/Group";
    public const string SetupHeader = "Setup";
    public const string DataSrcHeader = "Data Src";
    public const string DataFmtHeader = "Data Fmt";

    public static IReadOnlyList<ColumnConfig> ColumnConfigs { get; private set; } = [
        new(TimeSetFileHeader, true, true),
        new(FrequencyHeader, true, false),
        new(TimeSetHeader, true, true),
        new(PinGroupHeader, true, true),
        new(SetupHeader, true, true),
        new(DataSrcHeader, true, false),
        new(DataFmtHeader, true, false),
    ];

    public static IReadOnlyList<ColumnConfig> RequiredColumns
    {
        get { return [.. ColumnConfigs.Where(cf => cf.Required)]; }
    }

    public static IReadOnlyList<ColumnConfig> ValueRequiredColumns
    {
        get { return [.. ColumnConfigs.Where(cf => cf.ValueRequired)]; }
    }

    public static string GetRequiredColumnNames()
    {
        return string.Join(", ", ColumnConfigs.Where(cf => cf.Required).Select(cf => $"`{cf.Name}`"));
    }
}
