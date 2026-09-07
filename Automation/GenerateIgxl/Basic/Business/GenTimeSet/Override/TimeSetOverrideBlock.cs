using System.Collections.Generic;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override;

public record TimeSetOverrideBlock(
    string TimeSetFile,
    IReadOnlyList<TimeSetFrequencyOverride> FrequencyOverrides,
    IReadOnlyList<TimeSetPinValueOverride> PinValueOverrides
);

public sealed record TimeSetFrequencyOverride(string Name, string Value, int RowIndex);

public sealed record TimeSetPinValueOverride(
    int RowIndex,
    string TimeSet,
    string PinName,
    string Setup,
    string? DataSrc,
    string? DataFmt,
    string? Variable
);
