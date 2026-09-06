namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override;

public record TimeSetOverrideRow
{
    public required int RowIndex { get; init; }

    public required string TimeSetFile { get; init; }

    public required string? Frequency { get; init; }

    public required string TimeSet { get; init; }

    public required string PinGroupName { get; init; }

    public required string Setup { get; init; }

    public required string? DataSrc { get; init; }

    public required string? DataFmt { get; init; }
}
