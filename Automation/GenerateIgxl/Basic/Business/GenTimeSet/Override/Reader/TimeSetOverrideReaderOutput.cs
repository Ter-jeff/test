using System.Collections.Generic;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Reader;

public sealed record TimeSetOverrideReaderOutput(
    IReadOnlyList<TimeSetOverrideBlock> OverrideBlocks,
    IReadOnlyList<RowError> RowErrors,
    IReadOnlyList<CellError> CellErrors
);
