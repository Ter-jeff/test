using System.Collections.Generic;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Parser;

public sealed record TimeSetOverrideBlockParserOutput(
    IReadOnlyList<TimeSetOverrideBlock> OverrideBlocks,
    IReadOnlyList<RowError> RowErrors,
    IReadOnlyList<CellError> CellErrors
);
