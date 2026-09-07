using System.Collections.Generic;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Reader;

public sealed record TimeSetOverrideRowReaderOutPut(
    IReadOnlyList<TimeSetOverrideRow> Rows,
    IReadOnlyList<RowError> RowErrors
);
