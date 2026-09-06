using System.Collections.Generic;

using CommonLib.Enums;

using LogLib.Static;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Applier;

public sealed record TimeSetOverrideBatchApplierOutput(
    IReadOnlyList<SheetError> SheetErrors,
    IReadOnlyList<RowError> RowErrors,
    IReadOnlyList<CellError> CellErrors
)
{
    public void Report()
    {
        foreach (SheetError error in SheetErrors)
        {
            error.Report();
            Response.Report(error.GetMessage(), EnumMessageLevel.Error);
        }
        foreach (RowError rowError in RowErrors)
        {
            rowError.Report();
            Response.Report(rowError.GetMessage(), EnumMessageLevel.Error);
        }
        foreach (CellError cellError in CellErrors)
        {
            cellError.Report();
            Response.Report(cellError.GetMessage(), EnumMessageLevel.Error);
        }
    }
}
