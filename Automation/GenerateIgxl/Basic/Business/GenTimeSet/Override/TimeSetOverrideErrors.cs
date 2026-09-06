using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override;

public sealed record SheetError(ErrorCode Error, string SheetName, string[]? Args = null);

public sealed record RowError(ErrorCode Error, string SheetName, int RowIndex, string[]? Args = null);

public sealed record CellError(ErrorCode Error, string SheetName, int RowIndex, int ColumnIndex, string[]? Args = null);

public static class ErrorExtensions
{
    public static void Report(this SheetError sheetError)
    {
        ErrorReportManager.AddError(sheetError.Error, sheetError.SheetName, -1, -1, sheetError.Args);
    }

    public static void Report(this RowError rowError)
    {
        ErrorReportManager.AddError(rowError.Error, rowError.SheetName, rowError.RowIndex, -1, rowError.Args);
    }

    public static void Report(this CellError cellError)
    {
        ErrorReportManager.AddError(cellError.Error, cellError.SheetName, cellError.RowIndex, cellError.ColumnIndex, cellError.Args);
    }

    public static string GetMessage(this SheetError error)
    {
        string msg = error.Error.FormatMessage(error.Args);
        string fullMsg = $"[{error.SheetName}] {msg}";
        return fullMsg;
    }

    public static string GetMessage(this RowError error)
    {
        string msg = error.Error.FormatMessage(error.Args);
        string fullMsg = $"[{error.SheetName}] [row:{error.RowIndex}] {msg}";
        return fullMsg;
    }

    public static string GetMessage(this CellError error)
    {
        string msg = error.Error.FormatMessage(error.Args);
        string fullMsg = $"[{error.SheetName}] [row:{error.RowIndex},col:{error.ColumnIndex}] {msg}";
        return fullMsg;
    }
}
