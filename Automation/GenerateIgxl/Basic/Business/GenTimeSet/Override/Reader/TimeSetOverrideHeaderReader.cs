using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Dto;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Results;
using CommonLib.Tables;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override;

public class TimeSetOverrideHeaderReader
{
    public Result<TimeSetOverrideMetadata, SheetError> Read(ExcelWorksheet worksheet)
    {
        int headerIndex = GetValidHeaderRowIndex(worksheet, 1);
        if (headerIndex == -1)
        {
            string requiredColumns = TimeSetOverrideSchema.GetRequiredColumnNames();
            string msg = $"{requiredColumns} columns are required!";
            SheetError sheetError = new(BasicErrorType.E_TimeSetOverride_01, worksheet.Name, [msg]);
            return Result<TimeSetOverrideMetadata, SheetError>.Fail(sheetError);
        }

        IEnumerable<ColumnMetadata> headerMetadata = GetHeaderMetadata(worksheet, headerIndex);
        return Result<TimeSetOverrideMetadata, SheetError>
            .Ok(new(worksheet.Name, headerMetadata, headerIndex));
    }

    private static IEnumerable<ColumnMetadata> GetHeaderMetadata(ExcelWorksheet sheet, int rowIndex)
    {
        Dictionary<string, int> names = EpplusExtensions.GetHeaderOrder(sheet, rowIndex);
        return names.Select(pair => new ColumnMetadata(pair.Key, pair.Value));
    }

    private static int GetValidHeaderRowIndex(ExcelWorksheet sheet, int startRow)
    {
        for (int i = startRow; i <= sheet.Dimension.End.Row; i++)
        {
            if (IsRowValidHeader(sheet, i))
            {
                return i;
            }
        }
        return -1;
    }

    private static bool IsRowValidHeader(ExcelWorksheet sheet, int rowIndex)
    {
        Dictionary<string, int> names = EpplusExtensions.GetHeaderOrder(sheet, rowIndex);
        foreach (
            ColumnConfig colConf in TimeSetOverrideSchema.ColumnConfigs.Where(cf => cf.Required)
        )
        {
            bool exist = names.ContainsKey(colConf.Name);
            if (!exist)
            {
                return false;
            }
        }
        return true;
    }
}
