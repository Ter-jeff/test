using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Dto;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Reader;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Results;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override;

public class TimeSetOverrideRowReader
{
    public TimeSetOverrideRowReaderOutPut Read(ExcelWorksheet worksheet, TimeSetOverrideMetadata header)
    {
        int firstRowIndex = header.HeaderRow + 1;

        List<TimeSetOverrideRow> rows = [];
        List<RowError> rowErrors = [];
        for (int i = firstRowIndex; i <= worksheet.Dimension.End.Row; i++)
        {
            if (IsRowEmpty(worksheet, i))
            {
                continue;
            }

            Result<TimeSetOverrideRow, RowError> result = GetRow(worksheet, i, header);
            if (!result.Success)
            {
                rowErrors.Add(result.Error);
                continue;
            }
            rows.Add(result.Value);
        }
        return new(rows, rowErrors);
    }

    private static bool IsRowEmpty(ExcelWorksheet worksheet, int rowIndex)
    {
        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
        {
            string? value = worksheet.Cells[rowIndex, col].Text;

            if (!string.IsNullOrWhiteSpace(value))
            {
                return false;
            }
        }

        return true;
    }

    private static Result<TimeSetOverrideRow, RowError> GetRow(
        ExcelWorksheet sheet,
        int rowIndex,
        TimeSetOverrideMetadata header
    )
    {
        if (
            TimeSetOverrideSchema.ValueRequiredColumns.Any(cf =>
                string.IsNullOrWhiteSpace(
                    sheet.GetCellValue(rowIndex, header.GetColumnIndex(cf.Name))
                )
            )
        )
        {
            string requiredColumns = string.Join(", ", TimeSetOverrideSchema.ValueRequiredColumns
                .Select(cf => cf.Name));
            RowError rowError = new(BasicErrorType.E_TimeSetOverride_02, sheet.Name, rowIndex, [requiredColumns, $"{rowIndex}"]);
            return Result<TimeSetOverrideRow, RowError>.Fail(rowError);
        }

        return Result<TimeSetOverrideRow, RowError>.Ok(
            new()
            {
                RowIndex = rowIndex,
                TimeSetFile = sheet.GetCellValue(
                    rowIndex,
                    header.GetColumnIndex(TimeSetOverrideSchema.TimeSetFileHeader)
                ).Trim(),
                Frequency = sheet.GetCellValue(
                    rowIndex,
                    header.GetColumnIndex(TimeSetOverrideSchema.FrequencyHeader)
                ).Trim(),
                TimeSet = sheet.GetCellValue(
                    rowIndex,
                    header.GetColumnIndex(TimeSetOverrideSchema.TimeSetHeader)
                ).Trim(),
                PinGroupName = sheet.GetCellValue(
                    rowIndex,
                    header.GetColumnIndex(TimeSetOverrideSchema.PinGroupHeader)
                ).Trim(),
                Setup = sheet.GetCellValue(
                    rowIndex,
                    header.GetColumnIndex(TimeSetOverrideSchema.SetupHeader)
                ).Trim(),
                DataSrc = sheet.GetCellValue(
                    rowIndex,
                    header.GetColumnIndex(TimeSetOverrideSchema.DataSrcHeader)
                ).Trim(),
                DataFmt = sheet.GetCellValue(
                    rowIndex,
                    header.GetColumnIndex(TimeSetOverrideSchema.DataFmtHeader)
                ).Trim(),
            }
        );
    }
}
