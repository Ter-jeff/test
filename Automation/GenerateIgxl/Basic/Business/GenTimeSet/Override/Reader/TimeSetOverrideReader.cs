using System.Collections.Generic;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Dto;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Parser;

using CommonLib.Results;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Reader;

public class TimeSetOverrideReader
{
    public Result<TimeSetOverrideReaderOutput, SheetError> Read(ExcelWorksheet worksheet)
    {
        Result<TimeSetOverrideMetadata, SheetError> headerResult =
            new TimeSetOverrideHeaderReader().Read(worksheet);

        if (!headerResult.Success)
        {
            return Result<TimeSetOverrideReaderOutput, SheetError>.Fail(headerResult.Error);
        }

        List<RowError> rowErrors = [];
        List<CellError> cellErrors = [];

        TimeSetOverrideRowReader rowReader = new();
        TimeSetOverrideRowReaderOutPut rowResult = rowReader.Read(worksheet, headerResult.Value);

        rowErrors.AddRange(rowResult.RowErrors);

        // Parse row data to OverrideBlock
        TimeSetOverrideBlockParser blockParser = new();
        TimeSetOverrideBlockParserOutput blocksResult = blockParser
            .Parse(rowResult.Rows, headerResult.Value);
        rowErrors.AddRange(blocksResult.RowErrors);
        cellErrors.AddRange(blocksResult.CellErrors);

        return Result<TimeSetOverrideReaderOutput, SheetError>.Ok(
            new(blocksResult.OverrideBlocks, rowErrors, cellErrors)
        );
    }
}
