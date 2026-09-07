using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Reader;
using Automation.Static;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Results;

using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Applier;

public sealed class TimeSetOverrideBatchApplier
{
    public TimeSetOverrideBatchApplierOutput Apply(TimeSetSheets sheets, List<string> tsetFileList)
    {
        List<SheetError> sheetErrors = [];
        List<RowError> rowErrors = [];
        List<CellError> cellErrors = [];
        ExcelWorksheet? tsOverrideSheet = EpWorkbook.TestPlanWorkbook.Worksheets[TimeSetOverrideSchema.SheetName];
        if (tsOverrideSheet == null)
        {
            return new([], rowErrors, cellErrors);
        }
        Result<OverrideBlocksOutput, SheetError> blocksResult = GetTimeSetOverrideBlocks(tsOverrideSheet, tsetFileList);
        if (!blocksResult.Success)
        {
            return new([blocksResult.Error], rowErrors, cellErrors);
        }
        IReadOnlyList<TimeSetOverrideBlock>? overrideBlocks = blocksResult.Value.Blocks;
        sheetErrors.AddRange(blocksResult.Value.SheetErrors);
        rowErrors.AddRange(blocksResult.Value.RowErrors);
        cellErrors.AddRange(blocksResult.Value.CellErrors);

        foreach (ComTimeSetBasicSheet sheet in sheets)
        {
            TimeSetOverrideBlock? targetBlock = overrideBlocks?.FirstOrDefault(b =>
                string.Equals(
                    b.TimeSetFile,
                    GetTsNameWithoutVersion(sheet.Name),
                    StringComparison.OrdinalIgnoreCase
                )
            );
            if (targetBlock == null)
            {
                continue;
            }

            TimeSetOverrideApplier applier = new();
            Result<Unit, IReadOnlyList<RowError>> applyResult = applier
                .Apply(sheet, targetBlock, $"{sheet.Name}.txt");
            if (!applyResult.Success)
            {
                rowErrors.AddRange(applyResult.Error);
                continue;
            }
        }
        return new(sheetErrors, rowErrors, cellErrors);
    }

    private sealed record OverrideBlocksOutput(
        IReadOnlyList<TimeSetOverrideBlock> Blocks,
        IReadOnlyList<SheetError> SheetErrors,
        IReadOnlyList<RowError> RowErrors,
        IReadOnlyList<CellError> CellErrors
    );

    private Result<OverrideBlocksOutput, SheetError> GetTimeSetOverrideBlocks(
        ExcelWorksheet tsOverrideSheet,
        List<string> tsetFileList
    )
    {
        TimeSetOverrideReader reader = new();
        Result<TimeSetOverrideReaderOutput, SheetError> result = reader
            .Read(tsOverrideSheet);
        if (!result.Success)
        {
            return Result<OverrideBlocksOutput, SheetError>.Fail(result.Error);
        }
        List<SheetError> sheetErrors = [];
        List<RowError> rowErrors = [];
        List<CellError> cellErrors = [];

        rowErrors.AddRange(result.Value.RowErrors);
        cellErrors.AddRange(result.Value.CellErrors);

        IEnumerable<string> tsNoVersionFiles = tsetFileList.Select(GetTsNameWithoutVersion);

        IEnumerable<TimeSetOverrideBlock> missingTsets = result.Value.OverrideBlocks.Where(ob =>
            !tsNoVersionFiles.Any(ts =>
                string.Equals(ts, ob.TimeSetFile, StringComparison.OrdinalIgnoreCase)
            )
        );
        foreach (TimeSetOverrideBlock overrideBlock in missingTsets)
        {
            SheetError error = new(BasicErrorType.E_TimeSetOverride_13, TimeSetOverrideSchema.SheetName, [overrideBlock.TimeSetFile]);
            sheetErrors.Add(error);
        }

        IEnumerable<TimeSetOverrideBlock> existTsets = result.Value.OverrideBlocks.Where(ob =>
            tsNoVersionFiles.Any(ts =>
                string.Equals(ts, ob.TimeSetFile, StringComparison.OrdinalIgnoreCase)
            )
        );
        return Result<OverrideBlocksOutput, SheetError>
            .Ok(new([.. existTsets], sheetErrors, rowErrors, cellErrors));
    }

    private static string GetTsNameWithoutVersion(string tsName)
    {
        int lastUnderscore = tsName.LastIndexOf('_');
        return lastUnderscore >= 0
            ? tsName[..lastUnderscore]
            : tsName;
    }
}
