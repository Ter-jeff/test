using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Constants;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Dto;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Results;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Parser;

public sealed class TimeSetOverrideBlockParser
{
    private static readonly Dictionary<string, string> _validSetups = new(
        StringComparer.OrdinalIgnoreCase
    )
    {
        { "io", TimeSetOverrideSetups.Io },
        { "i/o", TimeSetOverrideSetups.Io },
        { "clock", TimeSetOverrideSetups.Clock },
        { "clock_2x", TimeSetOverrideSetups.Clock2X },
    };

    private static readonly Dictionary<string, string> _validDataSrcs = new(
        StringComparer.OrdinalIgnoreCase
    )
    {
        { "ALLHI", TimeSetOverrideDataSrcs.AllHi },
        { "ALLLO", TimeSetOverrideDataSrcs.AllLo },
        { "PA", TimeSetOverrideDataSrcs.Pa },
        { "PAT", TimeSetOverrideDataSrcs.Pat },
        { "PATHI", TimeSetOverrideDataSrcs.PatHi },
        { "PATLO", TimeSetOverrideDataSrcs.PatLo },
        { "PATNOT", TimeSetOverrideDataSrcs.PatNot },
    };

    private static readonly Dictionary<string, string> _validDataFmts = new(
        StringComparer.OrdinalIgnoreCase
    )
    {
        { "NR", TimeSetOverrideDataFmts.Nr },
        { "NR-2X", TimeSetOverrideDataFmts.Nr },
        { "RH", TimeSetOverrideDataFmts.Rh },
        { "RH-2X", TimeSetOverrideDataFmts.Rh },
        { "RL", TimeSetOverrideDataFmts.Rl },
        { "RL-2X", TimeSetOverrideDataFmts.Rl },
        { "STAY", TimeSetOverrideDataFmts.Stay },
        { "STAY-2X", TimeSetOverrideDataFmts.Stay },
    };

    public TimeSetOverrideBlockParserOutput Parse(
        IReadOnlyList<TimeSetOverrideRow> rows,
        TimeSetOverrideMetadata meta
    )
    {
        List<RowError> rowErrors = [];
        List<CellError> cellErrors = [];
        IEnumerable<IGrouping<string, TimeSetOverrideRow>> tsFileGroups = rows
            .GroupBy(r => r.TimeSetFile);

        List<TimeSetOverrideBlock> blocks = [];

        foreach (IGrouping<string, TimeSetOverrideRow> tsFileGroup in tsFileGroups)
        {
            IEnumerable<TimeSetOverrideRow> tsFiles = [.. tsFileGroup];
            OverrideBlockOutput blockOutput = GetBlockOutput(tsFileGroup.Key, tsFiles, meta);
            rowErrors.AddRange(blockOutput.RowErrors);
            cellErrors.AddRange(blockOutput.CellErrors);

            blocks.Add(blockOutput.Block);
        }
        return new(blocks, rowErrors, cellErrors);
    }

    private sealed record OverrideBlockOutput(
        TimeSetOverrideBlock Block,
        IReadOnlyList<RowError> RowErrors,
        IReadOnlyList<CellError> CellErrors
    );

    private static Result<Unit, RowError> ValidateDuplicateRow(
        List<TimeSetPinValueOverride> pinValueOverrides,
        TimeSetPinValueDto pinValDto,
        TimeSetOverrideRow row,
        TimeSetOverrideMetadata meta
    )
    {
        TimeSetPinValueOverride? duplicateRow = pinValueOverrides.FirstOrDefault(pv =>
                string.Equals(pv.TimeSet, pinValDto.TimeSet, StringComparison.OrdinalIgnoreCase)
                && string.Equals(pv.PinName, pinValDto.PinName, StringComparison.OrdinalIgnoreCase)
            );
        if (duplicateRow != null)
        {
            RowError rowError = new(
                BasicErrorType.E_TimeSetOverride_14,
                meta.SheetName,
                row.RowIndex,
                [pinValDto.TimeSet, pinValDto.PinName, $"{duplicateRow.RowIndex}"]
            );
            return Result<Unit, RowError>.Fail(rowError);
        }
        return Result<Unit, RowError>.Ok(Unit.Value);
    }

    private static bool HasConflictingValue(
        TimeSetFrequencyOverride existing,
        TimeSetFrequencyOverride incoming
    )
    {
        return !string.IsNullOrWhiteSpace(incoming.Value)
            && !string.Equals(existing.Value, incoming.Value, StringComparison.OrdinalIgnoreCase);
    }

    private static OverrideBlockOutput GetBlockOutput(
        string timeSetFileName,
        IEnumerable<TimeSetOverrideRow> tsFiles,
        TimeSetOverrideMetadata meta
    )
    {
        int colIndex = meta.GetColumnIndex(TimeSetOverrideSchema.FrequencyHeader);
        List<TimeSetFrequencyOverride> frequencyOverrides = [];
        List<TimeSetPinValueOverride> pinValueOverrides = [];
        List<RowError> rowErrors = [];
        List<CellError> cellErrors = [];
        foreach (TimeSetOverrideRow row in tsFiles)
        {
            Result<OverrideRowData, GetRowDataErrors> rowDataResult = GetRowData(row, meta);
            if (!rowDataResult.Success)
            {
                rowErrors.AddRange(rowDataResult.Error.RowErrors);
                cellErrors.AddRange(rowDataResult.Error.CellErrors);
                continue;
            }

            TimeSetFrequencyOverride? freqVar = rowDataResult.Value.OverrideOutput.Override;
            TimeSetPinValueDto pinValDto = rowDataResult.Value.PinValue;
            string variableName = freqVar != null ? freqVar.Name : string.Empty;

            Result<Unit, RowError> dupeValidateResult = ValidateDuplicateRow(
                pinValueOverrides,
                pinValDto,
                row,
                meta
            );
            if (!dupeValidateResult.Success)
            {
                rowErrors.Add(dupeValidateResult.Error);
                continue;
            }
            pinValueOverrides.Add(new(
                RowIndex: pinValDto.RowIndex,
                TimeSet: pinValDto.TimeSet,
                PinName: pinValDto.PinName,
                Setup: pinValDto.Setup,
                DataSrc: pinValDto.DataSrc,
                DataFmt: pinValDto.DataFmt,
                Variable: variableName
            ));

            if (freqVar == null)
            {
                continue;
            }

            TimeSetFrequencyOverride? targetFreq = frequencyOverrides.FirstOrDefault(f =>
                string.Equals(f.Name, freqVar.Name, StringComparison.OrdinalIgnoreCase)
            );

            if (targetFreq == null)
            {
                frequencyOverrides.Add(freqVar);
                continue;
            }

            if (!HasConflictingValue(targetFreq, freqVar))
            {
                continue;
            }

            string values = string.Join(", ", targetFreq.Value, freqVar.Value);
            CellError cellValError = new(
                BasicErrorType.E_TimeSetOverride_06,
                meta.SheetName,
                row.RowIndex,
                colIndex,
                [freqVar.Name, values]
            );
            cellErrors.Add(cellValError);

            // Replace old value
            frequencyOverrides.Remove(targetFreq);
            frequencyOverrides.Add(freqVar);
        }
        TimeSetOverrideBlock block = new(timeSetFileName, frequencyOverrides, pinValueOverrides);
        return new(block, rowErrors, cellErrors);
    }

    private sealed record OverrideRowData(
        FrequencyOverrideOutput OverrideOutput,
        TimeSetPinValueDto PinValue
    );

    private sealed record GetRowDataErrors(
        IReadOnlyList<RowError> RowErrors,
        IReadOnlyList<CellError> CellErrors
    );

    private static Result<OverrideRowData, GetRowDataErrors> GetRowData(
        TimeSetOverrideRow row,
        TimeSetOverrideMetadata meta
    )
    {
        List<RowError> currentRowErrors = [];
        List<CellError> currentCellErrors = [];
        Result<FrequencyOverrideOutput, CellError> rowFreqResult = GetFrequencyOverride(row, meta);
        if (!rowFreqResult.Success)
        {
            currentCellErrors.Add(rowFreqResult.Error);
        }

        Result<TimeSetPinValueDto, CombinedErrors> rowPinValueOutput = GetPinValueOverride(row, meta);
        if (!rowPinValueOutput.Success)
        {
            currentRowErrors.AddRange(rowPinValueOutput.Error.RowErrors);
            currentCellErrors.AddRange(rowPinValueOutput.Error.CellErrors);
        }

        if (!rowPinValueOutput.Success || !rowFreqResult.Success)
        {
            return Result<OverrideRowData, GetRowDataErrors>.Fail(new(currentRowErrors, currentCellErrors));
        }
        return Result<OverrideRowData, GetRowDataErrors>.Ok(new(rowFreqResult.Value, rowPinValueOutput.Value));
    }

    private sealed record FrequencyOverrideOutput(TimeSetFrequencyOverride? Override);

    private static Result<FrequencyOverrideOutput, CellError> GetFrequencyOverride(
        TimeSetOverrideRow row,
        TimeSetOverrideMetadata meta
    )
    {
        if (string.IsNullOrWhiteSpace(row.Frequency))
        {
            return Result<FrequencyOverrideOutput, CellError>.Ok(new(null));
        }

        Result<FrequencyVariableOutput, CellError> varResult = row.Frequency.Contains('=')
            ? GetVariableAndValue(row, row.Frequency, meta)
            : GetVariableOnly(row, row.Frequency, meta);
        if (!varResult.Success)
        {
            return Result<FrequencyOverrideOutput, CellError>.Fail(varResult.Error);
        }
        FrequencyVariableOutput freqVar = varResult.Value;
        return Result<FrequencyOverrideOutput, CellError>
            .Ok(new(new(freqVar.Name, freqVar.Value, row.RowIndex)));
    }

    private sealed record CombinedErrors(IReadOnlyList<RowError> RowErrors, IReadOnlyList<CellError> CellErrors);

    private sealed record TimeSetPinValueDto(int RowIndex,
        string TimeSet,
        string PinName,
        string Setup,
        string? DataSrc,
        string? DataFmt
    );

    private static Result<TimeSetPinValueDto, CombinedErrors> GetPinValueOverride(
        TimeSetOverrideRow row,
        TimeSetOverrideMetadata meta
    )
    {
        int dataFmtColIndex = meta.GetColumnIndex(TimeSetOverrideSchema.DataFmtHeader);

        List<CellError> cellErrors = CheckNotEmptyAndValid(row, meta);
        List<RowError> rowErrors = [];

        if (string.Equals(row.DataSrc, TimeSetOverrideDataSrcs.AllHi) && !string.Equals(row.DataFmt, TimeSetOverrideDataFmts.Rl))
        {
            CellError cellError = new(
                BasicErrorType.E_TimeSetOverride_11,
                meta.SheetName,
                row.RowIndex,
                dataFmtColIndex
            );
            cellErrors.Add(cellError);
        }
        if (string.Equals(row.DataSrc, TimeSetOverrideDataSrcs.AllLo) && !string.Equals(row.DataFmt, TimeSetOverrideDataFmts.Rh))
        {
            CellError cellError = new(
                BasicErrorType.E_TimeSetOverride_12,
                meta.SheetName,
                row.RowIndex,
                dataFmtColIndex
            );
            cellErrors.Add(cellError);
        }

        if (cellErrors.Count > 0 || rowErrors.Count > 0)
        {
            return Result<TimeSetPinValueDto, CombinedErrors>
                .Fail(new(rowErrors, cellErrors));
        }
        string dataSrc = GetValidString(_validDataSrcs, row.DataSrc);
        string dataFmt = GetValidString(_validDataFmts, row.DataFmt);
        return Result<TimeSetPinValueDto, CombinedErrors>.Ok(
            new(
                row.RowIndex,
                row.TimeSet,
                row.PinGroupName,
                _validSetups[row.Setup],
                dataSrc,
                dataFmt
            )
        );
    }

    private static string GetValidString(Dictionary<string, string> dict, string? input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return string.Empty;
        }
        if (dict.ContainsKey(input))
        {
            return dict[input];
        }
        return string.Empty;
    }

    private sealed record FrequencyVariableOutput(string Name, string Value);

    private static Result<FrequencyVariableOutput, CellError> GetVariableOnly(
        TimeSetOverrideRow row,
        string frequency,
        TimeSetOverrideMetadata meta
    )
    {
        Result<string, CellError> varResult = ValidateVariable(row, frequency, meta);
        if (!varResult.Success)
        {
            return Result<FrequencyVariableOutput, CellError>.Fail(varResult.Error);
        }
        return Result<FrequencyVariableOutput, CellError>.Ok(new(varResult.Value, string.Empty));
    }

    private static Result<FrequencyVariableOutput, CellError> GetVariableAndValue(
        TimeSetOverrideRow row,
        string frequency,
        TimeSetOverrideMetadata meta
    )
    {
        string[] args = frequency.Split(
            '=',
            StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries
        );

        int colIndex = meta.GetColumnIndex(TimeSetOverrideSchema.FrequencyHeader);

        if (args.Length != 2)
        {
            CellError cellError = new(
                BasicErrorType.E_TimeSetOverride_04,
                meta.SheetName,
                row.RowIndex,
                colIndex,
                [frequency]
            );
            return Result<FrequencyVariableOutput, CellError>.Fail(cellError);
        }

        string varName = args[0].Trim();
        Result<string, CellError> varResult = ValidateVariable(row, varName, meta);
        if (!varResult.Success)
        {
            return Result<FrequencyVariableOutput, CellError>.Fail(varResult.Error);
        }

        bool isFreqVal = args[1].Trim().TryConvertToFreq(out string freqVal);
        if (!isFreqVal)
        {
            CellError cellValError = new(
                BasicErrorType.E_TimeSetOverride_05,
                meta.SheetName,
                row.RowIndex,
                colIndex,
                [frequency]
            );
            return Result<FrequencyVariableOutput, CellError>.Fail(cellValError);
        }
        return Result<FrequencyVariableOutput, CellError>.Ok(new(varName, freqVal));
    }

    private static Result<string, CellError> ValidateVariable(
        TimeSetOverrideRow row,
        string variableName,
        TimeSetOverrideMetadata meta
    )
    {
        int colIndex = meta.GetColumnIndex(TimeSetOverrideSchema.FrequencyHeader);
        bool validVariable = Regex.IsMatch(variableName, VariableTemplate);
        if (!validVariable)
        {
            CellError cellError = new(
                BasicErrorType.E_TimeSetOverride_03,
                meta.SheetName,
                row.RowIndex,
                colIndex,
                [variableName]
            );
            return Result<string, CellError>.Fail(cellError);
        }
        return Result<string, CellError>.Ok(variableName);
    }

    private const string VariableTemplate = @"^[A-Za-z][A-Za-z0-9_]*$";

    private static List<CellError> CheckNotEmptyAndValid(
        TimeSetOverrideRow row,
        TimeSetOverrideMetadata meta
    )
    {
        List<CellError> cellErrors = [];
        int setupColIndex = meta.GetColumnIndex(TimeSetOverrideSchema.SetupHeader);
        int dataSrcColIndex = meta.GetColumnIndex(TimeSetOverrideSchema.DataSrcHeader);
        int dataFmtColIndex = meta.GetColumnIndex(TimeSetOverrideSchema.DataFmtHeader);
        if (string.IsNullOrWhiteSpace(row.Setup))
        {
            CellError cellError = new(
                BasicErrorType.E_TimeSetOverride_07,
                meta.SheetName,
                row.RowIndex,
                setupColIndex
            );
            cellErrors.Add(cellError);
        }
        if (!_validSetups.ContainsKey(row.Setup))
        {
            CellError cellError = new(
                BasicErrorType.E_TimeSetOverride_08,
                meta.SheetName,
                row.RowIndex,
                setupColIndex,
                [row.Setup]
            );
            cellErrors.Add(cellError);
        }
        if (!string.IsNullOrWhiteSpace(row.DataSrc) && !_validDataSrcs.ContainsKey(row.DataSrc))
        {
            CellError cellError = new(
                BasicErrorType.E_TimeSetOverride_09,
                meta.SheetName,
                row.RowIndex,
                dataSrcColIndex,
                [row.DataSrc]
            );
            cellErrors.Add(cellError);
        }
        if (!string.IsNullOrWhiteSpace(row.DataFmt) && !_validDataFmts.ContainsKey(row.DataFmt))
        {
            CellError cellError = new(
                BasicErrorType.E_TimeSetOverride_10,
                meta.SheetName,
                row.RowIndex,
                dataFmtColIndex,
                [row.DataFmt]
            );
            cellErrors.Add(cellError);
        }
        return cellErrors;
    }
}
