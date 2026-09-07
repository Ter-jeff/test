using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Constants;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Results;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Applier;

public sealed class TimeSetOverrideApplier
{
    private static readonly Regex _variableRegex = new(
        @"_[A-Za-z][A-Za-z0-9_]*",
        RegexOptions.Compiled
    );

    private static readonly Dictionary<string, Action<TimingRow, string>> _clockModeHandlers = new(
        StringComparer.OrdinalIgnoreCase
    )
    {
        [TimeSetOverrideSetups.Clock2X] = (row, baseVar) =>
        {
            row.PinGrpSetup = TimeSetOverrideSetups.Clock2X;
            row.PinGrpClockPeriod = $@"=(1/{baseVar})*2";
            row.DriveOn = $@"=(1/{baseVar})*0.5";
            row.DriveData = $@"=(1/{baseVar})*1";
            row.DriveReturn = $@"=(1/{baseVar})*1.5";
            row.DriveOff = $@"=(1/{baseVar})*2";
        },
        [TimeSetOverrideSetups.Clock] = (row, baseVar) =>
        {
            row.PinGrpSetup = TimeSetOverrideSetups.Clock;
            row.PinGrpClockPeriod = $@"=1/{baseVar}";
            row.DriveOn = "0";
            row.DriveData = $@"=1/(2*{baseVar})";
            row.DriveReturn = $@"=1/{baseVar}";
            row.DriveOff = "0";
        },
        [TimeSetOverrideSetups.Io] = (row, baseVar) =>
        {
            row.PinGrpSetup = TimeSetOverrideSetups.Io;
            row.DriveData = "0";
            row.DriveReturn = "";
        },
    };

    public Result<Unit, IReadOnlyList<RowError>> Apply(
        ComTimeSetBasicSheet tsSheet,
        TimeSetOverrideBlock block,
        string timeSetPath
    )
    {
        string txtName = Path.GetFileName(timeSetPath);

        // Modify rows
        List<RowError> rowErrors = [];

        // Modify VAR Definition
        IReadOnlyList<RowError> modifiedErrors = ModifyVarDefDictionary(tsSheet, block);
        rowErrors.AddRange(modifiedErrors);

        // Group by TimeSet column in TimeSet TXT file
        IEnumerable<IGrouping<string, TimeSetPinValueOverride>> pinValOverridesGroups = block
            .PinValueOverrides
            .GroupBy(pv => pv.TimeSet);

        foreach (IGrouping<string, TimeSetPinValueOverride> pinOverridesGroup in pinValOverridesGroups)
        {
            string timeSetName = pinOverridesGroup.Key;
            IEnumerable<TimeSetPinValueOverride> pinOverrides = [.. pinOverridesGroup];
            TSet? targetTs = tsSheet.Rows.FirstOrDefault(r =>
                string.Equals(r.Name, timeSetName, StringComparison.OrdinalIgnoreCase)
            );
            if (targetTs == null)
            {
                RowError rowError = new(
                    BasicErrorType.E_TimeSetOverride_21,
                    TimeSetOverrideSchema.SheetName,
                    pinOverrides.First().RowIndex,
                    [timeSetName, txtName]
                );
                rowErrors.Add(rowError);
                continue;
            }
            IReadOnlyList<RowError> errors = ModifyPinValuesByGroup(
                pinOverrides,
                targetTs,
                txtName,
                modifiedErrors
            );
            rowErrors.AddRange(errors);
        }
        if (rowErrors.Count > 0)
        {
            return Result<Unit, IReadOnlyList<RowError>>.Fail(rowErrors);
        }
        return Result<Unit, IReadOnlyList<RowError>>.Ok(Unit.Value);
    }

    private IReadOnlyList<RowError> ModifyPinValuesByGroup(
        IEnumerable<TimeSetPinValueOverride> pinOverrides,
        TSet targetTs,
        string txtName,
        IReadOnlyList<RowError> modifiedErrors
    )
    {
        List<RowError> rowErrors = [];
        foreach (TimeSetPinValueOverride pinOverride in pinOverrides)
        {
            TimingRow? targetTsRow = targetTs.TimingRows
                .FirstOrDefault(r => r.PinGrpName == pinOverride.PinName);
            if (targetTsRow == null)
            {
                RowError rowError = new(
                    BasicErrorType.E_TimeSetOverride_15,
                    TimeSetOverrideSchema.SheetName,
                    pinOverride.RowIndex,
                    [pinOverride.PinName, pinOverride.TimeSet, txtName]
                );
                rowErrors.Add(rowError);
                continue;
            }

            // If variable modify occurred on the same row, stop apply value
            if (modifiedErrors.Any(e => e.RowIndex == pinOverride.RowIndex))
            {
                continue;
            }

            Result<FinalOverrideOutput, RowError> finalResult = GetFinalOverrideOutput(
                targetTsRow,
                pinOverride,
                txtName
            );
            if (!finalResult.Success)
            {
                rowErrors.Add(finalResult.Error);
                continue;
            }

            FinalOverrideOutput output = finalResult.Value;

            targetTsRow.DataSrc = output.DataSrc;
            targetTsRow.DataFmt = output.DataFmt;
            output.Apply(targetTsRow, output.BaseVariable);
        }
        return rowErrors;
    }

    private sealed record FinalOverrideOutput(
        string BaseVariable,
        string DataSrc,
        string DataFmt,
        Action<TimingRow, string> Apply
    );

    private static Result<FinalOverrideOutput, RowError> GetFinalOverrideOutput(
        TimingRow targetTsRow,
        TimeSetPinValueOverride pinOverride,
        string txtName
    )
    {
        Result<string, RowError> baseVarResult = GetBaseVariableResult(
            targetTsRow,
            pinOverride,
            txtName
        );
        if (!baseVarResult.Success)
        {
            return Result<FinalOverrideOutput, RowError>.Fail(baseVarResult.Error);
        }
        string baseVar = baseVarResult.Value;

        // Get clock mode apply method
        if (!_clockModeHandlers.TryGetValue(pinOverride.Setup, out Action<TimingRow, string>? apply))
        {
            // This error should never happen because it should be removed at TimeSetOverrideReader.
            RowError rowError = new(
                BasicErrorType.E_TimeSetOverride_18,
                TimeSetOverrideSchema.SheetName,
                pinOverride.RowIndex,
                [pinOverride.Setup, txtName, $"{targetTsRow.RowNum}"]
            );
            return Result<FinalOverrideOutput, RowError>.Fail(rowError);
        }

        string finalSrc = string.IsNullOrWhiteSpace(pinOverride.DataSrc)
            ? targetTsRow.DataSrc
            : pinOverride.DataSrc;
        string finalFmt = string.IsNullOrWhiteSpace(pinOverride.DataFmt)
            ? targetTsRow.DataFmt
            : pinOverride.DataFmt;

        if (IsAllHButNoRl(finalSrc, finalFmt))
        {
            RowError rowError = new(
                BasicErrorType.E_TimeSetOverride_19,
                TimeSetOverrideSchema.SheetName,
                pinOverride.RowIndex,
                [txtName, $"{targetTsRow.RowNum}"]
            );
            return Result<FinalOverrideOutput, RowError>.Fail(rowError);
        }

        if (IsAllLoButNoRh(finalSrc, finalFmt))
        {
            RowError rowError = new(
                BasicErrorType.E_TimeSetOverride_20,
                TimeSetOverrideSchema.SheetName,
                pinOverride.RowIndex,
                [txtName, $"{targetTsRow.RowNum}"]
            );
            return Result<FinalOverrideOutput, RowError>.Fail(rowError);
        }

        if (string.Equals(pinOverride.Setup, TimeSetOverrideSetups.Clock2X))
        {
            finalFmt = $"{finalFmt}-2X";
        }

        return Result<FinalOverrideOutput, RowError>.Ok(new(baseVar, finalSrc, finalFmt, apply));
    }

    private static bool IsAllHButNoRl(string dataSrc, string dataFmt)
    {
        bool isAllHi = string.Equals(dataSrc, TimeSetOverrideDataSrcs.AllHi, StringComparison.OrdinalIgnoreCase);
        bool isRl = string.Equals(dataFmt, TimeSetOverrideDataFmts.Rl, StringComparison.OrdinalIgnoreCase)
            || string.Equals(dataFmt, TimeSetOverrideDataFmts.Rl2x, StringComparison.OrdinalIgnoreCase);
        return isAllHi && !isRl;
    }

    private static bool IsAllLoButNoRh(string dataSrc, string dataFmt)
    {
        bool isAllLo = string.Equals(dataSrc, TimeSetOverrideDataSrcs.AllLo, StringComparison.OrdinalIgnoreCase);
        bool isRh = string.Equals(dataFmt, TimeSetOverrideDataFmts.Rh, StringComparison.OrdinalIgnoreCase)
            || string.Equals(dataFmt, TimeSetOverrideDataFmts.Rh2x, StringComparison.OrdinalIgnoreCase);
        return isAllLo && !isRh;
    }

    private static IReadOnlyList<RowError> ModifyVarDefDictionary(
        ComTimeSetBasicSheet tsSheet,
        TimeSetOverrideBlock block
    )
    {
        List<RowError> rowErrors = [];
        ComTimeSetBasic? firstTsBasic = GetFirstTimeSetBasic(tsSheet);
        if (firstTsBasic == null)
        {
            return rowErrors;
        }
        foreach (TimeSetFrequencyOverride freqOverride in block.FrequencyOverrides)
        {
            bool isValueEmpty = string.IsNullOrWhiteSpace(freqOverride.Value);
            bool isValueValid = double.TryParse(freqOverride.Value, out double newVal);
            bool isVariableExist = firstTsBasic.SubCommentVariable.ContainsKey(freqOverride.Name);
            if (!isVariableExist && isValueEmpty)
            {
                RowError rowError = new(
                    BasicErrorType.E_TimeSetOverride_22,
                    TimeSetOverrideSchema.SheetName,
                    freqOverride.RowIndex,
                    [freqOverride.Name]
                );
                rowErrors.Add(rowError);
                continue;
            }
            if (!isValueValid)
            {
                continue;
            }
            if (isVariableExist)
            {
                firstTsBasic.SubCommentVariable[freqOverride.Name] = newVal;
                continue;
            }
            firstTsBasic.SubCommentVariable.Add(freqOverride.Name, newVal);
            firstTsBasic.SubContextVariable.Add(freqOverride.Name);
        }
        return rowErrors;
    }

    private static ComTimeSetBasic? GetFirstTimeSetBasic(ComTimeSetBasicSheet tsSheet)
    {
        foreach (TSet row in tsSheet.Rows)
        {
            if (row is not ComTimeSetBasic basicTs)
            {
                throw new Exception("");
            }
            return basicTs;
        }
        return null;
    }

    private static Result<string, RowError> GetBaseVariableResult(
        TimingRow row,
        TimeSetPinValueOverride pinOverride,
        string txtName
    )
    {
        if (!string.IsNullOrWhiteSpace(pinOverride.Variable))
        {
            return Result<string, RowError>.Ok($"_{pinOverride.Variable}");
        }
        List<string> candidates = GetVariableNames(row.PinGrpClockPeriod);
        if (candidates.Count == 0)
        {
            RowError rowError = new(
                BasicErrorType.E_TimeSetOverride_16,
                TimeSetOverrideSchema.SheetName,
                pinOverride.RowIndex,
                [row.PinGrpClockPeriod, $"{row.RowNum}", txtName]
            );
            return Result<string, RowError>.Fail(rowError);
        }
        if (candidates.Count > 1)
        {
            RowError rowError = new(
                BasicErrorType.E_TimeSetOverride_17,
                TimeSetOverrideSchema.SheetName,
                pinOverride.RowIndex,
                [row.PinGrpClockPeriod, $"{row.RowNum}", txtName]
            );
            return Result<string, RowError>.Fail(rowError);
        }
        return Result<string, RowError>.Ok(candidates[0]);
    }

    private static List<string> GetVariableNames(string formula)
    {
        return _variableRegex
            .Matches(formula)
            .Select(m => m.Value)
            .Distinct(StringComparer.Ordinal)
            .ToList();
    }
}
