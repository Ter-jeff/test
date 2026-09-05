using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.Static;
using Automation.Utility.Atpg.Data;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

using TestPlanLib.BinNumber;
using TestPlanLib.Singleton;

namespace Automation.Utility.Atpg
{
    public static class AtpgBinOutService
    {
        private static readonly Regex _regexEndsWithHiLoVol = new Regex("_(HV|LV)+$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static List<AtpgBinOutData> GetAtpgBinOutDatas(List<string> binItems, BinCutFinalInstanceRow dataRow)
        {
            var binOutItems = new List<AtpgBinOutData>();
            string domain = dataRow.Domain;
            string block = dataRow.Block;
            string binOutStage = dataRow.GetBinOutStage();
            bool isByPassBinOut = dataRow.BinCutInstanceRow.IsBypassBinOut;
            string testPlanSheetName = dataRow.BinCutInstanceRow.SheetName;
            string subFlowSheetName = dataRow.GetSubFlowName();
            bool isNeedEvsDeferredBinOut = TestPlanStatic.MainFlowSheet?.EvsDeferSubFlowSheets?.Contains($"{testPlanSheetName}:{subFlowSheetName}") ?? false;

            foreach (string binItem in binItems)
            {
                string binOp = "AND";
                if (binItem.Contains("&&") || binItem.Contains("||"))
                {
                    List<string> flagList = binItem.Replace("&&", ",").Replace("||", ",").Split(',').ToList();
                    if (binItem.Contains("&&") && binItem.Contains("||"))
                    {
                        binOp = binItem.IndexOf("&&", StringComparison.Ordinal) <= binItem.IndexOf("||", StringComparison.Ordinal) ? "AND" : "OR";
                    }
                    else
                    {
                        binOp = binItem.IndexOf("&&", StringComparison.Ordinal) != -1 ? "AND" : "OR";
                    }
                    binOutItems.Add(new AtpgBinOutData(domain, block, binOp, binOutStage, isByPassBinOut, false, isNeedEvsDeferredBinOut, flagList));
                }
                else
                {
                    if (_regexEndsWithHiLoVol.IsMatch(binItem) && string.IsNullOrEmpty(dataRow.BinCutInstanceRow.PinGroupBinoutFlag))
                    {
                        string flagName = _regexEndsWithHiLoVol.Replace(binItem, "");
                        binOutItems.Add(new AtpgBinOutData(domain, block, binOp, binOutStage, isByPassBinOut, true, isNeedEvsDeferredBinOut, [flagName]));
                    }
                    binOutItems.AddRange(dataRow.GetFlagWithPinFail(binItem).Select(x => new AtpgBinOutData(domain, block, binOp, binOutStage, isByPassBinOut, false, isNeedEvsDeferredBinOut, x)));
                }
            }
            return binOutItems;
        }

        public static List<BinTableRow> GetBinTableRow(AtpgBinOutData atpgBinOutData, string module)
        {
            var binTableRows = new List<BinTableRow>();
            List<(List<string>, string)> flagsVoltagePairs = new List<(List<string>, string)>();

            if (atpgBinOutData.IsHvLvBinOut)
            {
                string flagName = atpgBinOutData.FlagList.FirstOrDefault() ?? "";
                flagsVoltagePairs.Add(([flagName + "_HV", flagName + "_LV"], "_HLV"));
                flagsVoltagePairs.Add(([flagName + "_HV"], "_HV"));
                flagsVoltagePairs.Add(([flagName + "_LV"], "_LV"));
            }
            else
            {
                flagsVoltagePairs.Add((atpgBinOutData.FlagList, ""));
            }

            foreach ((List<string>, string) flagsVoltagePair in flagsVoltagePairs)
            {
                var bin = new BinTableRow
                {
                    Name = GetBinName(flagsVoltagePair.Item1, flagsVoltagePair.Item2),
                    ItemList = string.Join(",", flagsVoltagePair.Item1),
                    Op = atpgBinOutData.BinOp,
                    Items = Enumerable.Repeat("T", flagsVoltagePair.Item1.Count).ToList()
                };
                BinNumResult binNumInfoHlv = BinNumberSingleton.Instance.GetBinInfo(module, atpgBinOutData.Domain, atpgBinOutData.Block, bin);
                bin.Sort = binNumInfoHlv.SoftBin.ToString("G15");
                bin.Bin = binNumInfoHlv.BinNumInfo.HardBin.ToString("G15");
                bin.Result = binNumInfoHlv.BinNumInfo.Status;

                if (atpgBinOutData.IsNeedEvsDeferredBinOut)
                {
                    BinTableRow deferBinTable = bin.Copy();
                    deferBinTable.ItemList += ",F_EVS_Defer";
                    deferBinTable.Items.Add("T");
                    BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("EVSDEFER", atpgBinOutData.Domain, atpgBinOutData.Block, deferBinTable);
                    deferBinTable.Sort = binNumInfo.SoftBin.ToString("G15");
                    deferBinTable.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                    bin.ItemList = "";
                    binTableRows.Add(deferBinTable);
                }
                binTableRows.Add(bin);
            }
            return binTableRows;
        }

        public static List<FlowRow> GetBinFlowRow(AtpgBinOutData atpgBinOutData)
        {
            var bin = new FlowRow();
            bin.Parameter = GetBinName(atpgBinOutData.FlagList, atpgBinOutData.IsHvLvBinOut ? "_HLV" : "");
            bin.Opcode = OpCode.BinTable;
            bin.BinTableFlagCount = atpgBinOutData.FlagList.Count;
            return [bin];
        }

        public static List<string> GetBinNameList(AtpgBinOutData atpgBinOutData)
        {
            return [GetBinName(atpgBinOutData.FlagList, atpgBinOutData.IsHvLvBinOut ? "_HLV" : "")];
        }

        public static string GetBinName(List<string> flagList, string voltage, string prefix = "Bin_")
        {
            if (!string.IsNullOrEmpty(voltage))
            {
                flagList = [.. flagList.Select(x => _regexEndsWithHiLoVol.Replace(x, "")).Distinct()];
            }
            string flagsCombination = string.Join("_", flagList.Select(x => Regex.Replace(x, "^F_", "")));
            return $"{prefix}{flagsCombination}{voltage}";
        }

        public static IEqualityComparer<T> GetBinOutItemsComparer<T>()
        {
            if (typeof(T) == typeof(BinTableRow))
            {
                var comparer = EqualityComparer<BinTableRow>.Create(
                    (x, y) =>
                        string.Equals(x.Name, y.Name, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(x.ItemList, y.ItemList, StringComparison.OrdinalIgnoreCase),
                    x => HashCode.Combine(
                        StringComparer.OrdinalIgnoreCase.GetHashCode(x.Name ?? string.Empty),
                        StringComparer.OrdinalIgnoreCase.GetHashCode(x.ItemList ?? string.Empty)));

                return (IEqualityComparer<T>)(object)comparer;
            }

            if (typeof(T) == typeof(FlowRow))
            {
                var comparer = EqualityComparer<FlowRow>.Create(
                    (x, y) =>
                        string.Equals(x.Parameter, y.Parameter, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(x.Enable, y.Enable, StringComparison.OrdinalIgnoreCase),
                    x => HashCode.Combine(
                        StringComparer.OrdinalIgnoreCase.GetHashCode(x.Parameter ?? string.Empty),
                        StringComparer.OrdinalIgnoreCase.GetHashCode(x.Enable ?? string.Empty)));

                return (IEqualityComparer<T>)(object)comparer;
            }

            if (typeof(T) == typeof(string))
            {
                return (IEqualityComparer<T>)(object)
                    StringComparer.OrdinalIgnoreCase;
            }

            return EqualityComparer<T>.Default;
        }
    }
}
