using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTableRow;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.BinNumber;
using TestPlanLib.Singleton;
using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTable
{
    public abstract class BlockBinTableGeneratorBase
    {
        protected HardIpInputData HardIpInputData { get; }
        protected BinTableSheet HardIpBinTableSheet;
        protected List<HardIpPattern> PatternList;
        protected List<string> DuplicateParameter;
        protected BinTableRowGeneratorBase BinTableRowGenerator = null;

        protected BlockBinTableGeneratorBase(HardIpInputData hardIpInputData, BinTableSheet hardIpBinTableSheet, List<HardIpPattern> patternList, List<string> duplicateParameter)
        {
            HardIpInputData = hardIpInputData;
            HardIpBinTableSheet = hardIpBinTableSheet;
            PatternList = patternList;
            DuplicateParameter = duplicateParameter;
        }

        protected virtual string GetVoltageForBinTable(HardIpPattern pattern)
        {
            bool isSpecialInstance = HardIpConstData.RegInsInPatt.IsMatch(pattern.Pattern.GetLastPayload());
            return isSpecialInstance ? "NV" : "";
        }

        public virtual void GenerateBinTableRows(bool mergedHln = false)
        {
            var extraBinTables = new List<BinTableRow>();
            var voltageSeq = new List<string> { HardIpConstData.LabelNv };

            if (GetType() != typeof(RtosBinTableGenerator))
            {
                voltageSeq.Add(HardIpConstData.LabelLv);
                voltageSeq.Add(HardIpConstData.LabelHv);
            }

            foreach (HardIpPattern pattern in PatternList)
            {
                if (pattern.IsIgnorePatBinOut() ||
                    Regex.IsMatch(pattern.Pattern.GetLastPayload(), "^cz_", RegexOptions.IgnoreCase) ||
                    (pattern.Pattern.IsMultiTimeDomain() && !pattern.Pattern.IsFullMtdPattern()) ||
                    (HardIpConstData.RegOpcodeInPatt.IsMatch(pattern.Pattern.RealPatternName) && string.IsNullOrEmpty(pattern.Failflag)))
                {
                    continue;
                }

                List<BinTableRow> binTableList = new List<BinTableRow>();
                List<BinTableRow> extraBintableList = new List<BinTableRow>();
                string voltage = GetVoltageForBinTable(pattern);

                if (!Regex.IsMatch(pattern.SheetName, "DCTEST_IDS", RegexOptions.IgnoreCase) &&
                    !Regex.IsMatch(pattern.SheetName, "RTOS_IDS", RegexOptions.IgnoreCase) &&
                    !VbtFunctionLibShared.EfuseReadFunctionList.Exists(f => f.Equals(pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    BuildDeviceBinTableRows(pattern, voltage, voltageSeq, mergedHln, binTableList, ref extraBintableList);
                }
                if (!binTableList.Any())
                {
                    BinTableRow binTableRow = BinTableRowGenerator.GenBinTableRow(pattern, voltage);
                    binTableList = new List<BinTableRow> { binTableRow };
                }
                if (!binTableList.Exists(p => DuplicateParameter.Contains(p.Name, StringComparer.OrdinalIgnoreCase)))
                {
                    foreach (BinTableRow item in binTableList)
                    {
                        if (item != null)
                        {
                            HardIpBinTableSheet.AddRow(item);
                            DuplicateParameter.Add(item.Name);
                        }
                    }
                }
            }

            //Check if enable the adaptive cooling and generate it.
            if (HardIpInputData.ConfigData.NeedGenerateTmpsBin)
            {
                HardIpInputData.ConfigData.NeedGenerateTmpsBin = false;
                var binTmpsMonitorRow = new BinTableRow
                {
                    Name = "Bin_TMPS_Monitor",
                    ItemList = "F_TMPS_Monitor",
                    Op = "OR",
                    Result = BinTableRow.ResultFailStop,
                    Items = new List<string> { "T" }
                };

                if (!BinNumberSingleton.Instance.CheckDuplicateBinRows(binTmpsMonitorRow))
                {
                    BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("TMPS", "", "", binTmpsMonitorRow);
                    binTmpsMonitorRow.Sort = binNumInfo.SoftBin.ToString("G15");
                    binTmpsMonitorRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                    extraBinTables.Add(binTmpsMonitorRow);
                }
            }

            foreach (BinTableRow item in extraBinTables)
            {
                if (item != null)
                {
                    HardIpBinTableSheet.AddRow(item);
                }
            }
        }

        private void BuildDeviceBinTableRows(HardIpPattern pattern, string voltage, List<string> voltageSeq, bool mergedHln, List<BinTableRow> binTableList, ref List<BinTableRow> extraBintableList)
        {
            if (!(LocalSpecs.Options.Device == EnumDevice.AP || LocalSpecs.Options.Device == EnumDevice.RF || LocalSpecs.Options.Device == EnumDevice.LCD))
            {
                return;
            }

            if (voltage != "")
            {
                return;
            }

            if (!mergedHln)
            {
                if (LocalSpecs.Options.Device == EnumDevice.AP)
                {
                    binTableList.AddRange(voltageSeq.Select(x => BinTableRowGenerator.GenBinTableRow(pattern, x)));
                    CombineBinRowByVoltageSeq(binTableList);
                }
                else if (LocalSpecs.Options.Device == EnumDevice.RF)
                {
                    BinTableRow binTableRow = BinTableRowGenerator.GenBinTableRow(pattern, voltage);
                    binTableList.Add(binTableRow);
                    extraBintableList = ExtendBinTableRows(binTableRow);
                }
                else
                {
                    BinTableRow binTableRow = BinTableRowGenerator.GenBinTableRow(pattern, voltage);
                    List<BinTableRow> split = SplitBinRowByVoltageSeq(binTableRow, voltageSeq, ref extraBintableList);
                    binTableList.Clear();
                    binTableList.AddRange(split);
                }
            }
            else
            {
                BinTableRow binTableRow = BinTableRowGenerator.GenBinTableRow(pattern, voltage);
                binTableList.Add(binTableRow);
                extraBintableList = ExtendBinTableRows(binTableRow);
            }
        }

        private List<BinTableRow> SplitBinRowByVoltageSeq(BinTableRow binRow, List<string> volSeq, ref List<BinTableRow> extraBintableRows)
        {
            var result = new List<BinTableRow>();
            foreach (string vol in volSeq)
            {
                BinTableRow copy = binRow.Copy();
                copy.Op = "OR";
                copy.Items = SelectFlagStateByVoltage(vol);
                copy.Sort = SelectBinNunmberByVoltage(copy.Sort, vol);
                BinTableRow noBinRow = copy.Copy();
                noBinRow.Name = copy.Name + "_" + vol;
                noBinRow.ItemList = SelectFlagByVoltage(copy.ItemList, vol);
                noBinRow.Items = new List<string> { "T" };
                extraBintableRows.Add(noBinRow);
                result.Add(copy);
            }
            return result;
        }

        private void CombineBinRowByVoltageSeq(List<BinTableRow> binRows)
        {
            IEnumerable<string> allFlags = binRows.SelectMany(x => x.ItemList.Split(','));
            string combinedAllFlags = string.Join(",", allFlags);
            binRows.ForEach(x => x.ItemList = combinedAllFlags);
        }

        private string SelectBinNunmberByVoltage(string binNum, string vol)
        {
            int result = int.Parse(binNum);
            if (vol.Equals(HardIpConstData.LabelLv, StringComparison.OrdinalIgnoreCase))
            {
                result += 1000;
            }

            if (vol.Equals(HardIpConstData.LabelHv, StringComparison.OrdinalIgnoreCase))
            {
                result += 2000;
            }

            return result.ToString();
        }

        private List<string> SelectFlagStateByVoltage(string vol = "")
        {
            switch (vol)
            {
                case HardIpConstData.LabelHv:
                    return new List<string> { "", "T", "" };
                case HardIpConstData.LabelLv:
                    return new List<string> { "", "", "T" };
                case HardIpConstData.LabelNv:
                    return new List<string> { "T", "", "" };
                default:
                    return new List<string> { "T", "T", "T" };
            }
        }

        private string SelectFlagByVoltage(string flags, string voltage)
        {
            List<string> flagsets = flags.Split(',').ToList();
            switch (voltage)
            {
                case HardIpConstData.LabelHv:
                    return flagsets[1];
                case HardIpConstData.LabelLv:
                    return flagsets[2];
                case HardIpConstData.LabelNv:
                    return flagsets[0];
                default:
                    return "";
            }
        }

        private List<BinTableRow> ExtendBinTableRows(BinTableRow row)
        {
            var results = new List<BinTableRow>();
            BinTableRow noBinRow = row.Copy();
            noBinRow.Name = row.Name + "_" + HardIpConstData.LabelLv;
            noBinRow.ItemList = SelectFlagByVoltage(row.ItemList, HardIpConstData.LabelLv);
            noBinRow.Items = new List<string> { "T" };
            results.Add(noBinRow);
            noBinRow = row.Copy();
            noBinRow.Name = row.Name + "_" + HardIpConstData.LabelHv;
            noBinRow.ItemList = SelectFlagByVoltage(row.ItemList, HardIpConstData.LabelHv);
            noBinRow.Items = new List<string> { "T" };
            results.Add(noBinRow);
            noBinRow = row.Copy();
            noBinRow.Name = row.Name + "_" + HardIpConstData.LabelNv;
            noBinRow.ItemList = SelectFlagByVoltage(row.ItemList, HardIpConstData.LabelNv);
            noBinRow.Items = new List<string> { "T" };
            results.Add(noBinRow);
            return results;
        }

    }
}
