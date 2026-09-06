using System;
using System.Collections.Generic;
using System.IO;

using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib.Base
{
    internal static class SiteInfoHelpers
    {
        public static double GetVoltageOfEvaluation(PowerZone powerZone)
        {
            int step = powerZone.IsAdjust ? powerZone.FinalStep : powerZone.Step;
            return powerZone.PossibleSteps[step].Lvcc;
        }

        public static double GetEfuseGb(BinningTable binningTable, string mode, EnumJob enumJob)
        {
            int lvVddBinGbIdx = binningTable.GetGbIdx(enumJob);
            if (BinCutConfig.FlagT0TxHotFormat || BinCutConfig.FlagT0TxRoomFormat)
            {
                lvVddBinGbIdx = BinCutConfig.FlagT0TxHotFormat ? binningTable.HtolGbHotIdx : binningTable.HtolGbRoomIdx;
            }

            for (int rowIdx = 0; rowIdx < binningTable.Rows.Count; rowIdx++)
            {
                string modeInVddBin = binningTable.Rows[rowIdx].RowData[binningTable.ModeIdx];
                if (mode.EqualsIgnoreCase(modeInVddBin))
                {
                    _ = double.TryParse(binningTable.Rows[rowIdx].RowData[lvVddBinGbIdx], out double dGbTmp);
                    return dGbTmp;
                }
            }
            return 0;
        }

        public static double GetLvcc(PowerZone powerZone, string mode, List<EFuseRow> eFuseRows)
        {
            if (powerZone.Pin.Contains("SRAM"))//For sram as core power
            {
                return eFuseRows.Exists(x => x.Name.EqualsIgnoreCase(mode)) ? GetEfuseLvcc(mode, eFuseRows) : powerZone.GetFinalLvcc();
            }

            if (BinCutConfig.FlagT0TxHotFormat || BinCutConfig.FlagT0TxRoomFormat)
            {
                return
                    BinCutData.BinCutFlowTables.Find(
                        x => x.JobName.ContainsIgnoreCase(BinCutConfig.FlagT0TxHotFormat ? "T0TX_HOT" : "T0TX_ROOM"))!
                        .GetEvaluateModes()
                        .Contains(mode)
                        ? powerZone.GetFinalLvcc()
                        : GetEfuseLvcc(mode, eFuseRows);
            }

            return BinCutData.BinCutFlowTables.Find(x => x.FinalJob.Contains(BinCutData.Job.JobType.ToString()))!.GetEvaluateModes().Contains(mode) ? powerZone.GetFinalLvcc() : GetEfuseLvcc(mode, eFuseRows);
            //return (powerZone.IdsValue == 0 || BinCutConfig.Is_BinCutJob_for_StepSearch == false) ? GetEfuseLvcc(mode) : powerZone.GetFinalLvcc();
        }

        public static double GetEfuseLvcc(string mode, List<EFuseRow> eFuseRows)
        {
            double productValue = 0;
            if (eFuseRows.Exists(x => x.Name.EqualsIgnoreCase(mode)))
            {
                EFuseRow efuse = eFuseRows.Find(x => x.Name.EqualsIgnoreCase(mode))!;
                productValue = efuse.Value - efuse.Gb;
            }
            return productValue;
        }

        public static double GetEfuseProductValue(string mode, List<EFuseRow> eFuseRows)
        {
            double productValue = 0;
            if (eFuseRows.Exists(x => x.Name.EqualsIgnoreCase(mode)))
            {
                EFuseRow efuse = eFuseRows.Find(x => x.Name.EqualsIgnoreCase(mode))!;
                productValue = efuse.Value;
            }
            return productValue;
        }

        public static double GetProductValue(PowerZone powerZone, string mode, StreamWriter streamWriter, List<EFuseRow> eFuseRows)
        {
            try
            {
                if (powerZone.Pin.Contains("SRAM"))//For sram as core power
                {
                    return eFuseRows.Exists(x => x.Name.EqualsIgnoreCase(mode)) ? GetEfuseProductValue(mode, eFuseRows) : powerZone.GetFinalProductValue(streamWriter);
                }

                if (BinCutConfig.FlagT0TxHotFormat || BinCutConfig.FlagT0TxRoomFormat)
                {
                    return
                        BinCutData.BinCutFlowTables.Find(
                            x => x.JobName.ContainsIgnoreCase(BinCutConfig.FlagT0TxHotFormat ? "T0TX_HOT" : "T0TX_ROOM"))!
                            .GetEvaluateModes()
                            .Contains(mode)
                            ? powerZone.GetFinalProductValue(streamWriter)
                            : GetEfuseProductValue(mode, eFuseRows);
                }

                return BinCutData.BinCutFlowTables.Find(x => x.FinalJob.Contains(BinCutData.Job.JobType.ToString()))!.GetEvaluateModes().Contains(mode) ? powerZone.GetFinalProductValue(streamWriter) : GetEfuseProductValue(mode, eFuseRows);
            }
            catch (Exception)
            {
                return -1;
            }
        }

        public static List<AdjustVddBinningRow> GetAdjustVddBinningRows(int site, List<string> corePowers, List<PowerZone> powerZones)
        {
            var adjustVddBinningRows = new List<AdjustVddBinningRow>();
            var types = new List<string> { "BinCut", "EQN", "CP", "VDD", "IDS" };
            foreach (PowerZone power in powerZones)
            {
                if (!corePowers.Exists(x => x.EqualsIgnoreCase(power.PinMode)))
                {
                    continue;
                }

                foreach (string type in types)
                {
                    var adjustVddBinningRow = new AdjustVddBinningRow
                    {
                        PowerName = power.PinMode,
                        Type = type,
                        Site = site
                    };
                    adjustVddBinningRows.Add(adjustVddBinningRow);
                }
            }
            return adjustVddBinningRows;
        }
    }
}
