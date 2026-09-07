using System;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

using LogLib.Utility;

using TestPlanLib.BinCut.BinCutConfig;
using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib.Comparer.BVChecker.GetVoltage
{
    internal class GetVoltageNew(EnumJob enumJob, InstanceBinCut instanceBinCut) : GetVoltageBase(enumJob, instanceBinCut)
    {
        public override double GetValue(BinCutExpress binCutExpress, SiteInfo siteInfo, BvName bvName, StreamWriter streamWriter)
        {
            try
            {
                if (siteInfo.AllPowers.Count == 0)
                {
                    return 0.0;
                }

                double retVal = 0.0;

                if (binCutExpress.PowerType == EnumPowerType.Core_Power)
                {
                    PowerZone powerZone = siteInfo.GetPowerZone(binCutExpress);
                    BinningTable binningTable = BinCutData.BinningTables[siteInfo.Bin - 1];
                    string mode = BinCutAlgorithmService.GetModeByName(binCutExpress.Express);
                    double productValueSearch = 0.0;
                    if (InstanceBinCut.CurInstanceName.Contains("BINRESULT") && powerZone.SearchStatus == EnumSearchStatus.Search)
                    {
                        retVal = powerZone.SearchStatus != EnumSearchStatus.Search
                        ? powerZone.AllSteps.Last().Lvcc
                        : powerZone.GetFinalLvcc();
                        return retVal;
                    }
                    HandleCorePower(binCutExpress, siteInfo, streamWriter, ref retVal, powerZone, binningTable, mode, ref productValueSearch);
                }
                else
                {
                    BinningTable otherRailTable = BinCutData.OtherRailTables[siteInfo.Bin - 1];
                    retVal = GetRamValue(binCutExpress, otherRailTable);
                }
                return retVal;
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                throw new Exception();
            }
        }

        private void HandleCorePower(BinCutExpress binCutExpress, SiteInfo siteInfo, StreamWriter streamWriter, ref double retVal, PowerZone powerZone, BinningTable binningTable, string mode, ref double productValueSearch)
        {
            //0: MS001_0 for evaluation      ex: MS001 Evaluate Bin
            //1: MS002_1 for result bin      ex: MG004 Bin Result
            //2: MC607_2 for HVCC            ex: HVCC Level at MC607
            //3: MS001_3_E1 for equation E1  ex: MC601 E1 Voltage
            switch (binCutExpress.ExpressType)
            {
                case EnumExpressType.Product:
                    {
                        HandleProduct(siteInfo, streamWriter, out retVal, powerZone, mode, out productValueSearch);
                        break;
                    }
                case EnumExpressType.ProductGb:
                    {
                        HandleProductGb(siteInfo, streamWriter, out retVal, powerZone, binningTable, mode, out productValueSearch);
                        break;
                    }
                case EnumExpressType.E1Voltage:
                    {
                        int bin = siteInfo.Bin;
                        retVal = GetE1Voltage(binCutExpress, siteInfo, powerZone, bin);
                        break;
                    }
                case EnumExpressType.BinProduct:
                    {
                        int bin = 1;
                        if (binCutExpress.Express.Contains("BINX", StringComparison.OrdinalIgnoreCase))
                        {
                            bin = 2;
                        }
                        else if (binCutExpress.Express.Contains("BINY", StringComparison.OrdinalIgnoreCase))
                        {
                            bin = 3;
                        }

                        retVal = GetE1ProductVoltage(binCutExpress, siteInfo, powerZone, bin);
                        break;
                    }
                case EnumExpressType.BinProductGb:
                    {
                        retVal = HandleBinProductGb(binCutExpress, siteInfo, powerZone, binningTable, mode);
                        break;
                    }
                case EnumExpressType.E1Product:
                    {
                        int bin = siteInfo.Bin;
                        retVal = GetE1ProductVoltage(binCutExpress, siteInfo, powerZone, bin);
                        break;
                    }
                case EnumExpressType.E1ProductGb:
                    {
                        retVal = HandleE1ProductGb(binCutExpress, siteInfo, powerZone, binningTable, mode);
                        break;
                    }
                case EnumExpressType.Evaluate:
                    {
                        int step = powerZone.IsAdjust ? powerZone.FinalStep : powerZone.Step;
                        double dynamicoffSetValue = GetDynamicOffSetValue(powerZone.Mode, powerZone.PossibleSteps[step].EqName, binningTable, TestCat, InstanceBinCut.CurInstType);
                        double lvcc = SiteInfoHelpers.GetVoltageOfEvaluation(powerZone);
                        siteInfo.CurrentdynamicOffset = dynamicoffSetValue;
                        retVal = lvcc + dynamicoffSetValue;
                        break;
                    }
                case EnumExpressType.BinResult:
                    {
                        if (siteInfo.IsPreVddSearch)
                        {
                            int index = powerZone.PossibleSteps.FindIndex(x => x.Bin == siteInfo.Bin && x.EqName == 1);
                            retVal = index != -1 ? powerZone.PossibleSteps[index].Lvcc : powerZone.AllSteps.Last().Lvcc;
                            break;
                        }
                        retVal = powerZone.SearchStatus != EnumSearchStatus.Search
                            ? powerZone.AllSteps.Last().Lvcc
                            : powerZone.GetFinalLvcc();
                        break;
                    }
                case EnumExpressType.Hvcc:
                    {
                        retVal = powerZone.PossibleSteps[0].CpHv;
                        for (int pStepIdx = 0; pStepIdx < powerZone.GetPosCount(); pStepIdx++)
                        {
                            if (powerZone.PossibleSteps[pStepIdx].Bin == siteInfo.Bin)
                            //!!HVCC也有分Bin1和Bin2, 依據目前的Bin給HVCC值, 並注意, 這裡假設同Bin下所有的HV STEP皆相同, 否則會出錯
                            {
                                retVal = powerZone.PossibleSteps[pStepIdx].CpHv;
                            }
                        }
                        break;
                    }
                case EnumExpressType.VMax:
                    {
                        BinningTable binningTb = binningTable;
                        retVal = GetRamValue(binCutExpress, binningTb);
                        break;
                    }
            }
        }

        private double HandleE1ProductGb(BinCutExpress binCutExpress, SiteInfo siteInfo, PowerZone powerZone, BinningTable binningTable, string mode)
        {
            double retVal;
            int bin = siteInfo.Bin;
            retVal = GetE1ProductVoltage(binCutExpress, siteInfo, powerZone, bin);

            int lvVddBinGbIdx = GetLvVddBinGbIdx(binningTable);
            double dGbTmp = 0.0;
            for (int rowIdx = 0; rowIdx < binningTable.Rows.Count; rowIdx++)
            {
                string modeInVddBin = binningTable.Rows[rowIdx].RowData[binningTable.ModeIdx];
                if (modeInVddBin == mode)
                {
                    dGbTmp = double.Parse(binningTable.Rows[rowIdx].RowData[lvVddBinGbIdx]);
                    break;
                }
            }
            retVal -= dGbTmp;
            return retVal;
        }

        private double HandleBinProductGb(BinCutExpress binCutExpress, SiteInfo siteInfo, PowerZone powerZone, BinningTable binningTable, string mode)
        {
            double retVal;
            int bin = 1;
            if (binCutExpress.Express.Contains("BINX", StringComparison.OrdinalIgnoreCase))
            {
                bin = 2;
            }
            else if (binCutExpress.Express.Contains("BINY", StringComparison.OrdinalIgnoreCase))
            {
                bin = 3;
            }

            retVal = GetE1ProductVoltage(binCutExpress, siteInfo, powerZone, bin);
            int lvVddBinGbIdx = GetLvVddBinGbIdx(binningTable);
            double dGbTmp = 0.0;
            for (int rowIdx = 0; rowIdx < binningTable.Rows.Count; rowIdx++)
            {
                string modeInVddBin = binningTable.Rows[rowIdx].RowData[binningTable.ModeIdx];
                if (modeInVddBin == mode)
                {
                    dGbTmp = double.Parse(binningTable.Rows[rowIdx].RowData[lvVddBinGbIdx]);
                    break;
                }
            }
            retVal -= dGbTmp;
            return retVal;
        }

        private void HandleProductGb(SiteInfo siteInfo, StreamWriter streamWriter, out double retVal, PowerZone powerZone, BinningTable binningTable, string mode, out double productValueSearch)
        {
            if (powerZone.Pin.Contains("SRAM") && BinCutData.PinInfos.Find(x => x.PinMode == powerZone.PinMode)!.Binned.EqualsIgnoreCase("True"))
            {
                double value = SiteInfoHelpers.GetEfuseProductValue(mode, siteInfo.EFuseValues);
                productValueSearch = value != 0 ? value : powerZone.GetFinalProductValue(streamWriter);
            }
            else
            {
                productValueSearch = BinCutData.BinCutFlowTables.Find(x => x.FinalJob.Contains(BinCutData.Job.JobType.ToString()))!.GetEvaluateModes().Contains(mode) ? powerZone.GetFinalProductValue(streamWriter) : SiteInfoHelpers.GetEfuseProductValue(mode, siteInfo.EFuseValues);
            }
            int lvVddBinGbIdx = GetLvVddBinGbIdx(binningTable);
            double dGbTmp = 0.0;
            for (int rowIdx = 0; rowIdx < binningTable.Rows.Count; rowIdx++)
            {
                string modeInVddBin = binningTable.Rows[rowIdx].RowData[binningTable.ModeIdx];
                if (modeInVddBin == mode)
                {
                    dGbTmp = double.Parse(binningTable.Rows[rowIdx].RowData[lvVddBinGbIdx]);
                    break;
                }
            }
            retVal = BinCutAlgorithmService.Floor3P125(productValueSearch - dGbTmp);
        }

        private static void HandleProduct(SiteInfo siteInfo, StreamWriter streamWriter, out double retVal, PowerZone powerZone, string mode, out double productValueSearch)
        {
            if (powerZone.Pin.Contains("SRAM") && BinCutData.PinInfos.Find(x => x.PinMode == powerZone.PinMode)!.Binned.EqualsIgnoreCase("True"))
            {
                productValueSearch = SiteInfoHelpers.GetEfuseProductValue(mode, siteInfo.EFuseValues) != 0 ? SiteInfoHelpers.GetEfuseProductValue(mode, siteInfo.EFuseValues) : powerZone.GetFinalProductValue(streamWriter);
            }
            else
            {
                productValueSearch = BinCutData.BinCutFlowTables.Find(x => x.FinalJob.Contains(BinCutData.Job.JobType.ToString()))!.GetEvaluateModes().Contains(mode) ? powerZone.GetFinalProductValue(streamWriter) : SiteInfoHelpers.GetEfuseProductValue(mode, siteInfo.EFuseValues);
            }
            retVal = productValueSearch;
        }
    }
}
