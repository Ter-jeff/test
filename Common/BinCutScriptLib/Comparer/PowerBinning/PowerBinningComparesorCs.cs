using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

namespace BinCutScriptLib.Comparer.PowerBinning
{
    internal class PowerBinningComparesorCs(OneTouchDown oneTouchDown, CheckManager checkManager, EnumJob enumJob) : PowerBinningComparesor(oneTouchDown, checkManager, enumJob)
    {
        public void CheckPowerBinningCs(ref SiteInfo[] siteInfoArray, StreamWriter streamWriter)
        {
            if (BinCutConfig.IsEnablePowerBinningHarvest && BinCutData.PowerBinHavrSheet != null && CheckPowerBinningTable())
            {
                GetHarvestBin(ref OneTouchDown, ref siteInfoArray);
                List<PowerBinningRowIdx> siteRowId = FindFirstPowerBinningRow(siteInfoArray);
                List<BinCutLineBase> powerBinningLines;
                while (OneTouchDown.GetPowerBinningHarvCs(out powerBinningLines, out string sheetName, "Power Binning Summary"))
                {
                    CheckManager.TotalPowerBinningCnt += powerBinningLines.Count;
                    CheckManager.MissingBv.AddRange(powerBinningLines);
                    if (!BinCutData.PowerBinningSheetList.First()
                        .Value.BinnedModeList.First()
                        .FactorDictionary.Keys.Any(x => x.Contains("AC_")))
                    {
                        CheckPowerBinningHarvest(streamWriter, siteInfoArray, ref powerBinningLines, siteRowId, sheetName);
                    }
                }

                //Check all site with all sheet should be check pass
                CheckSiteMissPwrBinSheet(streamWriter, ref siteInfoArray, siteRowId);

                //Power binning summary by site
                if (GetPowerBinningSummary(ref OneTouchDown, ref powerBinningLines))
                {
                    CheckManager.TotalPowerBinningCnt += powerBinningLines.Count;
                    CheckManager.MissingBv.AddRange(powerBinningLines);

                    for (int site = 0; site < siteInfoArray.Length; site++)
                    {
                        if (siteRowId.Exists(x => x.Site == site && x.BPass))
                        {
                            PowerBinningRowIdx powerBinningRowIdx = siteRowId.Find(x => x.Site == site)!;
                            CheckPTotalSummaryLine(streamWriter, ref siteInfoArray, ref powerBinningLines, site, powerBinningRowIdx.RowIdx);
                        }
                        else
                        {
                            int lineIdx = 0;
                            Dictionary<string, string> logPowerBinningSumRow = GetLogPotal(site, ref powerBinningLines, ref lineIdx);
                            if (logPowerBinningSumRow.Count != 0)
                            {
                                if (siteInfoArray[site].Bin.ToString(CultureInfo.InvariantCulture) != logPowerBinningSumRow["power_binning"].Trim())
                                {
                                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, powerBinningLines[lineIdx], site, logPowerBinningSumRow["power_binning"], $"{siteInfoArray[site].Bin:F4}", "power_binning");
                                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                                }
                                else if (BinCutConfig.IsDebugPrint)
                                {
                                    BinCutPrint.PrintDifferencePass(streamWriter, siteInfoArray, powerBinningLines[lineIdx], site, logPowerBinningSumRow["power_binning"], $"{siteInfoArray[site].Bin:F4}", "power_binning");
                                }

                                CheckManager.CheckPowerBinningLine(powerBinningLines[lineIdx]);
                            }
                        }
                    }
                }
            }

            CheckPowerBinningFlow(streamWriter, siteInfoArray);
        }

        protected override Dictionary<string, string> GetLogPotal(int site, ref List<BinCutLineBase> binCutLineBases, ref int lineIdx)
        {
            //[INFO]  [Site 0] power_binning = -1, spec_name = F_Power_Binning_Fail
            var powerBinningSumRow = new Dictionary<string, string>();
            for (int i = 0; i < binCutLineBases.Count; i++)
            {
                if (binCutLineBases[i].Line.StartsWithIgnoreCase("[INFO]  [Site "))
                {
                    int siteLog = binCutLineBases[i].GetSite();
                    if (siteLog == site)
                    {
                        powerBinningSumRow.Add("Site", siteLog.ToString());
                        string[] arr = binCutLineBases[i].Line.Split([',', ']'], StringSplitOptions.RemoveEmptyEntries);
                        foreach (string item in arr)
                        {
                            lineIdx = i;
                            if (item.Trim().Contains('='))
                            {
                                string[] array = item.Split(['='], StringSplitOptions.RemoveEmptyEntries);
                                powerBinningSumRow.Add(array[0].Trim(), array[1].Trim());
                            }
                        }
                    }
                }
            }
            return powerBinningSumRow;
        }

        protected override void ComparePower_Binning_P_total(StreamWriter streamWriter, SiteInfo[] siteInfoArray, int lineNo, int site, double valueLog, double specLog, double expectspec, double ptotal, string _)
        {
            // Inherited Override Method: Uses static string tags for generic passes
            EvaluatePowerMetrics(streamWriter, siteInfoArray, lineNo, site, valueLog, specLog, ptotal, expectspec, "Spec", "P_total");
        }
    }
}
