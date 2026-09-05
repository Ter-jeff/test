using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib.BinCut;
using TestPlanLib.BinCut.PowerBinning;

namespace BinCutScriptLib.Comparer.PowerBinning
{
    internal class PowerBinningComparesor(OneTouchDown oneTouchDown, CheckManager checkManager, EnumJob enumJob)
    {
        protected OneTouchDown OneTouchDown = oneTouchDown;
        protected readonly CheckManager CheckManager = checkManager;
        protected readonly EnumJob Job = enumJob;

        public void CheckPowerBinning(ref SiteInfo[] siteInfoArray, StreamWriter streamWriter)
        {
            var powerBinningLines = new List<BinCutLineBase>();
            string sheetName;

            if (BinCutConfig.IsEnablePowerBinningHarvest && BinCutData.PowerBinHavrSheet != null && CheckPowerBinningTable())
            {
                GetHarvestBin(ref OneTouchDown, ref siteInfoArray);
                List<PowerBinningRowIdx> siteRowId = FindFirstPowerBinningRow(siteInfoArray);
                while (OneTouchDown.GetPowerBinningHarvCs(out powerBinningLines, out sheetName, "Power Binning Summary"))
                {
                    //20/03/06 remove repeate power binning data, ex: ids. Except pTotal. Recover other power binning lines.
                    //UpdatePowerBinnningLines(sheetName, ref powerBinningLines);
                    CheckManager.TotalPowerBinningCnt += powerBinningLines.Count;
                    CheckManager.MissingBv.AddRange(powerBinningLines);
                    if (!BinCutData.PowerBinningSheetList.First()
                        .Value.BinnedModeList.First()
                        .FactorDictionary.Keys.Any(x => x.Contains("AC_")))
                    {
                        CheckPowerBinningHarvest(streamWriter, siteInfoArray, ref powerBinningLines, siteRowId, sheetName);
                    }
                    else
                    {
                        ComparePowerBinningTurks(streamWriter, siteInfoArray, ref powerBinningLines, siteRowId, sheetName, BinCutData.PowerBinningSheetList.TryGetValue(sheetName, out PowerBinningSheet? value) ? value! : BinCutData.PowerBinningSheetList.First().Value);
                    }
                }

                //Check all site with all sheet should be check pass
                CheckSiteMissPwrBinSheet(streamWriter, ref siteInfoArray, siteRowId);

                //Power binning summary by site
                var powerBinningSummary = new List<BinCutLineBase>();
                if (GetPowerBinningSummary(ref OneTouchDown, ref powerBinningSummary))
                {
                    CheckManager.TotalPowerBinningCnt += powerBinningSummary.Count;
                    CheckManager.MissingBv.AddRange(powerBinningSummary);

                    for (int site = 0; site < siteInfoArray.Length; site++)
                    {
                        if (siteRowId.Exists(x => x.Site == site && x.BPass))
                        {
                            PowerBinningRowIdx powerBinningRowIdx = siteRowId.Find(x => x.Site == site)!;
                            CheckPTotalSummaryLine(streamWriter, ref siteInfoArray, ref powerBinningSummary, site, powerBinningRowIdx.RowIdx);
                        }
                        else
                        {
                            int lineIdx = 0;
                            Dictionary<string, string> logPowerBinningSumRow = GetLogPotal(site, ref powerBinningLines, ref lineIdx);
                            if (logPowerBinningSumRow.Count != 0)
                            {
                                //Check Currentpassbinnum
                                if (siteInfoArray[site].Bin.ToString(CultureInfo.InvariantCulture) != logPowerBinningSumRow["currentpassbinnum"].Trim())
                                {
                                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, powerBinningLines[lineIdx].LineNo.ToString(CultureInfo.InvariantCulture), site, logPowerBinningSumRow["currentpassbinnum"], $"{siteInfoArray[site].Bin:F4}", "currentpassbinnum");
                                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                                }

                                CheckManager.CheckPowerBinningLine(powerBinningLines[lineIdx]);
                            }
                        }
                    }
                }
            }
            else
            {
                #region  Old power binning method
                while (OneTouchDown.GetPowerBinningHarvCs(out powerBinningLines, out sheetName, "<PrintOutVddBinning>"))
                {
                    CheckManager.TotalPowerBinningCnt += powerBinningLines.Count;
                    CheckManager.MissingBv.AddRange(powerBinningLines);
                    if (powerBinningLines.Count > 0 && BinCutData.PowerBinningSheetList.Count > 0)
                    {

                        {
                            PowerBinningSheet powerBinningSheet =
                                BinCutData.PowerBinningSheetList.TryGetValue(sheetName, out PowerBinningSheet? value)
                                    ? value!
                                    : BinCutData.PowerBinningSheetList.First().Value;
                            PowerBinningMain powerBinningMain = new PowerBinningMain();
                            var sheet = (PowerBinningSheetCebu)powerBinningSheet;
                            PowerBinningMain.CreatePowerBinning(sheet, siteInfoArray);
                            PowerBinningMain.SetPowerBinningIds(streamWriter, siteInfoArray, ref powerBinningLines, sheetName);
                            powerBinningMain.CalculatePowerBinning(siteInfoArray, Job, streamWriter, powerBinningSheet.NewMethod);

                            ComparePowerBinningCebu(streamWriter, siteInfoArray, ref powerBinningLines, sheetName);
                        }
                    }
                }

                #endregion
            }

            CheckPowerBinningFlow(streamWriter, siteInfoArray);
        }

        protected static List<PowerBinningRowIdx> FindFirstPowerBinningRow(SiteInfo[] siteInfoArray)
        {
            var siteRowId = new List<PowerBinningRowIdx>();
            if (BinCutData.PowerBinHavrSheet != null)
            {
                for (int site = 0; site < siteInfoArray.Length; site++)
                {
                    if (siteInfoArray[site].SiteIsBinOut || !siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                    {
                        continue;
                    }

                    string bin = siteInfoArray[site].Bin <= BinCutData.BinningTables.Count ? (siteInfoArray[site].Bin - 1).ToString(CultureInfo.InvariantCulture) : "x";
                    var harvbins = new List<string>();
                    foreach (int harvBin in siteInfoArray[site].HarvBin)
                    {
                        harvbins.Add(harvBin >= 0 && harvBin < 2 ? harvBin.ToString(CultureInfo.InvariantCulture) : "x");
                    }
                    //var harvbin = (allDice[site].HarvBin >= 0 && allDice[site].HarvBin < 2) ? allDice[site].HarvBin.ToString(CultureInfo.InvariantCulture) : "x";
                    for (int rowid = 0; rowid < BinCutData.PowerBinHavrSheet.Rows.Count; rowid++)
                    {
                        string tblbin = BinCutData.PowerBinHavrSheet.Rows.ElementAt(rowid).Inputinfo[0].ElementAtOrDefault(0).Value;
                        var tblharvbins = new List<string>();
                        bool binCompare = true;
                        for (int idx = 0; idx < harvbins.Count; idx++)
                        {
                            string tblharvbin = BinCutData.PowerBinHavrSheet.Rows.ElementAt(rowid).Inputinfo[idx + 1].ElementAtOrDefault(0).Value;
                            //if table's bin is x or empty, it shouldn't to compare
                            if (string.IsNullOrEmpty(tblharvbin))
                            {
                                tblharvbin = "x";
                            }

                            tblharvbin = tblharvbin.EqualsIgnoreCase("x") ? harvbins[idx] : tblharvbin;
                            tblharvbins.Add(tblharvbin);
                        }

                        for (int num = 0; num < harvbins.Count; num++)
                        {
                            if (tblharvbins[num] != harvbins[num])
                            {
                                binCompare = false;
                                break;
                            }
                        }
                        if (tblbin == bin && binCompare)
                        {
                            //200410 Ture: this site w/o Binning sheet value (ex. binX binY)
                            bool bSkip = false;
                            List<Dictionary<string, bool>> sheets = GetSheetInfoList(rowid, ref bSkip);
                            siteRowId.Add(new PowerBinningRowIdx { Site = site, RowIdx = rowid, BPass = bSkip, CheckSheetList = sheets });
                            break;
                        }
                    }
                }
            }
            return siteRowId;
        }

        protected void CheckPowerBinningHarvest(StreamWriter streamWriter, SiteInfo[] siteInfoArray, ref List<BinCutLineBase> binCutLineBases, List<PowerBinningRowIdx> powerBinningRowIdxs, string sheetname)
        {
            if (binCutLineBases.Count > 0 && BinCutData.PowerBinningSheetList.Count > 0)
            {
                PowerBinningSheet powerBinningSheet = BinCutData.PowerBinningSheetList.TryGetValue(sheetname, out PowerBinningSheet? value)
                    ? value!
                    : BinCutData.PowerBinningSheetList.First().Value;
                var sheet = (PowerBinningSheetCebu)powerBinningSheet;
                PowerBinningMain powerBinningMain = new PowerBinningMain();
                PowerBinningMain.CreatePowerBinning(sheet, siteInfoArray);
                PowerBinningMain.SetPowerBinningIds(streamWriter, siteInfoArray, ref binCutLineBases, powerBinningSheet.SheetName);
                powerBinningMain.CalculatePowerBinning(siteInfoArray, Job, streamWriter, powerBinningSheet.NewMethod);
                CompareWithHarvPowerBinning(streamWriter, siteInfoArray, ref binCutLineBases, powerBinningSheet.SheetName, powerBinningRowIdxs);
            }
        }

        protected void CheckPTotalSummaryLine(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, ref List<BinCutLineBase> binCutLineBases, int site, int rowid)
        {
            //site:0, currentpassbinnum=1, power_binning = 3, fuse_name2 = 95.2465, spec_name = bin1_high_power
            int lineIdx = 0;
            int binnum = 0;
            string comparesh = "";
            Dictionary<string, string> logPowerBinningSumRow = GetLogPotal(site, ref binCutLineBases, ref lineIdx);
            PowerBinningHarvRow powerBinHarvRow = BinCutData.PowerBinHavrSheet!.Rows[rowid];
            if (logPowerBinningSumRow.Count == 0)
            {
                return;
            }

            //Check PTotal
            //var isSumPtotal = isSumOfPtotal(rowid, ref binnum, ref comparesh);
            if (IsSumOfPtotal(rowid, ref binnum, ref comparesh))
            {
                if (!string.IsNullOrEmpty(comparesh))   //ignore fuse_name2 empty case, BC owner say it can skip to check
                {
                    Tuple<string, string>? p = siteInfoArray[site].PowerBinningPTotalSummary.Find(x => x.Item1.StartsWithIgnoreCase(comparesh));
                    if (p != null && !logPowerBinningSumRow["fuse_name2"].EqualsIgnoreCase(p.Item2))
                    {
                        BinCutPrint.PrintDifference(streamWriter, siteInfoArray, binCutLineBases[lineIdx].LineNo.ToString(CultureInfo.InvariantCulture), site, logPowerBinningSumRow["fuse_name2"], p.Item2, "P_total");
                        siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                    }
                }
            }
            else
            {
                if (!binnum.ToString(CultureInfo.InvariantCulture).EqualsIgnoreCase(logPowerBinningSumRow["fuse_name2"]))
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, binCutLineBases[lineIdx].LineNo.ToString(CultureInfo.InvariantCulture), site, logPowerBinningSumRow["fuse_name2"], $"{binnum:F4}", "P_total");
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
            }

            //Check Currentpassbinnum
            if (siteInfoArray[site].Bin.ToString(CultureInfo.InvariantCulture) != logPowerBinningSumRow["currentpassbinnum"].Trim())
            {
                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, binCutLineBases[lineIdx].LineNo.ToString(CultureInfo.InvariantCulture), site, logPowerBinningSumRow["currentpassbinnum"], $"{siteInfoArray[site].Bin:F4}", "currentpassbinnum");
                siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
            }

            //Check power_binning
            if (powerBinHarvRow.PowerBinning != logPowerBinningSumRow["power_binning"].Trim())
            {
                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, binCutLineBases[lineIdx].LineNo.ToString(CultureInfo.InvariantCulture), site, logPowerBinningSumRow["power_binning"], powerBinHarvRow.PowerBinning, "power_binning");
                siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
            }

            //Check spec_name
            if (powerBinHarvRow.Comment != logPowerBinningSumRow["spec_name"])
            {
                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, binCutLineBases[lineIdx].LineNo.ToString(CultureInfo.InvariantCulture), site, logPowerBinningSumRow["spec_name"], powerBinHarvRow.Comment, "spec_name");
                siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
            }
            CheckManager.CheckPowerBinningLine(binCutLineBases[lineIdx]);
        }

        private static bool IsSumOfPtotal(int rowId, ref int binNum, ref string compsheetname)
        {
            string fuseName = BinCutData.PowerBinHavrSheet!.Rows[rowId].FuseName2;
            if (int.TryParse(fuseName, out int val))
            {
                binNum = val;
                //Bin number
                return false;
            }
            compsheetname = fuseName;
            //sheet name
            return true;
        }

        protected virtual Dictionary<string, string> GetLogPotal(int site, ref List<BinCutLineBase> binCutLineBases, ref int lineIdx)
        {
            //site:0, currentpassbinnum=1, power_binning = 3, fuse_name2 = 95.2465, spec_name = bin1_high_power
            //site:1, currentpassbinnum=1, power_binning = 3, fuse_name2 = 94.6793, spec_name = bin1_high_power
            //site:2, currentpassbinnum=2, power_binning = 4, fuse_name2 = 7, spec_name = binX_power
            var powerBinningSumRow = new Dictionary<string, string>();
            for (int lnIdx = 0; lnIdx < binCutLineBases.Count; lnIdx++)
            {
                if (binCutLineBases[lnIdx].Line.StartsWithIgnoreCase("site"))
                {
                    bool findSite = false;
                    string[] spt = binCutLineBases[lnIdx].Line.Split([','], StringSplitOptions.RemoveEmptyEntries);
                    foreach (string st in spt)
                    {
                        if (st.Trim().StartsWithIgnoreCase("site"))
                        {
                            string[] s = st.Split([':'], StringSplitOptions.RemoveEmptyEntries);
                            if (s[1].Trim() == site.ToString(CultureInfo.InvariantCulture))
                            {
                                findSite = true;
                                powerBinningSumRow.Add(s[0].Trim(), s[1].Trim());
                                continue;
                            }
                        }

                        if (findSite)
                        {
                            lineIdx = lnIdx;
                            if (st.Trim().Contains('='))
                            {
                                string[] p = st.Split(['='], StringSplitOptions.RemoveEmptyEntries);
                                powerBinningSumRow.Add(p[0].Trim(), p[1].Trim());
                            }
                        }
                    }

                }
            }
            return powerBinningSumRow;
        }

        protected static void CheckSiteMissPwrBinSheet(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<PowerBinningRowIdx> powerBinningRowIdxs)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].SiteIsBinOut || !siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                {
                    continue;
                }

                PowerBinningRowIdx? idx = powerBinningRowIdxs.Find(x => x.Site == site && x.BPass);
                if (idx != null)
                {
                    List<Dictionary<string, bool>> sheets = idx.CheckSheetList.FindAll(x => x.ContainsValue(false));
                    for (int i = 0; i < sheets.Count; i++)
                    {
                        string errMsg =
                            $"Site: [{site}] : rowids {idx.RowIdx} ,and miss sheetname {sheets.ElementAt(i).FirstOrDefault().Key}";
                        BinCutPrint.PrintErrorMessage(streamWriter, errMsg);
                        siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                    }
                }
            }
        }

        protected static void CheckPowerBinningFlow(StreamWriter streamWriter, SiteInfo[] siteInfoArray)
        {
            if (BinCutData.PwrbinSeqSheet != null && BinCutData.PowerBinHavrSheet == null)
            {
                for (int site = 0; site < siteInfoArray.Length; site++)
                {
                    if (siteInfoArray[site].SiteIsBinOut || !siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                    {
                        continue;
                    }

                    if (siteInfoArray[site].PowerBinningSeq > 0)
                    {
                        string sheet = BinCutData.PwrbinSeqSheet.GetPwrbinSeqSheet(siteInfoArray[site].PowerBinningSeq, siteInfoArray[site].Bin);
                        string finalsheet = BinCutData.PwrbinSeqSheet.GetFinalPwrbinSeqSheet(siteInfoArray[site].Bin);
                        string errMsg = $"Site:{site} : Final Sheet was {finalsheet} ,and stopped as {sheet}";
                        BinCutController.Controller.RichTextBoxAppend(errMsg, Color.Red);
                        streamWriter.WriteLine("<PowerBinningFlow>");
                        streamWriter.WriteLine(errMsg);
                    }
                }
            }
        }

        public static bool CheckPowerBinningTable()
        {
            int binmax = 0;
            foreach (PowerBinningHarvRow r in BinCutData.PowerBinHavrSheet!.Rows)
            {
                if (int.TryParse(r.Inputinfo[0].ElementAt(0).Value, out int temp))
                {
                    binmax = temp > binmax ? temp : binmax;
                }
                else
                {
                    BinCutController.Controller.RichTextBoxAppend("Failed Product bin in PowerBinning table define : " + r.Inputinfo[0].ElementAt(0).Value, Color.Red);
                    return false;
                }
                for (int idx = 1; idx < r.Inputinfo.Count; idx++)
                {
                    if (r.Inputinfo[idx].ElementAt(0).Value == "0" || r.Inputinfo[idx].ElementAt(0).Value == "1" ||
                    r.Inputinfo[idx].ElementAt(0).Value.EqualsIgnoreCase("x") || string.IsNullOrEmpty(r.Inputinfo[idx].ElementAt(0).Value))
                    {
                    }
                    else
                    {
                        BinCutController.Controller.RichTextBoxAppend("Failed Harvest bin in PowerBinning table define : " + r.Inputinfo[idx].ElementAt(0).Value, Color.Red);
                        return false;
                    }

                }
            }
            return true;
        }

        public static bool GetPowerBinningSummary(ref OneTouchDown oneTouchDown, ref List<BinCutLineBase> binCutLineBases)
        {
            //STEP1. Search until found power binning summary
            bool isFoundPowerBinSummary = false;
            if (oneTouchDown.Lines.Count == 0)
            {
                return false;
            }

            int oneTouchIndex;
            for (oneTouchIndex = 0; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (oneTouchDown.Lines[oneTouchIndex].Line == "<PrintOutVddBinning>" || oneTouchDown.Lines[oneTouchIndex].Line == "<PrintOutVddBinning_CS>")
                {
                    return false;
                }

                if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("Power Binning Summary"))
                {
                    isFoundPowerBinSummary = true;
                    break;
                }
            }

            if (!isFoundPowerBinSummary)
            {
                return false;
            }

            oneTouchIndex++;
            oneTouchIndex++;

            for (; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (oneTouchDown.Lines[oneTouchIndex].Line.Length == 0 || oneTouchDown.Lines[oneTouchIndex].Line == "<PrintOutVddBinning>")
                {
                    break;
                }

                //STEP3a. read each power binning summary line

                //Need skip this line from customer's request 
                //Site(0)  CFGFuse SetWriteDecimal                                     power_binning = 1                    [0x1]
                if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("SetWriteDecimal"))
                {
                    continue;
                }

                //Case 1: site:0, currentpassbinnum=1, power_binning = 3, fuse_name2 = 95.2465, spec_name = bin1_high_power
                //Case 2: site:0, currentpassbinnum=3
                if (oneTouchDown.Lines[oneTouchIndex].Line.Length == 0 || oneTouchDown.Lines[oneTouchIndex].Line.Split([','], StringSplitOptions.RemoveEmptyEntries).Length < 2)
                {
                    break;
                }

                binCutLineBases.Add(oneTouchDown.Lines[oneTouchIndex]);
            }

            //delete original data from oneTchLines to speed up parse
            oneTouchDown.Lines.RemoveRange(0, oneTouchIndex);

            return true;
        }

        public void ComparePowerBinningCebu(StreamWriter streamWriter, SiteInfo[] siteInfoArray, ref List<BinCutLineBase> binCutLineBases, string sheetName)
        {
            bool flag = true;
            for (int lnIdx = 0; lnIdx < binCutLineBases.Count; lnIdx++)
            {

                string line = binCutLineBases[lnIdx].Line;
                int lineNo = binCutLineBases[lnIdx].LineNo;

                string[] spt = line.Split([" K ", " ", "Define", "(F)"], StringSplitOptions.RemoveEmptyEntries);
                if (!double.TryParse(spt[0], out double _))
                {
                    continue;
                }

                int site = int.Parse(spt[1]);
                PowerBinningSheet? powerBinning = siteInfoArray[site].PowerBinning;
                if (powerBinning == null)
                {
                    continue;
                }

                string itemName = spt[2].Replace(sheetName + "_", "");
                string pwrName = "";
                string pwrTypeLog = "";
                GetPowerBinningNameType(sheetName, itemName, spt, ref pwrName, ref pwrTypeLog);

                _ = double.TryParse(spt[5], out double log);
                double powerTotal = powerBinning.BinnedModeList.Sum(item => item.PowerValue) + powerBinning.OtherModeList.Sum(item => item.PowerValue);

                double ptotal = Math.Round(powerTotal + powerBinning.Offset, 4, MidpointRounding.AwayFromZero);
                siteInfoArray[site].PowerBinningPTotalSummary.Add(new Tuple<string, string>(sheetName + "_" + itemName, ptotal.ToString()));

                #region Check power binning sheet
                if (flag && BinCutData.PwrbinSeqSheet != null)
                {
                    string sheet = BinCutData.PwrbinSeqSheet.GetPwrbinSeqSheet(siteInfoArray[site].PowerBinningSeq, siteInfoArray[site].Bin);
                    if (!sheet.EqualsIgnoreCase(sheetName))
                    {
                        BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, sheetName, sheet, "PwrBin SheetName");
                        flag = false;
                    }
                }
                #endregion

                #region Compare Ids,Vbin,P per mode
                PowerBinningSheetRow? powerBinningRow = PowerBinningMain.GetPowerBinningRow(powerBinning, pwrName);
                if (powerBinningRow == null && pwrName != "Total")
                {
                    //PrintDiff.PrintMissingLines(sw, new List<LineData>() { powerBinningLines[lnIdx] });
                    continue;
                }

                switch (pwrTypeLog)
                {
                    //case "SRAM_Ids":
                    //    if (Math.Abs(log - sramIds) >= 0.0002)
                    //    {
                    //        PrintDifference(sw, allDice, lineNo.ToString(), site, string.Format("{0:F4}", log), string.Format("{0:F4}", sramIds), "SRAM_Ids", false);
                    //        allDice[site].IsCurrPowerBinningPass = false;
                    //    }
                    //    TotalPowerBinningCheckCnt++;
                    //    break;
                    //case "SRAM_P_Other":
                    //    sramP = GetPowerValueByMode("VDD_SRAM", sramIds, 0);
                    //    if (Math.Abs(log - sramP) >= 0.0002)
                    //    {
                    //        PrintDifference(sw, allDice, lineNo.ToString(), site, string.Format("{0:F4}", log), string.Format("{0:F4}", sramP), "SRAM_P_Other", false);
                    //        allDice[site].IsCurrPowerBinningPass = false;
                    //    }
                    //    TotalPowerBinningCheckCnt++;
                    //    break;
                    case "Power_Binning_P_total":
                        {
                            _ = double.TryParse(spt[6], out double spec);
                            double expectspec = Math.Round(powerBinning.Spec, 4, MidpointRounding.AwayFromZero);
                            if (Math.Abs(spec - expectspec) > 0.001)
                            {
                                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{spec:F4}", $"{expectspec:F4}", "Spec");
                                siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                            }

                            //#region Compare _SRAM_Ids
                            //double ptotalSram = Math.Round(powerTotalSram + powerBinningResult[site].Offset + sramP, 4, MidpointRounding.AwayFromZero);
                            //ptotal = ptotalSram; //Merge SRAM
                            //#endregion

                            if (Math.Abs(log - ptotal) > 0.001)
                            {
                                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{ptotal:F4}", "P_total");
                                siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                            }
                            PowerBinningMain.AdjustPowerBinningBin(siteInfoArray, ptotal, site, sheetName);
                            CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
                        }
                        break;
                    case "Ids":
                        double expectIds = powerBinningRow!.IdsValue;
                        expectIds = Math.Round(expectIds, 4, MidpointRounding.AwayFromZero);
                        if (Math.Abs(log - expectIds) > 0.001)
                        {
                            BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectIds:F4}", itemName);
                            siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                        }
                        CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
                        break;
                    case "Vbin":
                        double expectVbin = powerBinningRow!.ProductValue;
                        if (Math.Abs(log - expectVbin) > 0.001)
                        {
                            BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectVbin:F4}", itemName);
                            siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                        }
                        CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
                        break;
                    case "P":
                        double expectP = Math.Round(powerBinningRow!.PowerValue, 4, MidpointRounding.AwayFromZero);
                        if (Math.Abs(log - expectP) > 0.001)
                        {
                            BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectP:F4}", itemName);
                            siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                        }
                        CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
                        break;
                    case "C":
                        if (powerBinningRow!.FactorDictionary.TryGetValue("C", out double expectC))
                        {
                            expectC = Math.Round(expectC, 4, MidpointRounding.AwayFromZero);
                            if (Math.Abs(log - expectC) >= 0.0002)
                            {
                                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectC:F4}", itemName);
                                siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                            }
                            CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
                        }
                        break;
                    default:
                        //PrintDiff.PrintMissingLines(sw, new List<LineData>() { powerBinningLines[lnIdx] });
                        break;
                }
                #endregion
            }
        }

        public static List<Dictionary<string, bool>> GetSheetInfoList(int rowid, ref bool bSkip)
        {
            var sheetList = new List<Dictionary<string, bool>>();
            int emptyCount = 0;
            List<Dictionary<string, string>> items = BinCutData.PowerBinHavrSheet!.Rows.ElementAt(rowid).SheetInfo.FindAll(x => x.Values != null);

            foreach (Dictionary<string, string> item in items)
            {
                var dic = new Dictionary<string, bool>();
                if (string.IsNullOrEmpty(item.FirstOrDefault().Value))
                {
                    emptyCount++;
                    dic.Add(item.FirstOrDefault().Key, true);
                }
                else
                {
                    dic.Add(item.FirstOrDefault().Key, false);
                }

                sheetList.Add(dic);
            }

            // All values is empty in this row then clear this sheetList
            if (emptyCount == items.Count)
            {
                bSkip = true;
            }

            return sheetList;
        }

        protected bool CompareAll(StreamWriter streamWriter, SiteInfo[] siteInfoArray, List<BinCutLineBase> binCutLineBases, string sheetName, List<PowerBinningRowIdx> powerBinningRowIdxs, PowerBinningHarvRow powerBinningHarvRow, PowerBinningRowIdx powerBinningRowIdx, string sheetInfoSpec, int rowid, int lnIdx, bool isPtotalFail, int lineNo, int site, double valueLog, double specLog, bool bcheckcurrentsite, string itemName, string pwrTypeLog, double ptotal, PowerBinningSheetRow powerBinningSheetRow)
        {
            switch (pwrTypeLog)
            {
                case "Power_Binning_P_total":
                    {
                        _ = double.TryParse(sheetInfoSpec, out double binSpec);
                        double expectspec = Math.Round(binSpec, 4, MidpointRounding.AwayFromZero);
                        ComparePower_Binning_P_total(streamWriter, siteInfoArray, lineNo, site, valueLog, specLog, expectspec, ptotal, itemName);
                        PowerBinningMain.AdjustPowerBinningHarvBin(siteInfoArray, ptotal, site, sheetName, powerBinningHarvRow);
                        CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
                        if (ptotal > expectspec)
                        {
                            isPtotalFail = true;
                        }

                        if (bcheckcurrentsite && rowid != -1)
                        {
                            UpdateSiteRowIdBySheetResult(siteInfoArray, powerBinningRowIdxs, powerBinningRowIdx, site, sheetName, rowid, isPtotalFail);
                        }
                    }
                    break;
                case "Ids":
                    double expectIds = powerBinningSheetRow.IdsValue;
                    expectIds = Math.Round(expectIds, 4, MidpointRounding.AwayFromZero);
                    if (Math.Abs(valueLog - expectIds) > 0.001)
                    {
                        BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{valueLog:F4}", $"{expectIds:F4}", itemName);
                        siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                    }
                    else if (BinCutConfig.IsDebugPrint)
                    {
                        BinCutPrint.PrintDifferencePass(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{valueLog:F4}", $"{expectIds:F4}", itemName);
                    }
                    CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
                    break;
                case "Vbin":
                    double expectVbin = powerBinningSheetRow.ProductValue;
                    if (Math.Abs(valueLog - expectVbin) > 0.001)
                    {
                        BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{valueLog:F4}", $"{expectVbin:F4}", itemName);
                        siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                    }
                    else if (BinCutConfig.IsDebugPrint)
                    {
                        BinCutPrint.PrintDifferencePass(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{valueLog:F4}", $"{expectVbin:F4}", itemName);
                    }
                    CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
                    break;
                case "P":
                    double expectP = Math.Round(powerBinningSheetRow.PowerValue, 4, MidpointRounding.AwayFromZero);
                    if (Math.Abs(valueLog - expectP) > 0.001)
                    {
                        BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{valueLog:F4}", $"{expectP:F4}", itemName);
                        siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                    }
                    else if (BinCutConfig.IsDebugPrint)
                    {
                        BinCutPrint.PrintDifferencePass(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{valueLog:F4}", $"{expectP:F4}", itemName);
                    }
                    CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
                    break;
                case "C":
                    if (powerBinningSheetRow.FactorDictionary.TryGetValue("C", out double expectC))
                    {
                        expectC = Math.Round(expectC, 4, MidpointRounding.AwayFromZero);
                        if (Math.Abs(valueLog - expectC) >= 0.0002)
                        {
                            BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{valueLog:F4}", $"{expectC:F4}", itemName);
                            siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                        }
                        else if (BinCutConfig.IsDebugPrint)
                        {
                            BinCutPrint.PrintDifferencePass(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{valueLog:F4}", $"{expectC:F4}", itemName);
                        }
                        CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
                    }
                    break;
                default:
                    break;
            }

            return isPtotalFail;
        }

        protected virtual void ComparePower_Binning_P_total(StreamWriter streamWriter, SiteInfo[] siteInfoArray, int lineNo, int site, double valueLog, double specLog, double expectspec, double ptotal, string itemName)
        {
            // Base Virtual Method: Uses dynamic itemName parameter for tracking passes
            EvaluatePowerMetrics(streamWriter, siteInfoArray, lineNo, site, valueLog, specLog, ptotal, expectspec, itemName, "P_total");
        }

        protected static void EvaluatePowerMetrics(StreamWriter streamWriter, SiteInfo[] siteInfoArray, int lineNo, int site, double valueLog, double specLog, double ptotal, double expectspec, string specPassTag, string totalPassTag)
        {
            string lineStr = lineNo.ToString();
            string specLogStr = $"{specLog:F4}";
            string expectSpecStr = $"{expectspec:F4}";
            string valueLogStr = $"{valueLog:F4}";
            string ptotalStr = $"{ptotal:F4}";

            // 1. Process Spec Mismatch Validations
            if (Math.Abs(specLog - expectspec) > 0.001)
            {
                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineStr, site, specLogStr, expectSpecStr, "Spec");
                UpdateSitePassStatus(siteInfoArray, site);
            }
            else if (BinCutConfig.IsDebugPrint)
            {
                BinCutPrint.PrintDifferencePass(streamWriter, siteInfoArray, lineStr, site, specLogStr, expectSpecStr, specPassTag);
            }

            // 2. Process P_total Mismatch Validations
            if (Math.Abs(valueLog - ptotal) > 0.001)
            {
                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineStr, site, valueLogStr, ptotalStr, "P_total");
                UpdateSitePassStatus(siteInfoArray, site);
            }
            else if (BinCutConfig.IsDebugPrint)
            {
                BinCutPrint.PrintDifferencePass(streamWriter, siteInfoArray, lineStr, site, valueLogStr, ptotalStr, totalPassTag);
            }
        }

        private static void UpdateSitePassStatus(SiteInfo[] siteInfoArray, int site)
        {
            if (siteInfoArray != null && site >= 0 && site < siteInfoArray.Length && siteInfoArray[site]?.CheckResult != null)
            {
                siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
            }
        }

        protected static void UpdateSiteRowIdBySheetResult(SiteInfo[] siteInfoArray, List<PowerBinningRowIdx> powerBinningRowIdxs, PowerBinningRowIdx powerBinningRowIdx, int site, string sheetName, int rowid, bool isBitPtotalFail)
        {
            bool bSkip = false;
            if (isBitPtotalFail)  //test fail. next step
            {
                powerBinningRowIdxs.Remove(powerBinningRowIdx);
                powerBinningRowIdx.Site = site;
                powerBinningRowIdx.RowIdx = CheckPowerBinningTableNextRow(siteInfoArray, rowid, site);
                powerBinningRowIdx.CheckSheetList = GetSheetInfoList(powerBinningRowIdx.RowIdx, ref bSkip);
                powerBinningRowIdx.BPass = false;
                //it has next row, else test fail then bin out
                if (powerBinningRowIdx.RowIdx > rowid)
                {
                    powerBinningRowIdxs.Add(powerBinningRowIdx);
                }
            }
            else    //this site and this sheet pass but isn't the last sheet, update status
            {
                Dictionary<string, bool> sh = powerBinningRowIdx.CheckSheetList.Find(x => x.ContainsKey(sheetName))!;
                sh[sheetName] = true;
            }
        }

        public void ComparePowerBinningTurks(StreamWriter streamWriter, SiteInfo[] siteInfoArray, ref List<BinCutLineBase> binCutLineBases, List<PowerBinningRowIdx> powerBinningRowIdxs, string sheetName, PowerBinningSheet powerBinningSheet)
        {
            double powerBinningPTotal = 0;
            for (int lnIdx = 0; lnIdx < binCutLineBases.Count; lnIdx++)
            {
                string line = binCutLineBases[lnIdx].Line;
                int lineNo = binCutLineBases[lnIdx].LineNo;
                var row = new PowerBinningHarvRow();
                var idx = new PowerBinningRowIdx();
                int rowid = -1;
                bool isPtotalFail = false;

                string[] spt = line.Split([" K ", " ", "Define", "(F)"], StringSplitOptions.RemoveEmptyEntries);
                if (!double.TryParse(spt[0], out double _))
                {
                    continue;
                }

                int site = int.Parse(spt[1]);
                string itemName = spt[2];
                _ = double.TryParse(spt[5], out double log);
                //10913    0     PwrBin1_Sheet_1_VDD_WARM_MW001_Vbin_AC                                                      -1       N/A            0.5375             N/A            0.0000         0
                //10914    0     PwrBin1_Sheet_1_VDD_WARM_MW001_P_Binned_AC                                                  -1       N/A            0.2421             N/A            0.0000         0
                //10915    0     PwrBin1_Sheet_1_VDD_WARM_MW001_Ids                                                          -1       N/A            0.0012             N/A            0.0000         0
                //10916    0     PwrBin1_Sheet_1_VDD_WARM_MW001_Vbin_DC                                                      -1       N/A            0.5375             N/A            0.0000         0
                //10917    0     PwrBin1_Sheet_1_VDD_WARM_MW001_P_Binned_DC                                                  -1       N/A            0.0000             N/A            0.0000         0
                //10918    0     PwrBin1_Sheet_1_mw001_pg_Vbin_PG_AC                                                         -1       N/A            0.5375             N/A            0.0000         0
                //10919    0     PwrBin1_Sheet_1_mw001_pg_P_Binned_PG_AC                                                     -1       N/A            0.2421             N/A            0.0000         0
                //10920    0     PwrBin1_Sheet_1_mw001_pg_Ids                                                                -1       N/A            0.0012             N/A            0.0000         0
                //10921    0     PwrBin1_Sheet_1_mw001_pg_Vbin_PG_DC                                                         -1       N/A            0.5375             N/A            0.0000         0
                //10922    0     PwrBin1_Sheet_1_mw001_pg_P_Binned_PG_DC                                                     -1       N/A            0.0000             N/A            0.0000         0
                //10923    0     PwrBin1_Sheet_1_VDD_WARM_MW002_Vbin_AC                                                      -1       N/A            0.5813             N/A            0.0000         0
                if (itemName.EndsWithIgnoreCase("Power_Binning_P_total") && line.Contains("(F)"))
                {
                    isPtotalFail = true;
                }

                if (powerBinningRowIdxs.Exists(x => x.Site == site && !x.BPass))
                {
                    idx = powerBinningRowIdxs.Find(x => x.Site == site && !x.BPass)!;
                    rowid = idx.RowIdx;
                    row = BinCutData.PowerBinHavrSheet!.Rows.ElementAt(rowid);
                    try
                    {
                        var sh = row.SheetInfo.Find(x => x.ContainsKey(sheetName) && !x.ContainsValue(""))!.ToList();
                    }
                    catch (Exception)
                    {
                    }
                }

                string mode = "";
                foreach (string item in powerBinningSheet.BinnedModeList.Select(x => x.BinnedMode))
                {
                    if (line.ContainsIgnoreCase(item))
                    {
                        mode = item;
                        break;
                    }
                }

                if (mode.Length == 0)
                {
                    foreach (string item in powerBinningSheet.BinnedModeList.Select(x => x.BinVoltage))
                    {
                        if (line.ContainsIgnoreCase(item + "_SRAM"))
                        {
                            mode = item + "_SRAM";
                            break;
                        }
                    }
                }

                PowerBinningSheetRow? rowCg = powerBinningSheet.BinnedModeList.Find(x => x.BinnedMode.EqualsIgnoreCase(mode));
                PowerBinningSheetRow? rowPg = powerBinningSheet.BinnedModeList.Find(x => x.BinnedMode.EqualsIgnoreCase(mode + "_PG"));
                if (rowCg == null && rowPg == null)
                {
                    rowCg =
                        powerBinningSheet.BinnedModeList.Find(
                            x => x.BinVoltage.EqualsIgnoreCase(mode));
                    rowPg =
                        powerBinningSheet.BinnedModeList.Find(
                            x => x.Ids.EqualsIgnoreCase(mode + "_PG"));
                }
                if (rowCg != null && rowPg != null)
                {
                    powerBinningPTotal = ComparePowerBinning(streamWriter, siteInfoArray, binCutLineBases, powerBinningPTotal, lnIdx, lineNo, site, itemName, log, rowCg, rowPg);
                }
                else
                {

                    if (itemName.EndsWithIgnoreCase("Power_Binning_P_total"))
                    {
                        double total = powerBinningPTotal + BinCutData.PowerBinningSheetList.First().Value.Offset;
                        if (Math.Abs(log - total) >= 0.002)
                        {
                            BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{total:F4}", itemName);
                            siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                        }
                        CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
                        siteInfoArray[site].PowerBinningPTotalSummary.Add(new Tuple<string, string>(itemName, log.ToString()));
                        UpdateSiteRowIdBySheetResult(siteInfoArray, powerBinningRowIdxs, idx, site, sheetName, rowid, isPtotalFail);
                    }
                }
            }
        }

        private double ComparePowerBinning(StreamWriter streamWriter, SiteInfo[] siteInfoArray, List<BinCutLineBase> binCutLineBases, double powerBinningPTotal, int lnIdx, int lineNo, int site, string itemName, double log, PowerBinningSheetRow rowCg, PowerBinningSheetRow rowPg)
        {
            PowerBinningMain powerBinningMain = new PowerBinningMain();
            rowCg.BinVoltage = rowCg.BinVoltage.Contains("_CPU", StringComparison.OrdinalIgnoreCase) ? rowCg.BinVoltage.Replace("CPU", "ECPU")
                : rowCg.BinVoltage;
            rowCg.Ids = rowCg.Ids.Contains("_CPU", StringComparison.OrdinalIgnoreCase) ? rowCg.Ids.Replace("CPU", "ECPU")
                : rowCg.Ids;
            double product = PowerBinningMain.GetProductValue(rowCg, Job, siteInfoArray[site], streamWriter);
            double expectVbinCg = Math.Round(product, 4, MidpointRounding.AwayFromZero);
            double expectVIdsCg = PowerBinningMain.GetIdsValue(rowCg, siteInfoArray[site]);
            expectVIdsCg = expectVIdsCg == 0.0
                ? siteInfoArray[site].IdsOnList.First(x => x.Key == rowCg.BinVoltage).Value : expectVIdsCg;
            double idsRefCg = rowCg.FactorDictionary["IDS_REF"];
            double sumAcCg = rowCg.FactorDictionary["AC_P1"] + rowCg.FactorDictionary["AC_P2"] + rowCg.FactorDictionary["AC_P3"] + rowCg.FactorDictionary["AC_P4"] + rowCg.FactorDictionary["AC_P5"];
            double sumDcCg = rowCg.FactorDictionary["DC_P1"] + rowCg.FactorDictionary["DC_P2"] + rowCg.FactorDictionary["DC_P3"] + rowCg.FactorDictionary["DC_P4"] + rowCg.FactorDictionary["DC_P5"];
            double vddCalcCg = Math.Pow(product / rowCg.FactorDictionary["VDD_REF"], 2);
            double calcPowerAcCg = vddCalcCg * sumAcCg;

            double expectVbinPg = Math.Round(product, 4, MidpointRounding.AwayFromZero);
            double expectVIdsPg = siteInfoArray[site].IdsOffList.First(x => x.Key == rowPg.BinVoltage).Value;
            double idsRefPg = rowPg.FactorDictionary["IDS_REF"];
            double sumAcPg = rowPg.FactorDictionary["AC_P1"] + rowPg.FactorDictionary["AC_P2"] + rowPg.FactorDictionary["AC_P3"] + rowPg.FactorDictionary["AC_P4"] + rowPg.FactorDictionary["AC_P5"];
            double sumDcPg = rowPg.FactorDictionary["DC_P1"] + rowPg.FactorDictionary["DC_P2"] + rowPg.FactorDictionary["DC_P3"] + rowPg.FactorDictionary["DC_P4"] + rowPg.FactorDictionary["DC_P5"];
            double vddCalcPg = Math.Pow(product / rowPg.FactorDictionary["VDD_REF"], 2);
            double calcPowerAcPg = vddCalcPg * sumAcPg;

            if (itemName.EndsWithIgnoreCase("PG-Vbin_AC"))
            {
                if (Math.Abs(log - expectVbinPg) >= 0.002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectVbinPg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("PG-P-Binned_AC"))
            {
                if (Math.Abs(log - calcPowerAcPg) >= 0.002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{calcPowerAcPg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                powerBinningPTotal += calcPowerAcPg;
                siteInfoArray[site].PowerBinningPTotalSummary.Add(new Tuple<string, string>(itemName, log.ToString()));
                CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("PG-Ids"))
            {
                if (Math.Abs(log - expectVIdsPg) >= 0.002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectVIdsPg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("PG-Vbin_DC"))
            {
                if (Math.Abs(log - expectVbinPg) >= 0.002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectVbinPg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("PG-P-Binned_DC"))
            {
                double idsDcPg = expectVIdsPg;
                double currentCalcDcPg = idsDcPg / idsRefPg;
                double calcPowerDcPg = vddCalcPg * currentCalcDcPg * sumDcPg;
                if (Math.Abs(log - calcPowerDcPg) >= 0.002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{calcPowerDcPg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                powerBinningPTotal += calcPowerDcPg;
                siteInfoArray[site].PowerBinningPTotalSummary.Add(new Tuple<string, string>(itemName, log.ToString()));
                CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("Vbin_AC"))
            {
                if (Math.Abs(log - expectVbinCg) >= 0.002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectVbinCg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("P-Binned_AC"))
            {
                if (Math.Abs(log - calcPowerAcCg) >= 0.002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{calcPowerAcCg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                powerBinningPTotal += calcPowerAcCg;
                siteInfoArray[site].PowerBinningPTotalSummary.Add(new Tuple<string, string>(itemName, log.ToString()));
                CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("Ids"))
            {
                if (Math.Abs(log - expectVIdsCg) >= 0.002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectVIdsCg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("Vbin_DC"))
            {
                if (Math.Abs(log - expectVbinCg) >= 0.002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectVbinCg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("P-Binned_DC"))
            {
                double idsDcCg = expectVIdsCg;
                double currentCalcDcCg = idsDcCg / idsRefCg;
                double calcPowerDcCg = vddCalcCg * currentCalcDcCg * sumDcCg;
                if (Math.Abs(log - calcPowerDcCg) >= 0.002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{calcPowerDcCg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                powerBinningPTotal += calcPowerDcCg;
                siteInfoArray[site].PowerBinningPTotalSummary.Add(new Tuple<string, string>(itemName, log.ToString()));
                CheckManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }

            return powerBinningPTotal;
        }

        protected static void GetPowerBinningNameType(string sheetName, string itemName, string[] spt, ref string pwrName, ref string pwrTypeLog)
        {
            if (itemName.EndsWithIgnoreCase("Total_SRAM_Ids"))
            {
                pwrName = "SRAM";
                pwrTypeLog = "Ids";
            }
            else if (itemName.EndsWithIgnoreCase("Total_SRAM_P_Other"))
            {
                pwrName = "SRAM";
                pwrTypeLog = "P";
            }
            else if (itemName.EndsWithIgnoreCase("Power_Binning_P_total"))
            {
                pwrName = "Total";
                pwrTypeLog = "Power_Binning_P_total";
            }
            else if (itemName.EndsWithIgnoreCase("_Ids"))
            {
                pwrName = spt[2].Replace("_Ids", "");
                pwrTypeLog = "Ids";
            }
            else if (itemName.EndsWithIgnoreCase("_Id"))
            {
                pwrName = spt[2].Replace("_Id", "");
                pwrTypeLog = "Ids";
            }
            else if (itemName.EndsWithIgnoreCase("_Vbin"))
            {
                pwrName = spt[2].Replace("_Vbin", "");
                pwrTypeLog = "Vbin";
            }
            else if (itemName.EndsWithIgnoreCase("_P_Binned"))
            {
                int index = itemName.LastIndexOf("_P_Binned", StringComparison.CurrentCultureIgnoreCase);
                if (index != -1)
                {
                    pwrName = itemName[..index];
                }
                pwrTypeLog = "P";
            }
            else if (itemName.EndsWithIgnoreCase("_P_Other"))
            {
                pwrName = spt[2].Replace("_P_Other", "");
                pwrTypeLog = "P";
            }
            else if (itemName.EndsWithIgnoreCase("_C"))
            {
                pwrName = Reg.RegexC.Replace(spt[2], "");
                pwrTypeLog = "C";
            }
            //check if pwrName is mode
            //if (BinCutData.ModeVsPowerPins.ContainsKey(pwrName))
            //pwrName = BinCutAlgorithmLib.GetPowerByName(pwrName);
            pwrName = pwrName.Replace(sheetName + "_", "");
        }

        protected static int CheckPowerBinningTableNextRow(SiteInfo[] siteInfoArray, int rowid, int site)
        {
            int nextRowid = rowid;
            const string makeBin = "make Bin";
            string bin = siteInfoArray[site].Bin <= BinCutData.BinningTables.Count ? (siteInfoArray[site].Bin - 1).ToString() : "x";

            var harvbins = new List<string>();
            foreach (int harvBin in siteInfoArray[site].HarvBin)
            {
                harvbins.Add(harvBin >= 0 && harvBin < 2 ? harvBin.ToString(CultureInfo.InvariantCulture) : "x");
            }
            if (BinCutData.PowerBinHavrSheet != null)
            {
                if (BinCutData.PowerBinHavrSheet.Rows.ElementAt(rowid).FailCommand.EqualsIgnoreCase("FAIL"))
                {
                    nextRowid = rowid;
                }

                if (BinCutData.PowerBinHavrSheet.Rows.ElementAt(rowid).FailCommand.EqualsIgnoreCase("next"))
                {
                    nextRowid = rowid + 1;
                    string tblbin = BinCutData.PowerBinHavrSheet.Rows.ElementAt(nextRowid).Inputinfo[0].ElementAtOrDefault(0).Value;

                    var tblharvbins = new List<string>();
                    bool binCompare = true;
                    for (int idx = 0; idx < harvbins.Count; idx++)
                    {
                        string tblharvbin = BinCutData.PowerBinHavrSheet.Rows.ElementAt(rowid).Inputinfo[idx + 1].ElementAtOrDefault(0).Value;
                        //if table's bin is x or empty, it shouldn't to compare
                        if (string.IsNullOrEmpty(tblharvbin))
                        {
                            tblharvbin = "x";
                        }

                        tblharvbin = tblharvbin.EqualsIgnoreCase("x") ? harvbins[idx] : tblharvbin;
                        tblharvbins.Add(tblharvbin);
                    }
                    for (int num = 0; num < harvbins.Count; num++)
                    {
                        if (tblharvbins[num] != harvbins[num])
                        {
                            binCompare = false;
                            break;
                        }
                    }
                    if (tblbin == bin && binCompare)
                    {
                        return nextRowid;
                    }
                }

                if (BinCutData.PowerBinHavrSheet.Rows.ElementAt(rowid).FailCommand.StartsWithIgnoreCase("make Bin"))
                {
                    string str = BinCutData.PowerBinHavrSheet.Rows.ElementAt(rowid).FailCommand;
                    string nextBin = str[makeBin.Length..];
                    var testa = BinCutData.PowerBinHavrSheet.Rows.Select(x => x.Inputinfo[0]).ToList();

                    if (nextBin.EqualsIgnoreCase("x"))
                    {
                        nextRowid = testa.FindIndex(x => x.ContainsValue("1"));
                    }

                    if (nextBin.EqualsIgnoreCase("y"))
                    {
                        nextRowid = testa.FindIndex(x => x.ContainsValue("2"));
                    }

                    return nextRowid;
                }
            }
            return nextRowid;
        }

        private void CompareWithHarvPowerBinning(StreamWriter streamWriter, SiteInfo[] siteInfoArray, ref List<BinCutLineBase> binCutLineBases, string sheetName, List<PowerBinningRowIdx> powerBinningRowIdxs)
        {
            // 18100421 0     _Power_Binning_P_total                                                                                                                                  -1       0.0000         178.6793       (F) 80.0000        0.0000         0       
            // 41825472         1     PwrScreen_CS100F_Power_Binning_P_total                                                                                                             -1       N/A            65.7153              N/A            0.0000         0       
            var row = new PowerBinningHarvRow();
            var idx = new PowerBinningRowIdx();
            string sheetInfoSpec = "";
            int rowid = -1;
            for (int lnIdx = 0; lnIdx < binCutLineBases.Count; lnIdx++)
            {
                bool isPtotalFail = false;
                BinCutLineBase line = binCutLineBases[lnIdx];
                int lineNo = binCutLineBases[lnIdx].LineNo;
                string[] arr = line.Line.Split([" K ", " ", "Define", "(F)"], StringSplitOptions.RemoveEmptyEntries);
                if (!double.TryParse(arr[0], out _))
                {
                    continue;
                }

                int site = int.Parse(arr[1]);
                PowerBinningSheet? powerBinning = siteInfoArray[site].PowerBinning;
                if (powerBinning == null)
                {
                    continue;
                }

                int measureIndex = CommonLib.Datalog.LineBase.GetPowerBinningIndex(arr);
                _ = double.TryParse(arr[measureIndex], out double valueLog);
                _ = double.TryParse(arr[measureIndex + 1], out double specLog);

                if (arr[measureIndex + 1] == "m")
                {
                    valueLog /= 1000;
                }
                else if (arr[measureIndex + 1] == "u")
                {
                    valueLog /= 1000000;
                }

                const bool bcheckcurrentsite = true;
                if (powerBinningRowIdxs.Exists(x => x.Site == site && !x.BPass))
                {
                    idx = powerBinningRowIdxs.Find(x => x.Site == site && !x.BPass)!;
                    rowid = idx.RowIdx;
                    row = BinCutData.PowerBinHavrSheet!.Rows.ElementAt(rowid);
                    try
                    {
                        var sh = row.SheetInfo.Find(x => x.ContainsKey(sheetName) && !x.ContainsValue(""))!.ToList();
                        sheetInfoSpec = sh.FirstOrDefault().Value;
                    }
                    catch (Exception)
                    {

                    }
                }

                string itemName = arr[2].Replace(sheetName + "_", "");
                string pwrName = "";
                string pwrTypeLog = "";
                GetPowerBinningNameType(sheetName, itemName, arr, ref pwrName, ref pwrTypeLog);

                if (pwrTypeLog.EqualsIgnoreCase("Power_Binning_P_total") && line.Line.Contains("(F)"))
                {
                    isPtotalFail = true;
                }

                double powerTotal = powerBinning.BinnedModeList.Sum(item => item.PowerValue) + powerBinning.OtherModeList.Sum(item => item.PowerValue);

                double ptotal = Math.Round(powerTotal + powerBinning.Offset, 4, MidpointRounding.AwayFromZero);
                siteInfoArray[site].PowerBinningPTotalSummary.Add(new Tuple<string, string>(sheetName + "_" + itemName, ptotal.ToString()));

                #region Compare Ids,Vbin,P per mode
                PowerBinningSheetRow? powerBinningRow = PowerBinningMain.GetPowerBinningRow(powerBinning, pwrName);
                if (powerBinningRow == null && pwrName != "Total")
                {
                    continue;
                }

                isPtotalFail = CompareAll(streamWriter, siteInfoArray, binCutLineBases, sheetName, powerBinningRowIdxs, row, idx, sheetInfoSpec, rowid, lnIdx, isPtotalFail, lineNo, site, valueLog, specLog, bcheckcurrentsite, itemName, pwrTypeLog, ptotal, powerBinningRow!);
                #endregion
            }
            return;
        }

        public static bool GetHarvestBin(ref OneTouchDown oneTouchDown, ref SiteInfo[] siteInfoArray)
        {
            bool isFoundHarvestBin = false;
            Dictionary<string, int> inputHeader = BinCutData.PowerBinHavrSheet!.InputHeader;
            if (oneTouchDown.Lines.Count == 0)
            {
                return false;
            }

            int i;
            for (i = 0; i < oneTouchDown.Lines.Count; i++)
            {
                if (oneTouchDown.Lines[i].Line.Contains("Power Binning Summary"))
                {
                    return false;
                }

                if (oneTouchDown.Lines[i].Line.Contains("PwrBin Harvest_Bin"))
                {
                    isFoundHarvestBin = true;
                    break;
                }
            }

            if (!isFoundHarvestBin)
            {
                return false;
            }

            i++;
            for (; i < oneTouchDown.Lines.Count; i++)
            {
                if (oneTouchDown.Lines[i].Line.Contains("===    PwrBin Test Name :"))
                {
                    break;
                }

                if (oneTouchDown.Lines[i].Line.Contains("Power Binning Summary"))
                {
                    return false;
                }

                if (oneTouchDown.Lines[i].Line.Contains("Site", StringComparison.CurrentCultureIgnoreCase) &&
                    (oneTouchDown.Lines[i].Line.Contains("HARVESTING_BIN =", StringComparison.OrdinalIgnoreCase) || oneTouchDown.Lines[i].Line.Contains("P_BIN", StringComparison.OrdinalIgnoreCase)))
                {
                    int site = oneTouchDown.Lines[i].GetSite();
                    string[] spt = oneTouchDown.Lines[i].Line.Split([','], StringSplitOptions.RemoveEmptyEntries);
                    if (spt.Length >= 2)
                    {
                        var harvBins = new Dictionary<int, string>();
                        for (int num = 1; num < spt.Length; num++)
                        {
                            string name = spt[num].Split('=').First().Trim().ToUpper();
                            string value = spt[num].Split('=').Last().Trim();
                            if (BinCutConfig.FlagPowerBinningHarvestBinFieldName.Equals(true) || name.Contains("P_BIN", StringComparison.OrdinalIgnoreCase))
                            {
                                if (!inputHeader.TryGetValue(name, out int index))
                                {
                                    continue;
                                }

                                harvBins.Add(index, value);
                            }
                            else
                            {
                                harvBins.Add(0, value);
                            }
                        }
                        IOrderedEnumerable<KeyValuePair<int, string>> harvBinsOrdered = harvBins.OrderBy(x => x.Key);
                        if (harvBinsOrdered.Any())
                        {
                            foreach (KeyValuePair<int, string> harvBin in harvBinsOrdered)
                            {
                                _ = int.TryParse(harvBin.Value.Trim(), out int harvestBin);
                                if (harvestBin >= 0)
                                {
                                    siteInfoArray[site].HarvBin.Add(harvestBin);
                                }
                                else
                                {
                                    BinCutController.Controller.RichTextBoxAppend("=> allDice [{0}] harvest bin: {1}" + site + harvestBin, Color.Red);
                                }
                            }
                        }
                    }
                }
            }

            oneTouchDown.Lines.RemoveRange(0, i);
            return true;
        }
    }
}
