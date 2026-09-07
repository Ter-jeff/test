using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;

using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.Enums;

using TestPlanLib.BinCut;
using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.PowerBinning;

namespace BinCutScriptLib.Comparer.PowerBinning
{
    public class PowerBinningMain
    {
        public static void CreatePowerBinning(PowerBinningSheetCebu powerBinningSheetCebu, SiteInfo[] siteInfoArray)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].AllPowers.Count == 0 || !siteInfoArray[site].IsActiveSite)
                {
                    continue;
                }

                PowerBinningSheetCebu powerbinning = powerBinningSheetCebu.Copy();
                siteInfoArray[site].PowerBinning = powerbinning;
            }
        }

        public static double GetIdsValue(PowerBinningSheetRow powerBinningSheetRow, SiteInfo siteInfo)
        {
            if (!string.IsNullOrEmpty(powerBinningSheetRow.Ids))
            {
                double sum = 0;
                string divider = "";
                string[] items = Reg.RegexSplit1.Split(powerBinningSheetRow.Ids);
                foreach (string item in items)
                {
                    divider = GetOperator(item, divider);
                    PowerZone? row = GetMappingPowerData(siteInfo, item);
                    if (row == null)
                    {
                        continue;
                    }

                    switch (divider)
                    {
                        case "+":
                            {
                                sum += row.IdsValuePowerBinning;
                                break;
                            }
                        case "-":
                            {
                                sum -= row.IdsValuePowerBinning;
                                break;
                            }
                        case "*":
                            {
                                sum *= row.IdsValuePowerBinning;
                                break;
                            }
                        case "/":
                            {
                                sum /= row.IdsValuePowerBinning;
                                break;
                            }
                        default:
                            {
                                sum += row.IdsValuePowerBinning;
                                break;
                            }
                    }
                }
                return sum;
            }
            PowerZone? pwrRef = GetPowerRow(powerBinningSheetRow, siteInfo);
            if (pwrRef != null)
            {
                return pwrRef.IdsValuePowerBinning;
            }

            return 0;
        }

        private static string GetOperator(string item, string divider)
        {
            switch (item)
            {
                case "+":
                    {
                        divider = "+";
                        break;
                    }
                case "-":
                    {
                        divider = "-";
                        break;
                    }
                case "*":
                    {
                        divider = "*";
                        break;
                    }
                case "/":
                    {
                        divider = "/";
                        break;
                    }
                default:
                    {
                        break;
                    }
            }
            return divider;
        }

        public static double GetProductValue(PowerBinningSheetRow powerBinningSheetRow, EnumJob enumJob, SiteInfo siteInfo, StreamWriter streamWriter)
        {
            PowerZone? pwrRef = new PowerZone();
            if (!string.IsNullOrEmpty(powerBinningSheetRow.BinVoltage))
            {
                string mode = BinCutAlgorithmService.GetModeByName(powerBinningSheetRow.BinnedMode);
                var expressList = new List<string>();
                string function = "";
                GetFunction(powerBinningSheetRow.BinVoltage, ref function, ref expressList);

                if (expressList.Count != 0)
                {
                    var doubleList = new List<double>();
                    foreach (string item in expressList)
                    {
                        int pPowerIdx = BinCutAlgorithmService.GetPowerIndex(item + "_" + mode, siteInfo);
                        if (pPowerIdx == -1)
                        {
                            pPowerIdx = BinCutAlgorithmService.GetPowerIndex(item, siteInfo);
                        }

                        if (pPowerIdx != -1)
                        {
                            pwrRef = siteInfo.AllPowers[pPowerIdx];
                        }

                        doubleList.Add(GetOneDiceProductValue(streamWriter, siteInfo, pwrRef, mode));
                    }
                    return GetValue(function, doubleList);
                }

                string name = Combination.CombineByUnderLine(powerBinningSheetRow.BinVoltage, mode);
                int powerIdx = BinCutAlgorithmService.GetPowerIndex(name, siteInfo);
                if (powerIdx == -1)
                {
                    powerIdx = BinCutAlgorithmService.GetPowerIndex(powerBinningSheetRow.BinVoltage, siteInfo);
                }

                if (powerIdx == -1)
                {
                    string nameAffiliated = BinCutData.BinCutFlowTables.First().PowerPins.First(x => x.Value == BinCutData.BinCutFlowTables.First().AffiliatedPin.First(y => y.Key == powerBinningSheetRow.BinVoltage).Value).Key;
                    string modeNum = Reg.RegexModNum.Match(mode).Groups["modeNum"].ToString();
                    string modeAffiliated =
                        BinCutData.AllPowers.Find(x => x.Contains(nameAffiliated) && x.Contains(modeNum)) ?? "";
                    powerIdx = BinCutAlgorithmService.GetPowerIndex(modeAffiliated, siteInfo);
                }
                if (powerIdx == -1)
                {
                    string nameAffiliatedBinVoltage = BinCutData.BinCutFlowTables.First().PowerPins.First(x => x.Value == BinCutData.BinCutFlowTables.First()
                        .AffiliatedPin.First(y => y.Key == powerBinningSheetRow.BinVoltage).Value).Key;
                    powerIdx = BinCutAlgorithmService.GetPowerIndex(nameAffiliatedBinVoltage, siteInfo);
                }
                if (powerIdx != -1)
                {
                    pwrRef = siteInfo.AllPowers[powerIdx];
                }

                return GetOneDiceProductValue(streamWriter, siteInfo, pwrRef, mode);
            }
            pwrRef = GetPowerRow(powerBinningSheetRow, siteInfo);
            if (pwrRef != null)
            {
                return GetOneDiceProductValue(streamWriter, siteInfo, pwrRef);
            }

            return 0;
        }

        public static void SetPowerBinningIds(StreamWriter streamWriter, SiteInfo[] siteInfoArray, ref List<BinCutLineBase> binCutLineBases, string sheetName)
        {
            for (int lnIdx = 0; lnIdx < binCutLineBases.Count; lnIdx++)
            {
                string line = binCutLineBases[lnIdx].Line;
                string[] spt = line.Split([" K ", " ", "Define", "(F)"], StringSplitOptions.RemoveEmptyEntries);
                if (!double.TryParse(spt[0], out double _))
                {
                    continue;
                }

                int site = int.Parse(spt[1]);
                if (siteInfoArray[site].PowerBinning == null)
                {
                    continue;
                }

                string itemName = spt[2].Replace(sheetName + "_", "");
                string pwrName = "";
                string pwrTypeLog = "";
                GetPowerBinningNameType(sheetName, itemName, spt, ref pwrName, ref pwrTypeLog);

                _ = double.TryParse(spt[5], out double log);

                if (pwrTypeLog.EqualsIgnoreCase("Ids"))
                {
                    if (siteInfoArray[site].AllPowers.Exists(x => x.PinMode.EqualsIgnoreCase(pwrName)))
                    {
                        siteInfoArray[site].AllPowers.Find(x => x.PinMode.EqualsIgnoreCase(pwrName))!.IdsValuePowerBinning = log;
                    }
                }
            }
        }

        public void CalculatePowerBinning(SiteInfo[] siteInfoArray, EnumJob enumJob, StreamWriter streamWriter, bool newMethod)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].SiteIsBinOut || !siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                {
                    continue;
                }

                PowerBinningSheet powerbinning = siteInfoArray[site].PowerBinning!;
                foreach (PowerBinningSheetRow row in powerbinning.BinnedModeList)
                {
                    row.IdsValue = GetIdsValue(row, siteInfoArray[site]);
                    row.ProductValue = GetProductValue(row, enumJob, siteInfoArray[site], streamWriter);
                    if (row.FactorDictionary.ContainsKey("Vdd1") || row.FactorDictionary.ContainsKey("Vdd0")) //support power binning new format 201214, table only binned mode
                    {
                        HandlePowerValue(row, newMethod);
                    }
                    else
                    {
                        row.PowerValue = (row.FactorDictionary["A"] + (row.FactorDictionary["B"] * row.IdsValue)) * (row.ProductValue - row.FactorDictionary["C"]) * (row.ProductValue - row.FactorDictionary["C"]) / 1000000;
                    }
                }

                foreach (PowerBinningSheetRow row in powerbinning.OtherModeList)
                {
                    row.IdsValue = GetIdsValue(row, siteInfoArray[site]);
                    row.ProductValue = GetProductValue(row, enumJob, siteInfoArray[site], streamWriter);

                    if (row.FactorDictionary.TryGetValue("D", out double value) && row.FactorDictionary.TryGetValue("E", out double value1))
                    {
                        row.PowerValue = value1 + (value * row.IdsValue);
                    }
                }
            }

            #region for Cebu only
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].AllPowers.Count == 0 || !siteInfoArray[site].IsActiveSite)
                {
                    continue;
                }

                var sramList = new List<PowerBinningSheetRowCebu>();
                if (sramList.Count != 0)
                {
                    var powerbinning = (PowerBinningSheetCebu)siteInfoArray[site].PowerBinning!;
                    powerbinning.OtherModeList.AddRange(sramList);
                }
            }
            #endregion
        }

        public virtual double CalPowerValue(PowerBinningSheetRow powerBinningSheetRow, double deltaX, double vdd1, double vdd0, double jTerm, double k, double exp1, double exp0, double baseValue)
        {
            return CalPowerValueUnified(powerBinningSheetRow, deltaX, vdd1, vdd0, jTerm,
               leakTermCalculator1: () => k * Math.Pow(deltaX, exp1),
               leakTermCalculator0: () => k * Math.Pow(deltaX, exp0),
               elseTermCalculator: () => k * deltaX * Math.Pow(baseValue, deltaX),
               finalReturnCalculator: (jt, lt) => (jt + lt) / 1000000
           );
        }

        public virtual void HandlePowerValue(PowerBinningSheetRow powerBinningSheetRow, bool newMethod)
        {
            if (newMethod)
            {
                Dictionary<string, double> dict = powerBinningSheetRow.FactorDictionary;
                double deltaX = powerBinningSheetRow.ProductValue - dict["C"];
                double vdd1 = dict["Vdd1"];
                double vdd0 = dict["Vdd0"];

                double jTerm = dict["J"] * Math.Pow(deltaX, dict["Exp2"]) / 1000000;

                double leakTerm;
                if (deltaX <= vdd1)
                {
                    leakTerm = dict["K"] * Math.Pow(deltaX / vdd1, dict["Exp1"]) * powerBinningSheetRow.IdsValue;
                }
                else if (deltaX <= vdd0)
                {
                    leakTerm = dict["K"] * Math.Pow(deltaX / vdd1, dict["Exp0"]) * powerBinningSheetRow.IdsValue;
                }
                else
                {
                    double scaling = deltaX * vdd0 / Math.Pow(vdd1, dict["Exp0"]);
                    leakTerm = dict["K"] * scaling * Math.Pow(dict["Base"], deltaX - vdd0) * powerBinningSheetRow.IdsValue;
                }

                powerBinningSheetRow.PowerValue = jTerm + leakTerm;
            }
            else
            {
                Dictionary<string, double> dict = powerBinningSheetRow.FactorDictionary;
                double deltaX = powerBinningSheetRow.ProductValue - dict["C"];
                double jTerm = dict["J"] * Math.Pow(deltaX, dict["Exp2"]);

                double leakTerm;
                if (deltaX <= dict["Vdd1"])
                {
                    leakTerm = dict["K"] * powerBinningSheetRow.IdsValue * Math.Pow(deltaX, dict["Exp1"]);
                }
                else if (deltaX <= dict["Vdd0"])
                {
                    leakTerm = dict["K"] * powerBinningSheetRow.IdsValue * Math.Pow(deltaX, dict["Exp0"]);
                }
                else
                {
                    leakTerm = dict["K"] * powerBinningSheetRow.IdsValue * deltaX * Math.Pow(dict["Base"], deltaX);
                }

                powerBinningSheetRow.PowerValue = (jTerm + leakTerm) / 1000000;
            }
        }

        public static double CalPowerValueUnified(PowerBinningSheetRow powerBinningSheetRow, double deltaX, double vdd1, double vdd0, double jTerm, Func<double> leakTermCalculator1, Func<double> leakTermCalculator0, Func<double> elseTermCalculator, Func<double, double, double> finalReturnCalculator)
        {
            double leakTerm;
            if (deltaX <= vdd1)
            {
                // Inject variation for deltaX calculation (e.g., straight deltaX vs deltaX / vdd1)
                leakTerm = leakTermCalculator1() * powerBinningSheetRow.IdsValue;
            }
            else if (deltaX <= vdd0)
            {
                // Applies variant scaling if injected
                leakTerm = leakTermCalculator0() * powerBinningSheetRow.IdsValue;
            }
            else
            {
                // Inject the different else-block calculation layers dynamically
                leakTerm = elseTermCalculator() * powerBinningSheetRow.IdsValue;
            }

            // Handles the swap between dividing the whole sum vs only dividing jTerm
            return finalReturnCalculator(jTerm, leakTerm);
        }

        public static void AdjustPowerBinningBin(SiteInfo[] siteInfoArray, double ptotal, int site, string sheetName)
        {
            if (BinCutData.PwrbinSeqSheet != null)
            {
                siteInfoArray[site].PowerBinningSeq++;

                if (siteInfoArray[site].Bin == 1 && siteInfoArray[site].PowerBinningSeq == BinCutData.PwrbinSeqSheet.Bin1List.Count)
                {
                    siteInfoArray[site].PowerBinningSeq = -1;
                }

                if (siteInfoArray[site].Bin == 2 && siteInfoArray[site].PowerBinningSeq == BinCutData.PwrbinSeqSheet.BinXList.Count)
                {
                    siteInfoArray[site].PowerBinningSeq = -1;
                }
            }

            int maxBin = BinCutData.BinningTables.Count;
            if (ptotal > siteInfoArray[site].PowerBinning!.Spec)
            {
                if (sheetName.StartsWithIgnoreCase("Power_Binning"))
                {
                    siteInfoArray[site].Bin = maxBin >= 2 ? 2 : 1;
                    siteInfoArray[site].PowerBinningSeq = 0;
                    siteInfoArray[site].PowerBinningFail = "F";
                }
                if (sheetName.StartsWithIgnoreCase("PwrBin1"))
                {
                    siteInfoArray[site].Bin = maxBin >= 2 ? 2 : 1;
                    siteInfoArray[site].PowerBinningSeq = 0;
                    siteInfoArray[site].PowerBinningFail = "F";
                }
                if (sheetName.StartsWithIgnoreCase("PwrBinX"))
                {
                    siteInfoArray[site].Bin = maxBin >= 3 ? 3 : 2;
                    siteInfoArray[site].PowerBinningSeq = 0;
                    siteInfoArray[site].PowerBinningBinXFail = "F";
                }
                foreach (PowerZone pwrRef in siteInfoArray[site].AllPowers)
                {
                    if (pwrRef.PossibleSteps.Count != 0 && pwrRef.FinalStep < pwrRef.GetPosCount() - 1)
                    {
                        while (siteInfoArray[site].Bin > pwrRef.GetFinalBin())
                        {
                            pwrRef.FinalStep++;
                            pwrRef.Step++;
                            pwrRef.SearchStatus = EnumSearchStatus.Search;
                        }
                    }
                }
            }
            else
            {
                if (sheetName.StartsWithIgnoreCase("PwrBinX"))
                {
                    siteInfoArray[site].PowerBinningBinXFail = "P";
                }
                else
                {
                    siteInfoArray[site].PowerBinningFail = "P";
                }
            }
        }

        public static void AdjustPowerBinningHarvBin(SiteInfo[] siteInfoArray, double ptotal, int site, string sheetName, PowerBinningHarvRow powerBinningHarvRow)
        {
            string binNum = "";
            int maxbin = BinCutData.BinningTables.Count;
            if (BinCutData.PowerBinHavrSheet != null)
            {
                const string makeBin = "make Bin";
                if (powerBinningHarvRow.SheetInfo.Count == 0)
                {
                    return;
                }

                if (powerBinningHarvRow.FailCommand.StartsWithIgnoreCase(makeBin))
                {
                    binNum = powerBinningHarvRow.FailCommand[makeBin.Length..];
                }

                string strSpec = powerBinningHarvRow.SheetInfo.Find(x => x.ContainsKey(sheetName))!.Values.First();
                int spec = 0;
                if (!string.IsNullOrEmpty(strSpec))
                {
                    _ = int.TryParse(strSpec, out spec);
                }

                if (ptotal > spec)
                {
                    if (powerBinningHarvRow.FailCommand.StartsWithIgnoreCase(makeBin))
                    {
                        if (binNum.EqualsIgnoreCase("x"))
                        {
                            siteInfoArray[site].Bin = maxbin >= 2 ? 2 : 1;
                            siteInfoArray[site].PowerBinningFail = "F";
                        }
                        if (binNum.EqualsIgnoreCase("y"))
                        {
                            siteInfoArray[site].Bin = maxbin >= 3 ? 3 : 2;
                            siteInfoArray[site].PowerBinningBinXFail = "F";
                        }
                    }
                    if (powerBinningHarvRow.FailCommand.EqualsIgnoreCase("FAIL"))
                    {
                        siteInfoArray[site].Bin = maxbin;
                        siteInfoArray[site].PowerBinningFail = "F";
                        siteInfoArray[site].PowerBinningBinXFail = "F";
                    }

                    foreach (PowerZone pwrRef in siteInfoArray[site].AllPowers)
                    {
                        //Skip non-Search mode
                        if (pwrRef.SearchStatus != EnumSearchStatus.Search)
                        {
                            break;
                        }

                        if (pwrRef.PossibleSteps.Count != 0 && pwrRef.FinalStep < pwrRef.GetPosCount() - 1)
                        {
                            if (BinCutConfig.GradeSearchMethodSelected == "Conventional")
                            {
                                while (siteInfoArray[site].Bin > pwrRef.GetFinalBin())
                                {
                                    pwrRef.FinalStep++;
                                    pwrRef.Step++;
                                }
                            }
                        }

                    }
                }
                else
                {
                    if (binNum.EqualsIgnoreCase("y"))
                    {
                        siteInfoArray[site].PowerBinningBinXFail = "P";
                    }
                    else
                    {
                        siteInfoArray[site].PowerBinningFail = "P";
                    }
                }
            }
        }

        public static void ComparePowerBinningTurks(PowerBinningSheetTurks powerBinningSheetTurks, StreamWriter streamWriter, SiteInfo[] siteInfoArray, ref List<BinCutLineBase> binCutLineBases, EnumJob enumJob, CheckManager checkManager)
        {
            double powerBinningPTotal = 0;
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

                string mode = "";
                foreach (string item in powerBinningSheetTurks.BinnedModeList.Select(x => x.BinnedMode))
                {
                    if (line.ContainsIgnoreCase(item))
                    {
                        mode = item;
                        break;
                    }
                }

                PowerBinningSheetRow? rowCg = powerBinningSheetTurks.BinnedModeList.Find(x => x.BinnedMode.EqualsIgnoreCase(mode));
                PowerBinningSheetRow? rowPg = powerBinningSheetTurks.BinnedModeList.Find(x => x.BinnedMode.EqualsIgnoreCase(mode + "_PG"));
                if (rowCg != null && rowPg != null)
                {
                    double product = GetProductValue(rowCg, enumJob, siteInfoArray[site], streamWriter) / 1000;
                    double expectVbinCg = Math.Round(product, 4, MidpointRounding.AwayFromZero);
                    GetIdsValue(rowCg, siteInfoArray[site]);
                    double idsRefCg = rowCg.FactorDictionary["IDS_REF"];
                    double sumAcCg = rowCg.FactorDictionary["AC_P1"] + rowCg.FactorDictionary["AC_P2"] + rowCg.FactorDictionary["AC_P3"] + rowCg.FactorDictionary["AC_P4"] + rowCg.FactorDictionary["AC_P5"];
                    double sumDcCg = rowCg.FactorDictionary["DC_P1"] + rowCg.FactorDictionary["DC_P2"] + rowCg.FactorDictionary["DC_P3"] + rowCg.FactorDictionary["DC_P4"] + rowCg.FactorDictionary["DC_P5"];
                    double vddCalcCg = Math.Pow(product / rowCg.FactorDictionary["VDD_REF"], 2);
                    double calcPowerAcCg = vddCalcCg * sumAcCg;

                    double expectVbinPg = Math.Round(product, 4, MidpointRounding.AwayFromZero);
                    GetIdsValue(rowPg, siteInfoArray[site]);
                    double idsRefPg = rowPg.FactorDictionary["IDS_REF"];
                    double sumAcPg = rowPg.FactorDictionary["AC_P1"] + rowPg.FactorDictionary["AC_P2"] + rowPg.FactorDictionary["AC_P3"] + rowPg.FactorDictionary["AC_P4"] + rowPg.FactorDictionary["AC_P5"];
                    double sumDcPg = rowPg.FactorDictionary["DC_P1"] + rowPg.FactorDictionary["DC_P2"] + rowPg.FactorDictionary["DC_P3"] + rowPg.FactorDictionary["DC_P4"] + rowPg.FactorDictionary["DC_P5"];
                    double vddCalcPg = Math.Pow(product / rowPg.FactorDictionary["VDD_REF"], 2);
                    double calcPowerAcPg = vddCalcPg * sumAcPg;

                    powerBinningPTotal = CheckPowerBinningLine(streamWriter, siteInfoArray, binCutLineBases, checkManager, powerBinningPTotal, lnIdx, lineNo, site, itemName, log, expectVbinCg, idsRefCg, sumDcCg, vddCalcCg, calcPowerAcCg, expectVbinPg, idsRefPg, sumDcPg, vddCalcPg, calcPowerAcPg);
                }
                else
                {

                    if (itemName.EndsWithIgnoreCase("Power_Binning_P_total"))
                    {
                        double total = powerBinningPTotal + siteInfoArray[site].PowerBinning!.Offset;
                        if (Math.Abs(log - total) >= 0.0002)
                        {
                            BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{total:F4}", itemName);
                            siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                        }
                        checkManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
                        AdjustPowerBinningBin(siteInfoArray, powerBinningPTotal, site, powerBinningSheetTurks.SheetName);
                        siteInfoArray[site].PowerBinningPTotalSummary.Add(new Tuple<string, string>(itemName, log.ToString()));
                    }
                }
            }
        }

        private static double CheckPowerBinningLine(StreamWriter streamWriter, SiteInfo[] siteInfoArray, List<BinCutLineBase> binCutLineBases, CheckManager checkManager, double powerBinningPTotal, int lnIdx, int lineNo, int site, string itemName, double log, double expectVbinCg, double idsRefCg, double sumDcCg, double vddCalcCg, double calcPowerAcCg, double expectVbinPg, double idsRefPg, double sumDcPg, double vddCalcPg, double calcPowerAcPg)
        {
            if (itemName.EndsWithIgnoreCase("pg_Vbin_PG_AC"))
            {
                if (Math.Abs(log - expectVbinPg) >= 0.0002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectVbinPg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                checkManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("pg_P_Binned_PG_AC"))
            {
                if (Math.Abs(log - calcPowerAcPg) >= 0.0002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{calcPowerAcPg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                powerBinningPTotal += calcPowerAcPg;
                siteInfoArray[site].PowerBinningPTotalSummary.Add(new Tuple<string, string>(itemName, log.ToString()));
                checkManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("pg_Ids"))
            {
                checkManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("pg_Vbin_PG_DC"))
            {
                if (Math.Abs(log - expectVbinPg) >= 0.0002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectVbinPg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                checkManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("pg_P_Binned_PG_DC"))
            {
                string title = itemName[..itemName.LastIndexOf("pg_P_Binned_PG_DC", StringComparison.CurrentCultureIgnoreCase)];
                double idsDcPg = GetPowerBinningRow(title + "pg_Ids", binCutLineBases, site);
                double currentCalcDcPg = idsDcPg / idsRefPg;
                double calcPowerDcPg = vddCalcPg * currentCalcDcPg * sumDcPg;
                if (Math.Abs(log - calcPowerDcPg) >= 0.0002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{calcPowerDcPg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                powerBinningPTotal += calcPowerDcPg;
                siteInfoArray[site].PowerBinningPTotalSummary.Add(new Tuple<string, string>(itemName, log.ToString()));
                checkManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("Vbin_AC"))
            {
                if (Math.Abs(log - expectVbinCg) >= 0.0002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectVbinCg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                checkManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("P_Binned_AC"))
            {
                if (Math.Abs(log - calcPowerAcCg) >= 0.0002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{calcPowerAcCg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                powerBinningPTotal += calcPowerAcCg;
                siteInfoArray[site].PowerBinningPTotalSummary.Add(new Tuple<string, string>(itemName, log.ToString()));
                checkManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("Ids"))
            {
                checkManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("Vbin_DC"))
            {
                if (Math.Abs(log - expectVbinCg) >= 0.0002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{expectVbinCg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                checkManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }
            else if (itemName.EndsWithIgnoreCase("P_Binned_DC"))
            {
                string title = itemName[..itemName.LastIndexOf("P_Binned_DC", StringComparison.CurrentCultureIgnoreCase)];
                double idsDcCg = GetPowerBinningRow(title + "Ids", binCutLineBases, site);
                double currentCalcDcCg = idsDcCg / idsRefCg;
                double calcPowerDcCg = vddCalcCg * currentCalcDcCg * sumDcCg;
                if (Math.Abs(log - calcPowerDcCg) >= 0.0002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{log:F4}", $"{calcPowerDcCg:F4}", itemName);
                    siteInfoArray[site].CheckResult.IsPowerBinningPass = false;
                }
                powerBinningPTotal += calcPowerDcCg;
                siteInfoArray[site].PowerBinningPTotalSummary.Add(new Tuple<string, string>(itemName, log.ToString()));
                checkManager.CheckPowerBinningLine(binCutLineBases[lnIdx]);
            }

            return powerBinningPTotal;
        }

        public static PowerBinningSheetRow? GetPowerBinningRow(PowerBinningSheet powerBinningSheet, string powerName)
        {
            // 1. Combine lists once to avoid repeating loops for Binned vs Other
            var allRows = powerBinningSheet.BinnedModeList.Concat(powerBinningSheet.OtherModeList).ToList();

            // 2. Define the matching priority logic in a helper local function
            PowerBinningSheetRow? FindMatch(string name)
            {
                return allRows.FirstOrDefault(row =>
                    name.EqualsIgnoreCase(row.BinnedMode) ||
                    name.EqualsIgnoreCase(GetPowerName(row)) ||
                    name.EqualsIgnoreCase($"VDD_{row.BinnedMode}") ||
                    // Special complex match from your original code
                    name.EqualsIgnoreCase($"{row.Ids}_{row.BinnedMode}") ||
                    name.EqualsIgnoreCase($"{row.Ids}_{BinCutAlgorithmService.GetModeByName(row.BinnedMode)}")
                );
            }

            // 3. First pass with original name
            PowerBinningSheetRow? result = FindMatch(powerName);
            if (result != null)
            {
                return result;
            }

            // 4. Handle special VDD cases and retry
            string? normalizedName = null;
            if (powerName.ContainsIgnoreCase("VDD_FIXED"))
            {
                normalizedName = "VDD_FIXED_ALL";
            }
            else if (powerName.ContainsIgnoreCase("VDD_LOW"))
            {
                normalizedName = "VDD_LOW_ALL";
            }

            return normalizedName != null ? FindMatch(normalizedName) : null;
        }

        private static double GetPowerBinningRow(string title, List<BinCutLineBase> binCutLineBases, int targetSite)
        {
            for (int lnIdx = 0; lnIdx < binCutLineBases.Count; lnIdx++)
            {
                string line = binCutLineBases[lnIdx].Line;
                string[] spt = line.Split([" K ", " ", "Define", "(F)"], StringSplitOptions.RemoveEmptyEntries);

                if (!double.TryParse(spt[0], out _))
                {
                    continue;
                }

                int site = int.Parse(spt[1]);
                string itemName = spt[2];
                _ = double.TryParse(spt[5], out double log);
                if (targetSite == site && title.EqualsIgnoreCase(itemName))
                { return log; }
            }
            return 0;
        }

        private static double GetOneDiceProductValue(StreamWriter streamWriter, SiteInfo siteInfo, PowerZone powerZone, string mode = "")
        {
            double productValue = 0;
            BinningTables otherRailTables = BinCutData.OtherRailTables;
            if (!BinCutData.BinCutFlowTables.Find(x => x.FinalJob.Contains(BinCutData.Job.JobType.ToString()))!.GetEvaluateModes().Contains(mode))
            {
                productValue = siteInfo.EfuseProductValue.TryGetValue(powerZone.PinMode, out double value) ? value : powerZone.GetFinalProductValue(streamWriter);
            }
            else
            {
                bool bstepcheckfail = false;
                if (BinCutConfig.GradeSearchMethodSelected == "PerformanceModes")
                {
                    if (!(powerZone.FinalStep > powerZone.GetPosCount() ||
                          powerZone.FinalStep < 0))
                    {
                        productValue = powerZone.GetFinalProductValue(streamWriter);
                    }
                    else
                    {
                        bstepcheckfail = true;
                    }
                }
                else
                {
                    if (!(powerZone.FinalStep > powerZone.GetPosCount() || powerZone.FinalStep < 0))
                    {
                        productValue = powerZone.GetFinalProductValue(streamWriter);
                    }
                    else
                    {
                        bstepcheckfail = true;
                    }
                }

                if (bstepcheckfail)
                {
                    int binTbIdx = siteInfo.Bin - 1;
                    if (binTbIdx == -1 || binTbIdx >= otherRailTables.Count)
                    {
                    }
                    int rowIdx = BinCutAlgorithmService.GetOtherIndex(mode, otherRailTables[binTbIdx]);
                    if (rowIdx == -1)
                    {
                        rowIdx = BinCutAlgorithmService.GetOtherIndex(powerZone.PinMode, otherRailTables[binTbIdx]);
                    }

                    if (rowIdx == -1)
                    {
                        BinCutController.Controller.RichTextBoxAppend($"Can't find row index of  {powerZone.PinMode} !!!", Color.Red);
                    }
                    else
                    {
                        BinningRow row = otherRailTables[binTbIdx].Rows[rowIdx];
                        if (row != null)
                        {
                            productValue = double.Parse(row.RowData[otherRailTables[binTbIdx].CIdx]) +
                                           double.Parse(row.RowData[otherRailTables[binTbIdx].CpGbIdx]);
                        }
                    }
                }
            }
            return productValue;
        }

        private static void GetPowerBinningNameType(string sheetName, string itemName, string[] spt, ref string pwrName, ref string pwrTypeLog)
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
                pwrName = spt[2].Replace("_P_Binned", "");
                pwrTypeLog = "P";
            }
            else if (itemName.EndsWithIgnoreCase("_P_Other"))
            {
                pwrName = spt[2].Replace("_P_Other", "");
                pwrTypeLog = "P";
            }
            //check if pwrName is mode
            //if (BinCutData.ModeVsPowerPins.ContainsKey(pwrName))
            //    pwrName = BinCutAlgorithmLib.GetPowerByName(pwrName);
            pwrName = pwrName.Replace(sheetName + "_", "");
        }

        private static string GetPowerName(PowerBinningSheetRow powerBinningSheetRow)
        {
            string mode = BinCutAlgorithmService.GetModeByName(powerBinningSheetRow.BinnedMode);
            if (BinCutData.PinInfos.Exists(x => x.Mode.EqualsIgnoreCase(mode)))
            {
                return BinCutData.PinInfos.Find(x => x.Mode.EqualsIgnoreCase(mode))!.Pin + "_" + powerBinningSheetRow.BinnedMode;
            }

            if (BinCutConfig.PowerBiningMappingTable.TryGetValue(powerBinningSheetRow.BinnedMode, out string? name))
            {
                return name;
            }

            return powerBinningSheetRow.BinnedMode;
        }

        private static PowerZone? GetPowerRow(PowerBinningSheetRow powerBinningSheetRow, SiteInfo siteInfo)
        {
            int pPowerIdx = BinCutAlgorithmService.GetPowerIndex(GetPowerName(powerBinningSheetRow), siteInfo);
            if (pPowerIdx != -1)
            {
                return siteInfo.AllPowers[pPowerIdx];
            }

            return null;
        }

        private static PowerZone? GetMappingPowerData(SiteInfo siteInfo, string idsName)
        {
            try
            {
                foreach (PowerZone pwr in siteInfo.AllPowers)
                {
                    string mode = pwr.PinMode.Split('_').LastOrDefault() ?? "";
                    if (Reg.RegexRegexPerformance.IsMatch(mode))
                    {
                        //By pin
                        if (idsName.EqualsIgnoreCase(pwr.PinMode[..pwr.PinMode.LastIndexOf('_')]))
                        {
                            return pwr;
                        }
                    }
                    else
                    {
                        //By Pin_mode
                        if (idsName.EqualsIgnoreCase(pwr.PinMode))
                        {
                            return pwr;
                        }
                    }
                }
                if (idsName.Contains("VDD_FIXED", StringComparison.OrdinalIgnoreCase) || idsName.Contains("VDD_LOW", StringComparison.OrdinalIgnoreCase))
                {
                    if (idsName.Contains("VDD_FIXED", StringComparison.OrdinalIgnoreCase))
                    {
                        idsName = "VDD_FIXED_ALL";
                    }

                    if (idsName.Contains("VDD_LOW", StringComparison.OrdinalIgnoreCase))
                    {
                        idsName = "VDD_LOW_ALL";
                    }

                    foreach (PowerZone pwr in siteInfo.AllPowers)
                    {
                        string mode = pwr.PinMode.Split('_').LastOrDefault() ?? "";
                        if (Reg.RegexRegexPerformance.IsMatch(mode))
                        {
                            //By pin
                            if (idsName.EqualsIgnoreCase(pwr.PinMode[..pwr.PinMode.LastIndexOf('_')]))
                            {
                                return pwr;
                            }
                        }
                        else
                        {
                            //By Pin_mode
                            if (idsName.EqualsIgnoreCase(pwr.PinMode))
                            {
                                return pwr;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }

            return null;
        }

        private static void GetFunction(string binCutExpress, ref string function, ref List<string> expressList)
        {
            if (Reg.RegexFunction.IsMatch(binCutExpress))
            {
                string parameter = Reg.RegexFunction.Match(binCutExpress).Groups["parameter"].ToString();
                function = binCutExpress[..binCutExpress.IndexOf('(')].Trim();
                if (parameter.Contains(','))
                {
                    expressList = [.. parameter.Split(',').Select(x => x.Trim())];
                }
            }
        }

        private static double GetValue(string function, List<double> doubleList)
        {
            if (function.EqualsIgnoreCase("Max"))
            {
                return doubleList.Max();
            }

            if (function.EqualsIgnoreCase("Min"))
            {
                return doubleList.Min();
            }

            return 0;
        }
    }
}
