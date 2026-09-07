using System;
using System.Collections.Generic;

using BinCutScriptLib.Base;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib.SetFunction
{
    internal class CreatePowerStepsMain
    {
        public static void CreatePowerSteps(ref SiteInfo[] siteInfoArray, List<PinInfo> pinInfos, EnumJob enumJob)
        {
            BinningTables binningTables = BinCutData.BinningTables;
            //STEP1. Gather ids from all VddBinDef sheet(include bin1/bin2) to calc ids zone
            //run VDD_SOC_MS001->VDD_SOC_MS002->VDD_GPU_MG002... iteration
            for (int powerIdx = 0; powerIdx < pinInfos.Count; powerIdx++)         //the order of this.powerNames comes from <Judge_stored_IDS>
            {
                if (pinInfos[powerIdx].Binned.EqualsIgnoreCase("TRUE"))
                {
                    HandleVddBin(siteInfoArray, pinInfos, enumJob, binningTables, powerIdx);
                }
                else
                {
                    for (int binTbIdx = 0; binTbIdx < BinCutData.OtherRailTables.Count; binTbIdx++)
                    {
                        BinningTable binTbRef = BinCutData.OtherRailTables[binTbIdx];
                        for (int tbRowIdx = 0; tbRowIdx < binTbRef.Rows.Count; tbRowIdx++)  //all rows in binTable, eg. MC601/MC602/MC603...
                        {
                            string tbModeName = binTbRef.Rows[tbRowIdx].RowData[binTbRef.ModeIdx];
                            if (pinInfos[powerIdx].Mode.EqualsIgnoreCase(tbModeName))
                            {
                                for (int i = 0; i < siteInfoArray.Length; i++)
                                {
                                    var powerStep = new PowerStep
                                    {
                                        Bin = binTbIdx + 1,
                                        Id = double.Parse(binTbRef.Rows[tbRowIdx].RowData[binTbRef.IdIdx]),
                                        C = double.Parse(binTbRef.Rows[tbRowIdx].RowData[binTbRef.CIdx]),
                                        M = double.Parse(binTbRef.Rows[tbRowIdx].RowData[binTbRef.MIdx]),
                                        EqName = int.Parse(binTbRef.Rows[tbRowIdx].RowData[binTbRef.EqnIdx][1..]),
                                        IdsMax = double.Parse(binTbRef.Rows[tbRowIdx].RowData[binTbRef.IdsMaxIdx]),
                                        CpVMax = double.Parse(binTbRef.Rows[tbRowIdx].RowData[binTbRef.CpVMaxIdx]),
                                        CpVMin = double.Parse(binTbRef.Rows[tbRowIdx].RowData[binTbRef.CpVMinIdx]),
                                        Cp1Gb = double.Parse(binTbRef.Rows[tbRowIdx].RowData[binTbRef.CpGbIdx])
                                    };
                                    if (BinCutConfig.FlagT0TxHotFormat)
                                    {
                                        powerStep.NonCp1Gb = double.Parse(binTbRef.Rows[tbRowIdx].RowData[binTbRef.HtolGbHotIdx]);
                                    }
                                    else if (BinCutConfig.FlagT0TxRoomFormat)
                                    {
                                        powerStep.NonCp1Gb = double.Parse(binTbRef.Rows[tbRowIdx].RowData[binTbRef.HtolGbRoomIdx]);
                                    }
                                    else
                                    {
                                        int gbIdx = binTbRef.GetGbIdx(enumJob);
                                        powerStep.NonCp1Gb = double.Parse(binTbRef.Rows[tbRowIdx].RowData[gbIdx]);
                                    }
                                    if (double.TryParse(binTbRef.Rows[tbRowIdx].RowData[binTbRef.CpHvIdx], out double tmpCpHv))
                                    {
                                        powerStep.CpHv = tmpCpHv;
                                    }
                                    else
                                    {
                                        powerStep.CpHv = 0;
                                    }

                                    if (siteInfoArray[i].AllPowers.Count == 0)  //根本沒有量到IDS電流值
                                    {
                                        continue;
                                    }

                                    double pPowerIds = siteInfoArray[i].AllPowers[powerIdx].IdsValue;
                                    // Don't need step size
                                    double lvccCp1 = powerStep.C;
                                    if (lvccCp1 < powerStep.CpVMin)
                                    {
                                        lvccCp1 = powerStep.CpVMin;
                                    }
                                    else if (lvccCp1 > powerStep.CpVMax)
                                    {
                                        lvccCp1 = powerStep.CpVMax;
                                    }

                                    powerStep.BinningProduct = siteInfoArray[i].EFuseValues.Find(x => x.Name != null && x.Name == tbModeName) == null ? 0 : lvccCp1 + powerStep.Cp1Gb;

                                    //Get lvcc/product method depend on Efuse result
                                    powerStep.Lvcc = lvccCp1;
                                    powerStep.ProductValue = siteInfoArray[i].EFuseValues.Find(x => x.Name != null && x.Name == tbModeName) == null ?
                                        powerStep.Lvcc + powerStep.Cp1Gb : powerStep.Lvcc + powerStep.NonCp1Gb;
                                    if (pPowerIds >= powerStep.IdsMax)
                                    {
                                        powerStep.IdsOut = true;
                                    }

                                    siteInfoArray[i].AllPowers[powerIdx].AllSteps.Add(powerStep);
                                }
                            }
                        }
                    }
                }
            }

            //sorting equation, step must go from E4->E3->E2-E1->(Bin2)->E2->E1... something like this
            for (int diceIdx = 0; diceIdx < siteInfoArray.Length; diceIdx++)
            {
                if (siteInfoArray[diceIdx].AllPowers.Count == 0)
                {
                    continue;
                }

                for (int pPowerIdx = 0; pPowerIdx < pinInfos.Count; pPowerIdx++)
                {
                    siteInfoArray[diceIdx].AllPowers[pPowerIdx].AllSteps.Sort();
                    if (!siteInfoArray[diceIdx].AllPowers[pPowerIdx].AllSteps.Exists(x => x.IdsOut.Equals(false)))
                    {
                        siteInfoArray[diceIdx].AllPowers[pPowerIdx].SearchStatus = EnumSearchStatus.BinOut;
                    }
                }
            }
        }

        private static void HandleVddBin(SiteInfo[] siteInfoArray, List<PinInfo> pinInfos, EnumJob enumJob, BinningTables binningTables, int powerIdx)
        {
            #region VddBin
            for (int binTbIdx = 0; binTbIdx < binningTables.Count; binTbIdx++)         //run all bin table iteration, eg. bin1->bin2....
            {
                BinningTable binningTable = binningTables[binTbIdx];
                for (int rowIdx = 0; rowIdx < binningTable.Rows.Count; rowIdx++)  //all rows in binTable, eg. MC601/MC602/MC603...
                {
                    string tbModeName = binningTable.Rows[rowIdx].RowData[binningTable.ModeIdx];
                    if (pinInfos[powerIdx].PinMode.Contains(tbModeName))
                    {
                        for (int i = 0; i < siteInfoArray.Length; i++)
                        {
                            var powerStep = new PowerStep
                            {
                                Bin = binTbIdx + 1
                            };
                            if (binningTables.IsEqnBin)
                            {
                                powerStep.EqnBin = binningTable.Rows[rowIdx].RowData[binningTable.EqnBinIdx];
                            }

                            powerStep.Id = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.IdIdx]);
                            powerStep.C = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.CIdx]);
                            powerStep.M = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.MIdx]);
                            powerStep.EqName = int.Parse(binningTable.Rows[rowIdx].RowData[binningTable.EqnIdx][1..]);
                            powerStep.IdsMax = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.IdsMaxIdx]);
                            powerStep.CpVMax = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.CpVMaxIdx]);
                            powerStep.CpVMin = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.CpVMinIdx]);
                            powerStep.Cp1Gb = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.CpGbIdx]);
                            if (BinCutConfig.FlagT0TxHotFormat)
                            {
                                powerStep.NonCp1Gb = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.HtolGbHotIdx]);
                            }
                            else if (BinCutConfig.FlagT0TxRoomFormat)
                            {
                                powerStep.NonCp1Gb = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.HtolGbRoomIdx]);
                            }
                            else
                            {
                                int gbIdx = binningTable.GetGbIdx(enumJob);
                                powerStep.NonCp1Gb = double.Parse(binningTable.Rows[rowIdx].RowData[gbIdx]);
                            }
                            powerStep.CpHv = double.TryParse(binningTable.Rows[rowIdx].RowData[binningTable.CpHvIdx], out double tmpCpHv) ?
                                tmpCpHv : 0;

                            if (binningTable.SramthreshCp1Idx != -1)
                            {
                                powerStep.SramthreshCp1 = binningTable.Rows[rowIdx].RowData[binningTable.SramthreshCp1Idx];
                            }

                            if (binningTable.SramthreshProductIdx != -1)
                            {
                                powerStep.SramthreshProduct = binningTable.Rows[rowIdx].RowData[binningTable.SramthreshProductIdx];
                            }

                            if (binningTable.MonoDeltaIdx != -1)
                            {
                                powerStep.MonotonicityOffset = binningTable.Rows[rowIdx].RowData[binningTable.MonoDeltaIdx].Length != 0
                                    ? double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.MonoDeltaIdx]) : 0.0;
                            }

                            if (siteInfoArray[i].AllPowers.Count == 0)  //根本沒有量到IDS電流值
                            {
                                continue;
                            }

                            double pPowerIds = siteInfoArray[i].AllPowers[powerIdx].IdsValue;

                            double lvccCp1 = BinCutAlgorithmService.Floor3P125(powerStep.C - (powerStep.M * Math.Log10(pPowerIds)));

                            #region overwrite SRAM Lvcc when SRAM as core power

                            if (binningTable.BincutCalcIdSrailIdx != -1)
                            {
                                if (binningTable.Rows[rowIdx].RowData[binningTable.DomainIdx] != binningTable.Rows[rowIdx].RowData[binningTable.BincutCalcIdSrailIdx])
                                {
                                    string oriMode = siteInfoArray[i].AllPowers[powerIdx].Mode;
                                    string oriModeNum = oriMode[^2..];
                                    PowerZone refMode = siteInfoArray[i].AllPowers.Find(x => x.Pin == "VDD_" + binningTable.Rows[rowIdx].RowData[binningTable.BincutCalcIdSrailIdx]
                                        && x.Mode[^2..] == oriModeNum)!;
                                    lvccCp1 = BinCutAlgorithmService.Floor3P125(powerStep.C - (powerStep.M * Math.Log10(refMode.IdsValue)));
                                    siteInfoArray[i].AllPowers[powerIdx].IdsValue = refMode.IdsValue;
                                    powerStep.IdsMax = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.BincutCalcIdSmax]);
                                }
                            }

                            #endregion
                            if (lvccCp1 < powerStep.CpVMin)
                            {
                                lvccCp1 = powerStep.CpVMin;
                            }
                            else if (lvccCp1 > powerStep.CpVMax)
                            {
                                lvccCp1 = powerStep.CpVMax;
                            }

                            powerStep.BinningProduct = lvccCp1 + powerStep.Cp1Gb;

                            //Get lvcc/product method depend on Efuse result
                            powerStep.Lvcc = lvccCp1;
                            powerStep.ProductValue = siteInfoArray[i].EFuseValues.Find(x => x.Name != null && x.Name == tbModeName) == null ?
                                powerStep.Lvcc + powerStep.Cp1Gb : powerStep.Lvcc + powerStep.NonCp1Gb;

                            if (pPowerIds >= powerStep.IdsMax)
                            {
                                powerStep.IdsOut = true;
                            }

                            siteInfoArray[i].AllPowers[powerIdx].AllSteps.Add(powerStep);
                        }
                    }
                }
            }
            #endregion
        }
    }
}
