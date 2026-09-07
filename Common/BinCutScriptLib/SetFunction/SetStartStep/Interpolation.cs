using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Reader;
using BinCutScriptLib.Static;

using CommonLib.Extension;

namespace BinCutScriptLib.SetFunction.SetStartStep
{
    public class Interpolation : InheritanceBase
    {
        public static void SetAllPowers(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<InterpolationRow> interpolationRows, InheritanceManager inheritanceManager)
        {
            List<InterpolatioNode> interpolationNodes = BinCutConfig.InterpolationNodes;
            if (interpolationRows.Count == 0 || interpolationNodes.Count == 0)
            {
                return;
            }

            foreach (InterpolationRow row in interpolationRows)
            {
                int site = row.Site;
                if (siteInfoArray[site].AllPowers.Count == 0 && !siteInfoArray[site].IsActiveSite)
                {
                    continue;
                }

                row.CheckeqNLinesSkip(streamWriter);

                GetPwerZones(siteInfoArray, row, site, out PowerZone currentPowerZone, out PowerZone lowPowerZone, out PowerZone highPowerZone1, out PowerZone highPowerZone2);
                if (currentPowerZone == null)
                {
                    return;
                }

                inheritanceManager.SetInheritModeEnable(currentPowerZone.Mode);

                SetByBin(streamWriter, siteInfoArray, interpolationNodes, row, site, currentPowerZone, lowPowerZone, highPowerZone1, highPowerZone2);
            }
        }

        private static void SetByBin(StreamWriter streamWriter, SiteInfo[] siteInfoArray, List<InterpolatioNode> interpolatioNodes, InterpolationRow interpolationRow, int site, PowerZone currentPowerZone, PowerZone lowPowerZone, PowerZone highPowerZone1, PowerZone highPowerZone2)
        {
            int bin = siteInfoArray[site].Bin;
            if (bin > 1)
            {
                //current bin > 1 set step to current bin eq 1
                int binStep = currentPowerZone.PossibleSteps.FindIndex(x => x.Bin.Equals(bin) && x.EqName.Equals(1));
                if (binStep != -1)
                {
                    currentPowerZone.Bin = bin;
                    currentPowerZone.Step = binStep;
                    currentPowerZone.StartStep = binStep;
                    currentPowerZone.FinalStep = binStep;
                    currentPowerZone.StopStep = currentPowerZone.AllSteps.Count - 1;
                    currentPowerZone.SearchStatus = EnumSearchStatus.Search;
                }
            }
            else
            {
                if (!interpolationRow.Bin4Cand)
                {
                    MoveStepByInterpolation(streamWriter, siteInfoArray, currentPowerZone, lowPowerZone, highPowerZone1, highPowerZone2, interpolatioNodes, bin, interpolationRow, site);
                }
                else
                {
                    currentPowerZone.SetBin4CandStep();
                }
            }
        }

        private static void GetPwerZones(SiteInfo[] siteInfoArray, InterpolationRow interpolationRow, int site, out PowerZone currentPowerZone, out PowerZone lowPowerZone, out PowerZone highPowerZone1, out PowerZone highPowerZone2)
        {
            currentPowerZone = null!;
            lowPowerZone = null!;
            highPowerZone1 = null!;
            highPowerZone2 = null!;
            foreach (PowerZone powerZone in siteInfoArray[site].AllPowers)
            {
                if (!string.IsNullOrEmpty(interpolationRow.LowMode) && powerZone.PinMode.Contains(interpolationRow.LowMode))
                {
                    lowPowerZone = powerZone;
                }

                if (!string.IsNullOrEmpty(interpolationRow.HighMode1) && powerZone.PinMode.Contains(interpolationRow.HighMode1))
                {
                    highPowerZone1 = powerZone;
                }

                if (!string.IsNullOrEmpty(interpolationRow.HighMode2) && powerZone.PinMode.Contains(interpolationRow.HighMode2))
                {
                    highPowerZone2 = powerZone;
                }

                if (!string.IsNullOrEmpty(interpolationRow.Mode) && powerZone.PinMode.Contains(interpolationRow.Mode))
                {
                    currentPowerZone = powerZone;
                }
            }
        }

        public static void SetAllPowersCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<InterpolationRow> interpolationRows, InheritanceManager inheritanceManager)
        {
            List<InterpolatioNode> interpolationNodes = BinCutConfig.InterpolationNodes;
            if (interpolationRows.Count == 0 || interpolationNodes.Count == 0)
            {
                return;
            }

            foreach (InterpolationRow row in interpolationRows)
            {
                int site = row.Site;
                if (siteInfoArray[site].AllPowers.Count == 0 && !siteInfoArray[site].IsActiveSite)
                {
                    continue;
                }

                row.CheckeqNLinesSkipCs(streamWriter);
                PowerZone? currentPowerZone = null;
                foreach (PowerZone powerZone in siteInfoArray[site].AllPowers)
                {
                    if (!string.IsNullOrEmpty(row.Mode) && powerZone.PinMode.Contains(row.Mode))
                    {
                        currentPowerZone = powerZone;
                    }
                }
                if (currentPowerZone == null)
                {
                    return;
                }

                inheritanceManager.SetInheritModeEnable(currentPowerZone.Mode);
                int bin = siteInfoArray[site].Bin;
                if (bin > 1)
                {
                    //current bin > 1 set step to current bin eq 1
                    int binStep = currentPowerZone.PossibleSteps.FindIndex(x => x.Bin.Equals(bin) && x.EqName.Equals(1));
                    if (binStep != -1)
                    {
                        currentPowerZone.Bin = bin;
                        currentPowerZone.Step = binStep;
                        currentPowerZone.StartStep = binStep;
                        currentPowerZone.FinalStep = binStep;
                        currentPowerZone.StopStep = currentPowerZone.AllSteps.Count - 1;
                        currentPowerZone.SearchStatus = EnumSearchStatus.Search;
                    }
                }
                else
                {
                    if (!row.Bin4Cand)
                    {
                        MoveStepByInterpolationCs(currentPowerZone, row);
                    }
                    else
                    {
                        currentPowerZone.SetBin4CandStep();
                    }
                }
            }
        }

        private static void MoveStepByInterpolationCs(PowerZone powerZone, InterpolationRow interpolationRow)
        {
            if (powerZone != null)
            {
                for (int i = 0; i < powerZone.AllSteps.Count; i++)
                {
                    PowerStep step = powerZone.AllSteps[i];
                    if (step.EqName == interpolationRow.Eqidx)
                    {
                        powerZone.Step = i;
                        powerZone.FinalStep = i;
                        powerZone.SearchStatus = EnumSearchStatus.Search;
                        break;
                    }
                }
            }
        }

        private static void MoveStepByInterpolation(StreamWriter streamWriter, SiteInfo[] siteInfoArray, PowerZone targetPowerZone, PowerZone lowPowerZone, PowerZone highPowerZone1, PowerZone highPowerZone2, List<InterpolatioNode> interpolatioNodes, int bin, InterpolationRow interpolationRow, int site)
        {
            if (targetPowerZone != null && lowPowerZone != null && highPowerZone1 != null)
            {
                double lowPwrLvcc = lowPowerZone.GetFinalLvcc();
                double highPwrLvcc = highPowerZone1.GetFinalLvcc();
                if (highPowerZone2 != null)
                {
                    highPwrLvcc = highPowerZone1.GetFinalLvcc() > highPowerZone2.GetFinalLvcc()
                        ? highPowerZone1.GetFinalLvcc()
                        : highPowerZone2.GetFinalLvcc();
                }

                double expecteqLvcc = bin != 1 ?
                    targetPowerZone.PossibleSteps.Find(x => x.Bin == bin)!.Lvcc :
                    GetEqNLvccAndOffset(lowPwrLvcc, highPwrLvcc, interpolatioNodes, targetPowerZone);

                bool bSkip = interpolatioNodes.Find(x => x.Mode == targetPowerZone.PinMode)!.Bskip;
                if (true) //New interpolation method , interpolation mode can select binX/binY (interpolation node is skip mode)
                {
                    SetStartStepsNew(ref targetPowerZone, expecteqLvcc, bSkip);
                }
                //else
                //{
                //    SetStartSteps(ref targetPowerZone, expecteqLvcc);
                //    MoveStepByBin(targetPowerZone, bin); //Add for interpolation because interpolation will skip this function
                //}
                if (targetPowerZone.CurrentStep != null)
                {
                    int expectEqIdx = targetPowerZone.CurrentStep.EqName;
                    double expectSelectedLvcc = targetPowerZone.CurrentStep.Lvcc;
                    if (!interpolationRow.SkipTest)
                    {
                        if (interpolationRow.EnValue != expecteqLvcc)
                        {
                            string msg = "The En Voltage of Interpolation mismatch - The interpolated Voltage";
                            BinCutPrint.PrintInterpolationError(streamWriter, msg, siteInfoArray, interpolationRow.EnLine, site, interpolationRow.Mode, expecteqLvcc, interpolationRow.EnValue);
                        }
                        if (interpolationRow.EnLine.Line.Trim().Contains("THE SELECTED EQN", StringComparison.OrdinalIgnoreCase) &&
                            interpolationRow.Eqidx != expectEqIdx)
                        {
                            string msg = "The En Voltage of Interpolation mismatch - THE SELECTED EQN";
                            BinCutPrint.PrintInterpolationError(streamWriter, msg, siteInfoArray, interpolationRow.EnLine, site, interpolationRow.Mode, expectEqIdx, interpolationRow.Eqidx);
                        }
                        if (interpolationRow.EnLine.Line.Trim().Contains("THE SELECTED VOLTAGE", StringComparison.OrdinalIgnoreCase) &&
                            interpolationRow.SelectedLvcc != expectSelectedLvcc)
                        {
                            string msg = "The En Voltage of Interpolation mismatch - THE SELECTED VOLTAGE";
                            BinCutPrint.PrintInterpolationError(streamWriter, msg, siteInfoArray, interpolationRow.EnLine, site, interpolationRow.Mode, expectSelectedLvcc, interpolationRow.SelectedLvcc);
                        }
                    }
                }
            }
        }

        private static void SetStartStepsNew(ref PowerZone powerZone, double eqNLvcc, bool bSkip)
        {
            var possibleSteps = powerZone.PossibleSteps.ToList();
            //Delete steps which bin number is smaller than site bin number
            powerZone.PossibleSteps.ForEach(x => x.IsInterpolationDelete = true);
            PowerStep? firstStep = bSkip ? possibleSteps.Find(x => x.Lvcc >= eqNLvcc) : possibleSteps.Find(x => x.Bin == 1 && x.Lvcc >= eqNLvcc);
            if (firstStep == null)
            {
                if (bSkip) // For skip mode BinOut
                {
                    powerZone.SearchStatus = EnumSearchStatus.BinOut;
                    powerZone.FinalStep = -1;
                    powerZone.Step = -1;
                    powerZone.IsInterplation = true;
                    powerZone.StopStep = powerZone.GetPosCount() - 1;
                    powerZone.IdsMode = EnumIdsMode.Linear;
                    return;
                }
                firstStep = possibleSteps.FindLast(x => x.Bin == 1)!;
            }
            firstStep.IsInterpolationDelete = false;
            IEnumerable<IGrouping<int, PowerStep>> stepByBin = possibleSteps.Where(x => x.Bin >= firstStep.Bin).GroupBy(x => x.Bin);
            foreach (IGrouping<int, PowerStep> powerSteps in stepByBin)
            {
                powerSteps.Last().IsInterpolationDelete = false;
            }
            powerZone.SearchStatus = EnumSearchStatus.Search;
            powerZone.Bin = firstStep.Bin;
            powerZone.FinalStep = 0;
            powerZone.Step = 0;
            powerZone.IsInterplation = true;
            powerZone.StopStep = powerZone.GetPosCount() - 1;
            powerZone.IdsMode = EnumIdsMode.Linear;
        }

        private static double GetEqNLvccAndOffset(double lowPwrLvcc, double highPwrLvcc, List<InterpolatioNode> interpolatioNodes, PowerZone powerZone)
        {
            InterpolatioNode eqn = interpolatioNodes.Find(x => x.Mode.EqualsIgnoreCase(powerZone.PinMode))!;
            double eqNLvcc = lowPwrLvcc + ((highPwrLvcc - lowPwrLvcc) * eqn.Ratio);
            eqNLvcc += eqn.Offset;
            eqNLvcc = BinCutAlgorithmService.Ceiling3P125(eqNLvcc);
            //Remove for use interpolation result only
            //var binStep = powerZone.PossibleSteps.ToList();
            //var max = binStep.Max(y => y.CpVMax);
            //var min = binStep.Min(y => y.CpVMin);
            //if (eqNLvcc > max)
            //    eqNLvcc = max;
            //if (eqNLvcc < min)
            //    eqNLvcc = min;
            return eqNLvcc;
        }
    }
}
