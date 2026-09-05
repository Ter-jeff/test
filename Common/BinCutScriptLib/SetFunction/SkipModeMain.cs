using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Reader;
using BinCutScriptLib.SetFunction.SetStartStep;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib.SetFunction
{
    internal class SkipModeMain
    {
        public static void BeSkipReArrangePowerTableCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<VBinResultWoTestLine1> vBinResultWoTestLine1s, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, List<Tuple<string, string>> skipPwrList)
        {
            ModifyLvccByParMode(siteInfoArray, inheritanceManager, allowEqualBases);
            //if (enRows.Any())
            //{
            //    var skips = enRows.Where(x => x.SkipTest).ToList();
            //    new Interpolation(curInstanceName).SetAllPowers(sw, ref allDice, skips, inheritanceList);
            //}
            if (vBinResultWoTestLine1s.Count == 0 || skipPwrList.Count == 0)
            {
                return;
            }

            CheckResultSkipLinesCs(streamWriter, ref siteInfoArray, ref vBinResultWoTestLine1s, skipPwrList, allowEqualBases);
        }

        public static void BeSkipReArrangePowerTable(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<VBinResultWoTestLine1> vBinResultWoTestLine1s, List<InterpolationRow> interpolationRows, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, List<Tuple<string, string>> skipPwrList, EnumJob enumJob)
        {
            ModifyLvccByParMode(siteInfoArray, inheritanceManager, allowEqualBases);
            if (interpolationRows.Count != 0)
            {
                var skips = interpolationRows.Where(x => x.SkipTest).ToList();
                Interpolation.SetAllPowers(streamWriter, ref siteInfoArray, skips, inheritanceManager);
            }
            if (vBinResultWoTestLine1s.Count == 0 || skipPwrList.Count == 0)
            {
                return;
            }

            CheckResultSkipLines(streamWriter, ref siteInfoArray, ref vBinResultWoTestLine1s, enumJob, skipPwrList, allowEqualBases);
        }

        private static void CheckResultSkipLines(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, ref List<VBinResultWoTestLine1> vBinResultWoTestLine1s, EnumJob enumJob, List<Tuple<string, string>> skipPwrList, List<AllowEqualBase> allowEqualBases)
        {
            List<Tuple<string, string>> skipList = skipPwrList.FindAll(x => x.Item1.EqualsIgnoreCase(enumJob.ToString()));
            foreach (VBinResultWoTestLine1 line in vBinResultWoTestLine1s)
            {
                VBinResultWoTestRow vBinResultWoTestRow = line.GetVBinResultWoTestRow();

                if (string.IsNullOrEmpty(vBinResultWoTestRow.PowerMode) || string.IsNullOrEmpty(vBinResultWoTestRow.AllowEqualPowerMode))
                {
                    if (vBinResultWoTestRow.SetEqn != 0)
                    {
                        for (int site = 0; site < siteInfoArray.Length; site++)
                        {
                            if (siteInfoArray[site].AllPowers.Count == 0 && !siteInfoArray[site].IsActiveSite)
                            {
                                continue;
                            }

                            if (vBinResultWoTestRow.Site != site)
                            {
                                continue;
                            }

                            int pwrIdx =
                                siteInfoArray[site].AllPowers.FindIndex(x => x.PinMode == vBinResultWoTestRow.PowerMode);
                            int stepNum = siteInfoArray[site].AllPowers[pwrIdx].PossibleSteps.FindIndex(
                                x => x.Bin == 1 && x.EqName == vBinResultWoTestRow.SetEqn);
                            siteInfoArray[site].AllPowers[pwrIdx].SearchStatus = EnumSearchStatus.Search;
                            siteInfoArray[site].AllPowers[pwrIdx].FinalStep = stepNum;
                            siteInfoArray[site].AllPowers[pwrIdx].Step = stepNum;
                            siteInfoArray[site].AllPowers[pwrIdx].StopStep = siteInfoArray[site].AllPowers[pwrIdx].GetPosCount() - 1;
                            siteInfoArray[site].AllPowers[pwrIdx].Bin = siteInfoArray[site].AllPowers[pwrIdx].CurrentStep.Bin;
                        }
                    }
                    continue;
                }

                if (allowEqualBases == null)
                {
                    string msg = "Exist vdd bin result skip line but w/o allow equal";
                    streamWriter.WriteLine(msg);
                    msg = $"Fail line:{line.LineNo} ";
                    streamWriter.WriteLine(msg);
                    msg = $"   Skip Mode :{vBinResultWoTestRow.Mode}";
                    streamWriter.WriteLine(msg);
                    msg = $"   allow Mode :{vBinResultWoTestRow.AllowEqualMode}";
                    streamWriter.WriteLine(msg);
                    msg = "";
                    streamWriter.WriteLine(msg);
                    break;
                }

                string aeEqn = allowEqualBases.Any(x => x.Mode.EqualsIgnoreCase(vBinResultWoTestRow.Mode)) ?
                    allowEqualBases.Find(x => x.Mode.EqualsIgnoreCase(vBinResultWoTestRow.Mode))!.AllowEqual : "";
                if (!skipList.Exists(x => x.Item2.EqualsIgnoreCase(vBinResultWoTestRow.Mode)) ||
                    !aeEqn.EqualsIgnoreCase(vBinResultWoTestRow.AllowEqualMode))
                {
                    string msg = "The skip vddBinResult Line mismatch";
                    streamWriter.WriteLine(msg);
                    msg = $"Fail line:{line.LineNo} ";
                    streamWriter.WriteLine(msg);
                    msg = $"   Skip Mode :{vBinResultWoTestRow.Mode}";
                    streamWriter.WriteLine(msg);
                    msg = $"   Allow Mode :{aeEqn}";
                    streamWriter.WriteLine(msg);
                    msg = "";
                    streamWriter.WriteLine(msg);
                }

                for (int site = 0; site < siteInfoArray.Length; site++)
                {
                    if (siteInfoArray[site].AllPowers.Count == 0 && !siteInfoArray[site].IsActiveSite)
                    {
                        continue;
                    }

                    if (vBinResultWoTestRow.Site != site)
                    {
                        continue;
                    }

                    PowerZone? targetPowerZone = null;
                    PowerZone? allowPwrRef = null;
                    foreach (PowerZone pwrRef in siteInfoArray[site].AllPowers)
                    {
                        if (vBinResultWoTestRow.PowerMode.Contains(pwrRef.PinMode))
                        {
                            targetPowerZone = pwrRef;
                        }

                        if (vBinResultWoTestRow.AllowEqualPowerMode.Contains(pwrRef.PinMode))
                        {
                            allowPwrRef = pwrRef;
                        }
                    }

                    if (targetPowerZone == null || allowPwrRef == null)
                    {
                        break;
                    }
                    //link to reference pmode

                    targetPowerZone.AllowPwrRef = allowPwrRef;
                    targetPowerZone.SearchStatus = EnumSearchStatus.Search;
                    double eqNLvcc = allowPwrRef.GetFinalLvcc();
                    for (int stepIdx = 0; stepIdx < targetPowerZone.GetPosCount(); stepIdx++)
                    {
                        if (targetPowerZone.PossibleSteps[stepIdx].Lvcc >= eqNLvcc)
                        {
                            targetPowerZone.Step = stepIdx;
                            targetPowerZone.FinalStep = stepIdx;
                            break;
                        }
                    }
                }
            }
        }

        public static void CheckResultSkipLinesCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, ref List<VBinResultWoTestLine1> vBinResultWoTestLine1s, List<Tuple<string, string>> skipList, List<AllowEqualBase> allowEqualBases)
        {
            foreach (VBinResultWoTestLine1 line in vBinResultWoTestLine1s)
            {
                VBinResultWoTestRow row = line.GetVBinResultWoTestRowCs();
                if (row.SetEqn != 0)
                {
                    SetByVBinResultWoTestRows(siteInfoArray, row);
                }
                else
                {
                    if (string.IsNullOrEmpty(row.PowerMode) || string.IsNullOrEmpty(row.AllowEqualPowerMode))
                    {
                        continue;
                    }

                    if (allowEqualBases == null)
                    {
                        string msg = "Exist vdd bin result skip line but w/o allow equal";
                        streamWriter.WriteLine(msg);
                        msg = $"Fail line:{line.LineNo} ";
                        streamWriter.WriteLine(msg);
                        msg = $"   Skip Mode :{row.Mode}";
                        streamWriter.WriteLine(msg);
                        msg = $"   allow Mode :{row.AllowEqualMode}";
                        streamWriter.WriteLine(msg);
                        msg = "";
                        streamWriter.WriteLine(msg);
                        break;
                    }

                    string aeEqn = allowEqualBases.Any(x => x.Mode.EqualsIgnoreCase(row.Mode)) ? allowEqualBases.Find(x => x.Mode.EqualsIgnoreCase(row.Mode))!.AllowEqual : "";
                    if (!skipList.Exists(x => x.Item2.EqualsIgnoreCase(row.Mode)) ||
                        !aeEqn.EqualsIgnoreCase(row.AllowEqualMode))
                    {
                        string msg = "The skip vddBinResult Line mismatch";
                        streamWriter.WriteLine(msg);
                        msg = $"Fail line:{line.LineNo} ";
                        streamWriter.WriteLine(msg);
                        msg = $"   Skip Mode :{row.Mode}";
                        streamWriter.WriteLine(msg);
                        msg = $"   Allow Mode :{aeEqn}";
                        streamWriter.WriteLine(msg);
                        msg = "";
                        streamWriter.WriteLine(msg);
                    }

                    for (int site = 0; site < siteInfoArray.Length; site++)
                    {
                        if (siteInfoArray[site].AllPowers.Count == 0 && !siteInfoArray[site].IsActiveSite)
                        {
                            continue;
                        }

                        if (row.Site != site)
                        {
                            continue;
                        }

                        PowerZone? targetPowerZone = null;
                        PowerZone? allowPwrRef = null;
                        foreach (PowerZone pwrRef in siteInfoArray[site].AllPowers)
                        {
                            if (row.PowerMode.EqualsIgnoreCase(pwrRef.Mode))
                            {
                                targetPowerZone = pwrRef;
                            }

                            if (row.AllowEqualPowerMode.EqualsIgnoreCase(pwrRef.Mode))
                            {
                                allowPwrRef = pwrRef;
                            }
                        }

                        if (targetPowerZone == null || allowPwrRef == null)
                        {
                            break;
                        }

                        targetPowerZone.AllowPwrRef = allowPwrRef;
                        targetPowerZone.SearchStatus = EnumSearchStatus.Search;
                        double eqNLvcc = allowPwrRef.GetFinalLvcc();
                        for (int stepIdx = 0; stepIdx < targetPowerZone.GetPosCount(); stepIdx++)
                        {
                            if (targetPowerZone.PossibleSteps[stepIdx].Lvcc >= eqNLvcc)
                            {
                                targetPowerZone.Step = stepIdx;
                                targetPowerZone.FinalStep = stepIdx;
                                break;
                            }
                        }
                    }
                }
            }
        }

        private static void SetByVBinResultWoTestRows(SiteInfo[] siteInfoArray, VBinResultWoTestRow vBinResultWoTestRow)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].AllPowers.Count == 0 && !siteInfoArray[site].IsActiveSite)
                {
                    continue;
                }

                if (vBinResultWoTestRow.Site != site)
                {
                    continue;
                }

                int pwrIdx = siteInfoArray[site].AllPowers.FindIndex(x => x.Mode == vBinResultWoTestRow.PowerMode);
                int stepNum = siteInfoArray[site].AllPowers[pwrIdx].PossibleSteps.FindIndex(x => x.Bin == 1 && x.EqName == vBinResultWoTestRow.SetEqn);
                siteInfoArray[site].AllPowers[pwrIdx].SearchStatus = EnumSearchStatus.Search;
                siteInfoArray[site].AllPowers[pwrIdx].FinalStep = stepNum;
                siteInfoArray[site].AllPowers[pwrIdx].Step = stepNum;
                siteInfoArray[site].AllPowers[pwrIdx].StopStep = siteInfoArray[site].AllPowers[pwrIdx].GetPosCount() - 1;
                siteInfoArray[site].AllPowers[pwrIdx].Bin = siteInfoArray[site].AllPowers[pwrIdx].CurrentStep.Bin;
            }
        }

        private static void ModifyLvccByParMode(SiteInfo[] siteInfoArray, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases)
        {
            foreach (List<string> oneInherit in inheritanceManager.GetEnableInheritLists())
            {
                for (int i = 0; i < oneInherit.Count; i++)
                {
                    for (int site = 0; site < siteInfoArray.Length; site++)
                    {
                        if (siteInfoArray[site].AllPowers.Count == 0 || !siteInfoArray[site].IsActiveSite)
                        {
                            continue;
                        }

                        string sonName = oneInherit[i];
                        string parName = "";
                        if (allowEqualBases.Any(x => x.Mode.EqualsIgnoreCase(sonName)))
                        {
                            string allowEqual = allowEqualBases.Find(x => x.Mode.EqualsIgnoreCase(sonName))!.AllowEqual;
                            if (!string.IsNullOrEmpty(allowEqual))
                            {
                                parName = allowEqual;
                            }
                        }

                        siteInfoArray[site].GetInherit(parName, sonName, out int allowEqlParIdx, out int allowEqlSonIdx);

                        if (allowEqlParIdx == -1 || allowEqlSonIdx == -1)
                        {
                            continue;
                        }

                        PowerZone parRef = siteInfoArray[site].AllPowers[allowEqlParIdx];
                        PowerZone sonRef = siteInfoArray[site].AllPowers[allowEqlSonIdx];

                        if (!BinCutData.BinningTables.First().IsTheDomain(parName, sonName))
                        {
                            for (int index = 0; index < sonRef.AllSteps.Count; index++)
                            {
                                PowerStep step = sonRef.AllSteps[index];
                                if (index < parRef.AllSteps.Count)
                                {
                                    step.Lvcc = parRef.AllSteps[index].Lvcc;
                                    step.ProductValue = parRef.AllSteps[index].ProductValue;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
