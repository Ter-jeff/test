using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Reader;
using BinCutScriptLib.Static;

using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib.SetFunction.SetStartStep
{
    public class Inheritance : InheritanceBase
    {
        private enum EnumCode
        {
            Vbt,
            CSharp,
        }

        public static void SetAllPowers(List<int> testedSiteInMode, ref SiteInfo[] siteInfoArray, BvName bvName, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, EnumBcConfig enumBcConfig, List<BvName> bvNames, StreamWriter streamWriter)
        {
            SetAllPowersBase(testedSiteInMode, ref siteInfoArray, bvName, inheritanceManager, allowEqualBases, enumBcConfig, bvNames, streamWriter, isCs: false);
        }

        public static void SetAllPowersCs(List<int> testedSiteInMode, ref SiteInfo[] siteInfoArray, BvName bvName, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, EnumBcConfig enumBcConfig, List<BvName> bvNames, StreamWriter streamWriter)
        {
            SetAllPowersBase(testedSiteInMode, ref siteInfoArray, bvName, inheritanceManager, allowEqualBases, enumBcConfig, bvNames, streamWriter, isCs: true);
        }

        public static void SetAllPowersBase(List<int> testedSiteInMode, ref SiteInfo[] siteInfoArray, BvName bvName, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, EnumBcConfig enumBcConfig, List<BvName> bvNames, StreamWriter streamWriter, bool isCs)
        {
            // 1. Resolve inherit mode lists and mode loop boundaries
            List<string> inheritModeList = isCs ? GetInheritModeListCs(inheritanceManager.GetAllInheritLists(), bvName.Mode) : GetInheritModeList(inheritanceManager.GetAllInheritLists(), bvName.Mode);

            List<string> lastTwoModes = (!isCs && inheritModeList.Count > 2) ? inheritModeList.GetRange(inheritModeList.Count - 2, 2) : inheritModeList;

            int loopLimit = isCs ? lastTwoModes.Count : lastTwoModes.Count - 1;

            // 2. Step inheritance initialization check
            EnumBcConfig stepInheritConfig = isCs ? EnumBcConfig.Debug_BinCutCOF_Stored : EnumBcConfig.Vddbin_COF_StepInheritance_New_Logic;

            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].AllPowers.Count == 0 || !siteInfoArray[site].IsActiveSite)
                {
                    continue;
                }

                SetPowerZone(testedSiteInMode, siteInfoArray, bvName, allowEqualBases, enumBcConfig, bvNames, streamWriter, isCs, inheritModeList, lastTwoModes, loopLimit, stepInheritConfig, site);
            }
        }

        private static void SetPowerZone(List<int> testedSiteInMode, SiteInfo[] siteInfoArray, BvName bvName, List<AllowEqualBase> allowEqualBases, EnumBcConfig method, List<BvName> bvNames, StreamWriter streamWriter, bool isCs, List<string> inheritModeList, List<string> lastTwoModes, int loopLimit, EnumBcConfig stepInheritConfig, int site)
        {
            PowerZone currentPower = siteInfoArray[site].AllPowers[bvName.Index];

            if (method == stepInheritConfig && currentPower.LastFinalStep == -1)
            {
                currentPower.Step = currentPower.StartStep;
                currentPower.FinalStep = currentPower.StartStep;
                currentPower.Bin = 1;
            }

            // 3. Loop through active comparison modes
            for (int i = 0; i < loopLimit; i++)
            {
                PowerZone? parPower = siteInfoArray[site].AllPowers.Find(x => x.Mode == lastTwoModes[i]);
                if (parPower == null || parPower.SearchStatus != EnumSearchStatus.Search)
                {
                    continue;
                }

                bool allowEqual = allowEqualBases.Exists(s => currentPower.PinMode.Contains(s.Mode) && s.AllowEqual.Length != 0 && inheritModeList.Contains(s.AllowEqual));

                double sProductVal;
                if (currentPower.SearchStatus == EnumSearchStatus.Init)
                {
                    currentPower.SearchStatus = EnumSearchStatus.Search;
                    currentPower.FinalStep = 0;
                    sProductVal = currentPower.GetFinalProductValue(streamWriter);
                }
                else
                {
                    sProductVal = currentPower.GetFinalProductValue(streamWriter);
                }
                double sLvcc = currentPower.GetFinalLvcc();

                #region Conventional Method
                if (method == EnumBcConfig.Conventional)
                {
                    HandleConventionalMethod(testedSiteInMode, siteInfoArray, bvName, streamWriter, isCs, site, currentPower, parPower, allowEqual, sProductVal, sLvcc);
                }
                #endregion
                #region Step Inheritance / Stored Method
                else if (method == EnumBcConfig.Debug_BinCutCOF_Stored || method == EnumBcConfig.Vddbin_COF_StepInheritance_New_Logic || method == EnumBcConfig.Vddbin_COF_StepInheritance)
                {
                    double pProductVal = parPower.GetFinalProductValue(streamWriter);
                    double pLvcc = parPower.GetFinalLvcc();

                    if (parPower.IsFail)
                    {
                        PowerStep lastPassStep = parPower.SearchStatus == EnumSearchStatus.Search ? parPower.PossibleSteps[parPower.FinalStep] : siteInfoArray[site].GetFinalPassStep(bvNames);

                        if (lastPassStep != null)
                        {
                            pProductVal = lastPassStep.ProductValue;
                            pLvcc = lastPassStep.Lvcc;
                            currentPower.Bin = lastPassStep.Bin;
                        }
                    }

                    if (isCs)
                    {
                        MoveStepByInheritanceCs(allowEqual, pProductVal, sProductVal, pLvcc, sLvcc, currentPower);
                    }
                    else
                    {
                        MoveStepByInheritance(allowEqual, pProductVal, sProductVal, pLvcc, sLvcc, currentPower);
                    }
                }
                #endregion
            }
        }

        private static void HandleConventionalMethod(List<int> testedSiteInMode, SiteInfo[] siteInfoArray, BvName bvName, StreamWriter streamWriter, bool isCs, int site, PowerZone currentPower, PowerZone parPower, bool allowEqual, double sProductVal, double sLvcc)
        {
            siteInfoArray[site].AllPowers[bvName.Index].Bin = siteInfoArray[site].Bin;
            if (!testedSiteInMode.Exists(x => x == site))
            {
                if (parPower.SearchStatus == EnumSearchStatus.Search)
                {
                    double pProductVal = parPower.GetFinalProductValue(streamWriter);
                    double pLvcc = parPower.GetFinalLvcc();
                    if (currentPower.IdsMode == EnumIdsMode.Ids)
                    {
                        CheckByIdsMode(allowEqual, currentPower, pProductVal, pLvcc);
                    }

                    if (isCs)
                    {
                        MoveStepByInheritanceCs(allowEqual, pProductVal, sProductVal, pLvcc, sLvcc, currentPower);
                    }
                    else
                    {
                        MoveStepByInheritance(allowEqual, pProductVal, sProductVal, pLvcc, sLvcc, currentPower);
                    }
                }
            }
        }

        private static void CheckByIdsMode(bool allowEqual, PowerZone powerZone, double pProductVal, double pLvcc)
        {
            //Rule1
            if (allowEqual)
            {
                if (powerZone.Step != 0 && powerZone.PossibleSteps[powerZone.Step - 1].ProductValue < pProductVal)
                {
                    powerZone.IdsMode = EnumIdsMode.Linear;
                    powerZone.StopStep = powerZone.GetPosCount() - 1;
                    return;
                }
                if (!BinCutConfig.IsCompareByProductValueOnly)
                {
                    if (powerZone.Step != 0 && powerZone.PossibleSteps[powerZone.Step - 1].Lvcc < pLvcc)
                    {
                        powerZone.IdsMode = EnumIdsMode.Linear;
                        powerZone.StopStep = powerZone.GetPosCount() - 1;
                        return;
                    }
                }
            }
            else
            {
                if (powerZone.Step != 0 && powerZone.PossibleSteps[powerZone.Step - 1].ProductValue <= pProductVal)
                {
                    powerZone.IdsMode = EnumIdsMode.Linear;
                    powerZone.StopStep = powerZone.GetPosCount() - 1;
                    return;
                }
                if (!BinCutConfig.IsCompareByProductValueOnly)
                {
                    if (powerZone.Step != 0 && powerZone.PossibleSteps[powerZone.Step - 1].Lvcc <= pLvcc)
                    {
                        powerZone.IdsMode = EnumIdsMode.Linear;
                        powerZone.StopStep = powerZone.GetPosCount() - 1;
                        return;
                    }
                }
            }
            //!!sonRef.stopStep很重要, 這樣就不會把SPI已測的, 只會掃到前次IDS點給覆蓋掉

            //Rule2. IDS mode向下有空間, 但是可能不是每個Eq都大於父電壓時, Stop point必需停在最後一個大於父電壓的Eq(P.Power間的繼承)
            int curRefStep = powerZone.StopStep;
            if (allowEqual)
            {
                for (int stepIdx = powerZone.Step - 1; stepIdx >= curRefStep; stepIdx--)
                {
                    if (BinCutConfig.IsCompareByProductValueOnly)
                    {
                        if (powerZone.PossibleSteps[stepIdx].ProductValue >= pProductVal)
                        {
                            powerZone.StopStep = stepIdx;
                        }
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        if (powerZone.PossibleSteps[stepIdx].ProductValue >= pProductVal && powerZone.PossibleSteps[stepIdx].Lvcc >= pLvcc)
                        {
                            powerZone.StopStep = stepIdx;
                        }
                        else
                        {
                            break;
                        }
                    }
                }

            }
            else
            {
                for (int stepIdx = powerZone.Step - 1; stepIdx >= curRefStep; stepIdx--)
                {
                    if (BinCutConfig.IsCompareByProductValueOnly)
                    {
                        if (powerZone.PossibleSteps[stepIdx].ProductValue > pProductVal)
                        {
                            powerZone.StopStep = stepIdx;
                        }
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        if (powerZone.PossibleSteps[stepIdx].ProductValue > pProductVal && powerZone.PossibleSteps[stepIdx].Lvcc > pLvcc)
                        {
                            powerZone.StopStep = stepIdx;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
        }

        private static void MoveStepByInheritance(bool allowEqual, double pProductVal, double sProductVal, double pLvcc, double sLvcc, PowerZone powerZone)
        {
            while (true)
            {
                // 1. Evaluate the unique VB product evaluation rule
                bool loopFlag;
                if (allowEqual)
                {
                    loopFlag = BinCutConfig.IsCompareByProductValueOnly ? pProductVal > sProductVal : pProductVal > sProductVal || pLvcc > sLvcc;
                }
                else
                {
                    loopFlag = BinCutConfig.IsCompareByProductValueOnly ? pProductVal >= sProductVal : pProductVal >= sProductVal || pLvcc >= sLvcc;
                }

                // 2. Pass execution parameters and loop state down to the engine
                if (!ExecuteStepMoveIteration(loopFlag, powerZone, ref sProductVal, ref sLvcc))
                {
                    break;
                }
            }
        }

        private static void MoveStepByInheritanceCs(bool allowEqual, double pProductVal, double sProductVal, double pLvcc, double sLvcc, PowerZone powerZone)
        {
            while (true)
            {
                double offset = powerZone.GetMonotonicityOffset();

                // 1. Evaluate the unique C# product evaluation rule containing offset mechanics
                bool loopFlag;
                if (allowEqual)
                {
                    loopFlag = BinCutConfig.IsCompareByProductValueOnly ? sProductVal - pProductVal < offset : pProductVal > sProductVal || pLvcc > sLvcc;
                }
                else
                {
                    loopFlag = BinCutConfig.IsCompareByProductValueOnly ? sProductVal - pProductVal <= offset : pProductVal >= sProductVal || pLvcc >= sLvcc;
                }

                // 2. Pass execution parameters and loop state down to the engine
                if (!ExecuteStepMoveIteration(loopFlag, powerZone, ref sProductVal, ref sLvcc))
                {
                    break;
                }
            }
        }

        private static bool ExecuteStepMoveIteration(bool loopFlag, PowerZone powerZone, ref double sProductVal, ref double sLvcc)
        {
            // Abort loop if evaluation rule criteria fails
            if (!loopFlag)
            {
                return false;
            }

            // Check if there are remaining valid tracking channels available
            if (powerZone.Step < powerZone.GetPosCount() - 1)
            {
                powerZone.FinalStep++;
                powerZone.Step++;
                powerZone.IdsMode = EnumIdsMode.Linear;
                powerZone.StopStep = powerZone.GetPosCount() - 1;

                // Dynamically update the lookup arguments for subsequent loop comparisons
                sProductVal = powerZone.PossibleSteps[powerZone.Step].ProductValue;
                sLvcc = powerZone.PossibleSteps[powerZone.Step].Lvcc;

                return true;
            }
            else
            {
                if (!BinCutConfig.IsDoAll)
                {
                    powerZone.SearchStatus = EnumSearchStatus.BinOut;
                    powerZone.IdsMode = EnumIdsMode.Linear;
                }
                return false;
            }
        }

        private static List<string> GetInheritModeList(List<List<string>> inheritLists, string mode)
        {
            return GetInheritModesBase(inheritLists, mode, EnumCode.Vbt);
        }

        private static List<string> GetInheritModeListCs(List<List<string>> inheritLists, string mode)
        {
            return GetInheritModesBase(inheritLists, mode, EnumCode.CSharp);
        }

        private static List<string> GetInheritModesBase(List<List<string>> inheritLists, string mode, EnumCode enumCode)
        {
            var allInheritPwr = new List<string>();
            foreach (List<string> inherit in inheritLists)
            {
                if (!inherit.Contains(mode))
                {
                    continue;
                }
                for (int pwrIdx = 0; pwrIdx < inherit.Count; pwrIdx++)
                {
                    if (enumCode == EnumCode.Vbt)
                    {
                        allInheritPwr.Add(inherit[pwrIdx]);
                    }

                    if (inherit[pwrIdx].Contains(mode))
                    {
                        break;
                    }

                    if (enumCode == EnumCode.CSharp)
                    {
                        allInheritPwr.Add(inherit[pwrIdx]);
                    }
                }
                break;
            }
            return allInheritPwr;
        }

        public static void MoveStepByBin(List<int> testedSiteInMode, ref SiteInfo[] siteInfoArray, BvName bvName, EnumBcConfig enumBcConfig, List<BvName> bvNames)
        {
            MoveStepByBinBase(testedSiteInMode, ref siteInfoArray, bvName, enumBcConfig, bvNames, EnumBcConfig.Vddbin_COF_StepInheritance);
        }

        public static void MoveStepByBinCs(List<int> testedSiteInMode, ref SiteInfo[] siteInfoArray, BvName bvName, EnumBcConfig enumBcConfig, List<BvName> bvNames)
        {
            MoveStepByBinBase(testedSiteInMode, ref siteInfoArray, bvName, enumBcConfig, bvNames, EnumBcConfig.Debug_BinCutCOF_Stored);
        }

        private static void MoveStepByBinBase(List<int> testedSiteInMode, ref SiteInfo[] siteInfoArray, BvName bvName, EnumBcConfig method, List<BvName> bvNames, EnumBcConfig targetMethod)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].AllPowers.Count == 0 || !siteInfoArray[site].IsActiveSite)
                {
                    continue;
                }

                PowerZone currentPower = siteInfoArray[site].AllPowers[bvName.Index];

                if (bvNames != null && bvNames.Count != 0)
                {
                    PowerZone? lastPower = siteInfoArray[site].AllPowers.Find(x => x.Mode == bvNames.Last().Mode);
                    if (lastPower != null)
                    {
                        int bin = siteInfoArray[site].Bin;
                        #region Vddbin_COF_StepInheritance to modify bin by FinalPassStep
                        if (method == targetMethod)
                        {
                            if (!testedSiteInMode.Exists(x => x == site))
                            {
                                //if (!lastPower.IsFail)
                                //{
                                //    bin = lastPower.PossibleSteps[lastPower.FinalStep].Bin;
                                //}
                                if (lastPower.IsFail)
                                {
                                    PowerStep lastPassStep = siteInfoArray[site].GetFinalPassStep(bvNames);
                                    if (lastPassStep != null)
                                    {
                                        if (bin != lastPassStep.Bin)
                                        {
                                            bin = lastPassStep.Bin;
                                            currentPower.Bin = lastPassStep.Bin;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (currentPower.IsFail) //when all eq failing
                                {
                                    PowerStep lastPassStep = siteInfoArray[site].GetFinalPassStep(bvNames);
                                    if (lastPassStep != null)
                                    {
                                        bin = lastPassStep.Bin;
                                    }
                                }
                            }

                        }
                        #endregion

                        MoveStepByBin(currentPower, bin);
                    }
                }
                else
                {
                    //For first mode
                    MoveStepByBin(currentPower, siteInfoArray[site].Bin);
                }
            }
        }

        public static string GetFinalPassMode(List<BvName> bvNames, List<PowerZone> powerZones)
        {
            for (int i = bvNames.Count - 1; i >= 0; i--)
            {
                string mode = bvNames[i].Mode;
                PowerZone powerZone = powerZones.Find(x => x.Mode == mode)!;
                PowerStep? finalPassStep = powerZone.AllSteps.Find(x => x.IsLastPatternPassForCof);
                if (finalPassStep != null)
                {
                    return powerZone.Mode;
                }
            }
            return null!;
        }

        public static void ModifyStepWhenBinout(SiteInfo[] siteInfoArray, BvName bvName, List<BvName> bvNames, bool isNoBinout, bool hasHarvFlag = false, bool cofNew = false)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (!siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                {
                    continue;
                }

                PowerZone currentPower = siteInfoArray[site].AllPowers[bvName.Index];
                //To Set final step when binout
                if (siteInfoArray[site].AllPowers[bvName.Index].SearchStatus == EnumSearchStatus.BinOut)
                {
                    int lastFinalStep = siteInfoArray[site].AllPowers[bvName.Index].LastFinalStep;
                    PowerStep lastPassStep = siteInfoArray[site].GetFinalPassStep(bvNames);
                    string lastPassMode = GetFinalPassMode(bvNames, siteInfoArray[site].AllPowers);
                    if (lastFinalStep != -1) //reference lastFinalStep first
                    {
                        currentPower.Bin = siteInfoArray[site].AllPowers[bvName.Index].AllSteps[lastFinalStep].Bin;
                        currentPower.Step = lastFinalStep;
                        currentPower.FinalStep = lastFinalStep;
                        if (lastPassStep != null)
                        {
                            int bin = lastPassStep.Bin;
                            if (lastPassMode == currentPower.Mode) //if current mode has pattern pass , select pass step.
                            {
                                int step = currentPower.PossibleSteps.FindIndex(x => x.IsLastPatternPassForCof);
                                currentPower.Step = step;
                                currentPower.FinalStep = step;
                                currentPower.Bin = bin;
                                currentPower.SearchStatus = EnumSearchStatus.Search;
                            }
                        }
                    }
                    else // when first mode was failed
                    {
                        if (hasHarvFlag)
                        {
                            int harvFailCount = BinCutConfig.ProjectConfig.GetHarvFailCount(siteInfoArray[site].HarvesFlags);
                            if (harvFailCount == 1) //for case 6
                            {
                                if (siteInfoArray[site].IsPreVddSearch)
                                {
                                    int index = currentPower.PossibleSteps.FindIndex(x => x.Bin == 1 && x.EqName == 1);
                                    currentPower.FinalStep = index;
                                    currentPower.Bin = currentPower.PossibleSteps[index].Bin;
                                    currentPower.Step = index;
                                }
                                else
                                {
                                    currentPower.FinalStep = 0;
                                    currentPower.Bin = currentPower.PossibleSteps[0].Bin;
                                    currentPower.Step = 0;
                                }
                            }
                            else if (harvFailCount > 1) //for case 7
                            {
                                siteInfoArray[site].Bin = currentPower.Bin;
                                siteInfoArray[site].AllPowers[bvName.Index].SearchStatus = EnumSearchStatus.BinOut;
                                return;
                            }
                        }
                        else
                        {
                            MoveStepByBin(currentPower, currentPower.Bin);
                        }
                    }
                    siteInfoArray[site].Bin = currentPower.Bin;
                    siteInfoArray[site].AllPowers[bvName.Index].SearchStatus = EnumSearchStatus.Search;
                    siteInfoArray[site].AllPowers[bvName.Index].IsBinOut = false;
                }
                else
                {
                    currentPower.LastFinalStep = currentPower.FinalStep;
                }
            }
        }
    }
}
