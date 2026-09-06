using System;
using System.Collections.Generic;
using System.IO;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Reader;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib.SetFunction.SetStartStep
{
    internal class SetStartStepMain
    {
        public static void WorkFlow(InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, EnumBcConfig enumBcConfig, string curInstanceName, StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<InterpolationRow> interpolationRows, BvName bvName, List<BvName> bvNames)
        {
            WorkFlowBase(inheritanceManager, allowEqualBases, enumBcConfig, curInstanceName, streamWriter, ref siteInfoArray, interpolationRows, bvName, bvNames, isCs: false);
        }

        public static void WorkFlowCs(InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, EnumBcConfig enumBcConfig, string curInstanceName, StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<InterpolationRow> interpolationRows, List<VBinResultWoTestLine1> vBinResultWoTestLine1s, BvName bvName, List<BvName> bvNames, List<Tuple<string, string>> skipPwrList)
        {
            WorkFlowBase(inheritanceManager, allowEqualBases, enumBcConfig, curInstanceName, streamWriter, ref siteInfoArray, interpolationRows, bvName, bvNames, isCs: true, vBinResultWoTestLine1s, skipPwrList);
        }

        public static void WorkFlowBase(InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, EnumBcConfig enumBcConfig, string curInstanceName, StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<InterpolationRow> interpolationRows, BvName bvName, List<BvName> bvNames, bool isCs, List<VBinResultWoTestLine1> vBinResultWoTestLine1s = null!, List<Tuple<string, string>> skipPwrList = null!)
        {
            // 1. Extract tested sites
            List<int> testedSiteInMode = bvNames.Exists(x => x.Mode.EqualsIgnoreCase(bvName.Mode))
                ? bvNames.Find(x => x.Mode.EqualsIgnoreCase(bvName.Mode))!.Sites
                : [];

            // 2. Initialize
            InitialSteps(siteInfoArray, bvName, testedSiteInMode);

            // 3. Process IDS Distribution
            bool flagIdsEnable = isCs ? BinCutConfig.StartEqn : BinCutConfig.FlagIdsDistributionEnable;
            if (BinCutData.IdsdistributionTable != null && flagIdsEnable)
            {
                var idsDistribution = new IdsDistribution();
                IdsDistribution.SetAllPowers(testedSiteInMode, ref siteInfoArray, bvName, BinCutData.IdsdistributionTable, curInstanceName, flagIdsEnable);
            }

            // 4. Process Interpolation
            if (interpolationRows.Count != 0)
            {
                if (isCs)
                {
                    Interpolation.SetAllPowersCs(streamWriter, ref siteInfoArray, interpolationRows, inheritanceManager);
                }
                else
                {
                    Interpolation.SetAllPowers(streamWriter, ref siteInfoArray, interpolationRows, inheritanceManager);
                }
            }

            // 5. Process CS Specific Skip Logic
            if (isCs)
            {
                SkipModeMain.CheckResultSkipLinesCs(streamWriter, ref siteInfoArray, ref vBinResultWoTestLine1s, skipPwrList, allowEqualBases);
            }

            // 6. Process Inheritance
            if (isCs)
            {
                Inheritance.SetAllPowersCs(testedSiteInMode, ref siteInfoArray, bvName, inheritanceManager, allowEqualBases, enumBcConfig, bvNames, streamWriter);
                Inheritance.MoveStepByBinCs(testedSiteInMode, ref siteInfoArray, bvName, enumBcConfig, bvNames);
            }
            else
            {
                Inheritance.SetAllPowers(testedSiteInMode, ref siteInfoArray, bvName, inheritanceManager, allowEqualBases, enumBcConfig, bvNames, streamWriter);
                Inheritance.MoveStepByBin(testedSiteInMode, ref siteInfoArray, bvName, enumBcConfig, bvNames);
            }
        }

        private static void InitialSteps(SiteInfo[] siteInfoArray, BvName bvName, List<int> testedSiteInMode)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].AllPowers.Count == 0 || !siteInfoArray[site].IsActiveSite)
                {
                    continue;
                }

                PowerZone currentPower = siteInfoArray[site].AllPowers[bvName.Index];
                if (!testedSiteInMode.Exists(x => x == site) && currentPower.SearchStatus == EnumSearchStatus.Init)
                {
                    InitialStep(currentPower);
                }
            }
        }

        public static void InitialStep(PowerZone powerZone)
        {
            powerZone.SearchStatus = EnumSearchStatus.Search;
            powerZone.Step = 0;
            powerZone.FinalStep = 0;
            powerZone.StartStep = 0;
            powerZone.StopStep = powerZone.GetPosCount() - 1;
            powerZone.IdsMode = EnumIdsMode.Linear;
            powerZone.Bin = 1;
        }
    }
}
