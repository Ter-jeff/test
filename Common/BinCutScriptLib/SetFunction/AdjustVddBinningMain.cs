using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Reader;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib.SetFunction
{
    internal class AdjustVddBinningMain
    {
        #region By Bin
        public static void AdjustByBin(ref SiteInfo[] siteInfoArray)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].AllPowers.Count == 0 || !siteInfoArray[site].IsActiveSite)
                {
                    continue;
                }

                AdjustOneSiteByBin(siteInfoArray, site);
            }
        }

        private static void AdjustOneSiteByBin(SiteInfo[] siteInfoArray, int site)
        {
            if (siteInfoArray[site].Bin != 1)
            {
                for (int pPowerIdx = 0; pPowerIdx < siteInfoArray[site].AllPowers.Count; pPowerIdx++)
                {
                    PowerZone pwrRef = siteInfoArray[site].AllPowers[pPowerIdx];
                    if (pwrRef.GetPosCount() == 0)
                    {
                        continue;
                    }

                    if (pwrRef.SearchStatus == EnumSearchStatus.BinOut)
                    {
                        continue;
                    }

                    while (pwrRef.FinalStep < pwrRef.GetPosCount() - 1 && siteInfoArray[site].Bin > pwrRef.Bin)
                    {
                        pwrRef.FinalStep++;
                        pwrRef.Step++;
                        //200410 Mark apply new rule. Align all power perf mode bin to wrost bin. Case1 had this rule
                        pwrRef.Bin = pwrRef.PossibleSteps[pwrRef.FinalStep].Bin;
                        //binOut()後還有可能動到product value就這裡, 所以下面比較Product Value的Code還是寫成直接讀取LVCC及GB後相加
                    }
                }
            }
        }
        #endregion

        #region By AllowEqual
        public static void AdjustByAllowEqual(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases)
        {
            AdjustProductValByMaxPv(ref siteInfoArray, allowEqualBases, streamWriter);
            AdjustByAllowEqualStepBase(ref siteInfoArray, inheritanceManager, allowEqualBases, streamWriter, false);
        }

        public static void AdjustByAllowEqualCsharp(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases)
        {
            AdjustProductValByMaxPv(ref siteInfoArray, allowEqualBases, streamWriter);
            AdjustByAllowEqualStepBase(ref siteInfoArray, inheritanceManager, allowEqualBases, streamWriter, true);
        }

        private static void AdjustByAllowEqualStepBase(ref SiteInfo[] siteInfoArray, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, StreamWriter streamWriter, bool isCsharp)
        {
            foreach (List<string> oneInherit in inheritanceManager.GetEnableInheritLists())
            {
                for (int i = 0; i < oneInherit.Count - 1; i++)
                {
                    for (int site = 0; site < siteInfoArray.Length; site++)
                    {
                        if (siteInfoArray[site].AllPowers.Count == 0 || !siteInfoArray[site].IsActiveSite)
                        {
                            continue;
                        }

                        string parName = oneInherit[i];
                        string sonName = oneInherit[i + 1];

                        siteInfoArray[site].GetInherit(parName, sonName, out int allowEqlParIdx, out int allowEqlSonIdx);

                        if (allowEqlParIdx == -1 || allowEqlSonIdx == -1)
                        {
                            continue;
                        }

                        PowerZone parRef = siteInfoArray[site].AllPowers[allowEqlParIdx];
                        PowerZone sonRef = siteInfoArray[site].AllPowers[allowEqlSonIdx];

                        if (parRef.SearchStatus != EnumSearchStatus.Search || sonRef.SearchStatus != EnumSearchStatus.Search)
                        {
                            continue;
                        }

                        if (BinCutConfig.FlagUseNewInterpolationMonotonicity.Equals(true) && !isCsharp)
                        {
                            //initialize adjustMode
                            if (!parRef.IsAdjust)
                            {
                                parRef.SetAdjustTest(true);
                            }

                            if (!sonRef.IsAdjust)
                            {
                                sonRef.SetAdjustTest(true);
                            }
                        }

                        double pLvcc = parRef.GetFinalLvcc();
                        double sLvcc = sonRef.GetFinalLvcc();
                        double pProductVal = parRef.GetFinalProductValue(streamWriter);
                        double sProductVal = sonRef.GetFinalProductValue(streamWriter);
                        double monoDelta = sonRef.PossibleSteps[sonRef.FinalStep].MonotonicityOffset;
                        bool isallowEqual = allowEqualBases.Exists(s => sonRef.PinMode.Contains(s.Mode));
                        if (BinCutConfig.FlagCrossDomainCheckEnable.Equals(true))
                        {
                            if (parRef.Pin.Contains("VDD_GPU"))
                            {
                                int modeLength = parRef.Mode.Length;
                                string parentModeNum = parRef.Mode[(modeLength - 2)..];
                                pLvcc = siteInfoArray[site].AllPowers.FindAll(x => Reg.RegexPin.IsMatch(x.Pin) && x.Mode.Length != 0).FindAll(x => x.Mode[^2..] == parentModeNum).Max(x => x.GetFinalLvcc());
                                pProductVal = siteInfoArray[site].AllPowers.FindAll(x => Reg.RegexPin.IsMatch(x.Pin) && x.Mode.Length != 0).FindAll(x => x.Mode[^2..] == parentModeNum).Max(x => x.GetFinalProductValue(streamWriter));
                            }
                        }
                        sonRef.MonoAdjust = isallowEqual ? (sProductVal - pProductVal < monoDelta) : (sProductVal - pProductVal <= monoDelta);
                        sProductVal -= monoDelta;
                        MoveStep(isallowEqual, pProductVal, sProductVal, pLvcc, sLvcc, ref sonRef, streamWriter);
                        if (sonRef.GetFinalBin() > siteInfoArray[site].Bin)
                        {
                            sonRef.Bin = sonRef.GetFinalBin();
                            siteInfoArray[site].Bin = sonRef.GetFinalBin();
                        }

                        SetAdjustTest(isCsharp, parRef, sonRef);
                    }
                }
            }
        }

        private static void SetAdjustTest(bool isCsharp, PowerZone parRef, PowerZone sonRef)
        {
            bool flag = isCsharp ? BinCutConfig.FlagEnablePreAdjustVddbinning.Equals(true) : (BinCutConfig.FlagUseNewInterpolationMonotonicity.Equals(true) && BinCutConfig.FlagEnablePreAdjustVddbinning.Equals(true));
            //Remove BinCutConfig.Flag_use_new_Interpolation_Monotonicity.Equals(true)
            if (flag)
            {
                if (parRef.IsAdjust)
                {
                    parRef.SetAdjustTest(false);
                }

                if (sonRef.IsAdjust)
                {
                    sonRef.SetAdjustTest(false);
                }
            }
        }

        private static void AdjustProductValByMaxPv(ref SiteInfo[] siteInfoArray, List<AllowEqualBase> allowEqualBases, StreamWriter streamWriter)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].AllPowers.Count == 0 || !siteInfoArray[site].IsActiveSite)
                {
                    continue;
                }

                for (int allowEqualIdx = allowEqualBases.Count - 1; allowEqualIdx >= 0; allowEqualIdx--)
                {
                    if (!allowEqualBases[allowEqualIdx].Comment.StartsWithIgnoreCase("Max PV"))
                    {
                        continue;
                    }

                    string parName = allowEqualBases[allowEqualIdx].AllowEqual;
                    string sonName = allowEqualBases[allowEqualIdx].Mode;

                    siteInfoArray[site].GetInherit(parName, sonName, out int allowEqlParIdx, out int allowEqlSonIdx);

                    if (allowEqlParIdx == -1 || allowEqlSonIdx == -1)
                    {
                        continue;
                    }

                    PowerZone parRef = siteInfoArray[site].AllPowers[allowEqlParIdx];
                    PowerZone sonRef = siteInfoArray[site].AllPowers[allowEqlSonIdx];

                    if (parRef.SearchStatus != EnumSearchStatus.Search || sonRef.SearchStatus != EnumSearchStatus.Search)
                    {
                        continue;
                    }

                    double mcParProductVal = parRef.GetFinalProductValue(streamWriter);
                    double mcSonProductVal = sonRef.GetFinalProductValue(streamWriter);

                    if (mcParProductVal != mcSonProductVal)
                    {
                        double larger = mcParProductVal > mcSonProductVal ? mcParProductVal : mcSonProductVal;
                        if (parRef.PossibleSteps[parRef.FinalStep].BinningProduct != 0 &&
                            sonRef.PossibleSteps[sonRef.FinalStep].BinningProduct != 0)
                        {
                            //parRef.PossibleSteps[parRef.FinalStep].BinningProduct = larger;
                            //sonRef.PossibleSteps[sonRef.FinalStep].BinningProduct = larger;
                            siteInfoArray[site].AllPowers[allowEqlParIdx].PossibleSteps[parRef.FinalStep].BinningProduct =
                                larger;
                            siteInfoArray[site].AllPowers[allowEqlSonIdx].PossibleSteps[sonRef.FinalStep].BinningProduct =
                                larger;
                        }
                        else
                        {
                            //parRef.PossibleSteps[parRef.FinalStep].ProductValue = larger;
                            //sonRef.PossibleSteps[sonRef.FinalStep].ProductValue = larger;
                            siteInfoArray[site].AllPowers[allowEqlParIdx].PossibleSteps[parRef.FinalStep].ProductValue =
                                larger;
                            siteInfoArray[site].AllPowers[allowEqlSonIdx].PossibleSteps[sonRef.FinalStep].ProductValue =
                                larger;
                        }
                    }
                }
            }
        }

        private static void MoveStep(bool allowEqual, double pProductVal, double sProductVal, double pLvcc, double sLvcc, ref PowerZone powerZone, StreamWriter streamWriter)
        {
            bool loopBreak = true;
            while (loopBreak)
            {
                if (allowEqual)
                {
                    loopBreak = BinCutConfig.IsCompareByProductValueOnly ? pProductVal > sProductVal : pProductVal > sProductVal || pLvcc > sLvcc;
                }
                else
                {
                    loopBreak = BinCutConfig.IsCompareByProductValueOnly ? pProductVal >= sProductVal : pProductVal >= sProductVal || pLvcc >= sLvcc;
                }

                if (!loopBreak)
                {
                    break;
                }

                if (powerZone.FinalStep < powerZone.GetPosCount() - 1) //尚未至最後一步
                {
                    powerZone.FinalStep++;
                    if (powerZone.FinalStep < powerZone.GetPosCount())
                    {
                        //reCal production value
                        sProductVal = powerZone.GetFinalProductValue(streamWriter) - powerZone.PossibleSteps[powerZone.FinalStep].MonotonicityOffset;
                        sLvcc = powerZone.GetFinalLvcc();
                    }
                    else
                    {
                        powerZone.FinalStep = powerZone.GetPosCount() - 1;
                        break;
                    }
                }
                else
                {
                    if (!BinCutConfig.IsDoAll)
                    {
                        //若finalStep=-1表示繼承關係找不到任何更大的電壓, finalStep填為-1
                        //sonRef.FinalStep = -1;
                        powerZone.SearchStatus = EnumSearchStatus.BinOut;
                    }
                    break;
                }
            }
        }
        #endregion

        public static void ComparePre_Adjust_VddBinningVbt(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, OneTouchDown oneTouchDown, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases)
        {
            siteInfoArray = ComparePre_Adjust_VddBinningBase(streamWriter, siteInfoArray, oneTouchDown, inheritanceManager, allowEqualBases, false);
        }

        public static void ComparePre_Adjust_VddBinningCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, OneTouchDown oneTouchDown, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases)
        {
            siteInfoArray = ComparePre_Adjust_VddBinningBase(streamWriter, siteInfoArray, oneTouchDown, inheritanceManager, allowEqualBases, true);
        }

        private static SiteInfo[] ComparePre_Adjust_VddBinningBase(StreamWriter streamWriter, SiteInfo[] siteInfoArray, OneTouchDown oneTouchDown, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, bool isCsharp)
        {
            if (isCsharp)
            {
                #region Compare <pre_Adjust_VddBinning>
                var adjustVddBinning = new AdjustVddBinningMain();
                BinCutConfig.FlagEnablePreAdjustVddbinning = oneTouchDown.Lines.Any(x => x.Line == "----------Pre_Adjust_Vddbinning Start----------") && oneTouchDown.Lines.Any(y => y.Line == "----------Pre_Adjust_Vddbinning End----------");
                if (BinCutConfig.FlagEnablePreAdjustVddbinning.Equals(true))
                {
                    bool binChange = false;
                    int[] siteBinList = new int[siteInfoArray.Length];

                    for (int i = 0; i < siteBinList.Length; i++)
                    {
                        siteBinList[i] = siteInfoArray[i].Bin;
                    }
                    while (true)
                    {
                        var preadjustVddBinning = new AdjustVddBinningMain();
                        AdjustByBin(ref siteInfoArray);
                        AdjustByAllowEqualCsharp(streamWriter, ref siteInfoArray, inheritanceManager, allowEqualBases);
                        for (int i = 0; i < siteBinList.Length; i++)
                        {
                            if (!siteBinList[i].Equals(siteInfoArray[i].Bin))
                            {
                                binChange = true;
                                siteBinList[i] = siteInfoArray[i].Bin;
                            }
                        }
                        if (!binChange)
                        {
                            break;
                        }

                        binChange = false;
                    }
                }
                else
                {
                    AdjustByBin(ref siteInfoArray);
                }
                #endregion
            }
            else
            {
                #region Compare <pre_Adjust_VddBinning>
                BinCutConfig.FlagEnablePreAdjustVddbinning = oneTouchDown.Lines.Any(x => x.Line == "----------Pre_Adjust_Vddbinning Start----------") && oneTouchDown.Lines.Any(y => y.Line == "----------Pre_Adjust_Vddbinning End----------");
                if (BinCutConfig.FlagEnablePreAdjustVddbinning.Equals(true))
                {
                    bool binChange = false;
                    int[] siteBinList = new int[siteInfoArray.Length];

                    for (int i = 0; i < siteBinList.Length; i++)
                    {
                        siteBinList[i] = siteInfoArray[i].Bin;
                    }
                    while (true)
                    {
                        var preadjustVddBinning = new AdjustVddBinningMain();
                        AdjustByBin(ref siteInfoArray);
                        AdjustByAllowEqual(streamWriter, ref siteInfoArray, inheritanceManager, allowEqualBases);
                        for (int i = 0; i < siteBinList.Length; i++)
                        {
                            if (!siteBinList[i].Equals(siteInfoArray[i].Bin))
                            {
                                binChange = true;
                                siteBinList[i] = siteInfoArray[i].Bin;
                            }
                        }
                        if (!binChange)
                        {
                            break;
                        }

                        binChange = false;
                    }
                }
                else
                {
                    AdjustByBin(ref siteInfoArray);
                }
                #endregion
            }

            return siteInfoArray;
        }

        public static void CompareAdjust_VddBinningCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, OneTouchDown oneTouchDown, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, EnumJob enumJob)
        {
            siteInfoArray = CompareAdjust_VddBinningBase(streamWriter, siteInfoArray, oneTouchDown, inheritanceManager, allowEqualBases, enumJob, true);
        }

        internal static void CompareAdjust_VddBinningVbt(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, OneTouchDown oneTouchDown, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, EnumJob enumJob)
        {
            siteInfoArray = CompareAdjust_VddBinningBase(streamWriter, siteInfoArray, oneTouchDown, inheritanceManager, allowEqualBases, enumJob, false);
        }

        private static SiteInfo[] CompareAdjust_VddBinningBase(StreamWriter streamWriter, SiteInfo[] siteInfoArray, OneTouchDown oneTouchDown, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, EnumJob enumJob, bool isCsharp)
        {
            if (isCsharp)
            {
                #region Compare <Adjust_VddBinning>
                if (oneTouchDown.GetAdjustVddVddBinningCs(out List<AdjustVddBinningRow> adjustVddBinningRows, out List<ProductIdentifierLineRow> productIdentifierRows))
                {
                    var targetAdjustVddBinningRows = new List<AdjustVddBinningRow>();
                    for (int site = 0; site < siteInfoArray.Length; site++)
                    {
                        if (!siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                        {
                            continue;
                        }

                        var corePowers = BinCutData.ModeVsCorePowerName.Select(x => x.Value + "_" + x.Key).Distinct().ToList();
                        targetAdjustVddBinningRows.AddRange(SiteInfoHelpers.GetAdjustVddBinningRows(site, corePowers, siteInfoArray[site].AllPowers));
                    }
                    AdjustByAllowEqualCsharp(streamWriter, ref siteInfoArray, inheritanceManager, allowEqualBases);
                    var adjustVddBinComparer = new AdjustVddBinningComparer();
                    AdjustVddBinningComparer.CompareAdjustVddBin(streamWriter, siteInfoArray, adjustVddBinningRows, enumJob, BinCutData.BinningTables.IsEqnBin, out List<int> binXList);
                    if (BinCutData.BinningTables.IsEqnBin)
                    {
                        AdjustVddBinningComparer.CompareProductIdentifier(streamWriter, productIdentifierRows, binXList);
                    }
                }
                #endregion
            }
            else
            {
                #region Compare <Adjust_VddBinning>
                if (oneTouchDown.GetAdjustVddBinning(out List<AdjustVddBinningRow> adjustVddBinningRows, out List<ProductIdentifierLineRow> productIdentifierRows))
                {
                    var targetAdjustVddBinningRows = new List<AdjustVddBinningRow>();
                    for (int site = 0; site < siteInfoArray.Length; site++)
                    {
                        if (!siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                        {
                            continue;
                        }

                        var corePowers = BinCutData.ModeVsCorePowerName.Select(x => x.Value + "_" + x.Key).Distinct().ToList();
                        targetAdjustVddBinningRows.AddRange(SiteInfoHelpers.GetAdjustVddBinningRows(site, corePowers, siteInfoArray[site].AllPowers));
                    }
                    AdjustByAllowEqual(streamWriter, ref siteInfoArray, inheritanceManager, allowEqualBases);
                    var adjustVddBinComparer = new AdjustVddBinningComparer();
                    AdjustVddBinningComparer.CompareAdjustVddBin(streamWriter, siteInfoArray, adjustVddBinningRows, enumJob, BinCutData.BinningTables.IsEqnBin, out List<int> binXList);
                    if (BinCutData.BinningTables.IsEqnBin)
                    {
                        AdjustVddBinningComparer.CompareProductIdentifier(streamWriter, productIdentifierRows, binXList);
                    }
                }
                #endregion
            }
            return siteInfoArray;
        }
    }
}
