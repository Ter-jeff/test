using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;

namespace BinCutScriptLib.SetFunction
{
    internal class JudgePassFailMain
    {
        public static bool IsStepPf(ref SiteInfo[] siteInfoArray, ref OneStep oneStep, int powerIdx, int searchStep, bool isinitSkip)
        {
            List<PatternRow> oneStepPat = oneStep.OneStepPatternRows;
            List<BinCutLineBase> oneStepTmps = oneStep.OneStepTmps;
            List<BinCutLineBase> oneStepIlbs = oneStep.OneStepIlbs;
            List<BinCutLineBase> oneStepElbs = oneStep.OneStepElbs;
            List<BinCutLineBase> oneStepLimits = oneStep.OneStepLimits;
            List<BinCutLineBase> oneStepMeasures = oneStep.OneStepMeasures;
            bool isStepFail = oneStepPat.Count > 0 && IsPatPf(ref siteInfoArray, oneStepPat, powerIdx, searchStep, isinitSkip);
            bool isStepFail1 = oneStepTmps.Count > 0 && CheckStepPf(ref siteInfoArray, oneStepTmps, 3, powerIdx, searchStep, isinitSkip);
            bool isStepFail2 = oneStepIlbs.Count > 0 && CheckStepPf(ref siteInfoArray, oneStepIlbs, 3, powerIdx, searchStep, isinitSkip);
            bool isStepFail3 = oneStepElbs.Count > 0 && CheckStepPf(ref siteInfoArray, oneStepElbs, 3, powerIdx, searchStep, isinitSkip);
            bool isStepFail4 = oneStepLimits.Count > 0 && CheckStepPf(ref siteInfoArray, oneStepLimits, 3, powerIdx, searchStep, isinitSkip);
            bool isStepFail5 = oneStepMeasures.Count > 0 && CheckStepPf(ref siteInfoArray, oneStepMeasures, 2, powerIdx, searchStep, isinitSkip);
            return isStepFail || isStepFail1 || isStepFail2 || isStepFail3 || isStepFail4 || isStepFail5;
        }

        public static bool IsPatPfWithoutBinout(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<PatternRow> patternRows, BvName bvName, bool isNoBinOut, string curInstanceName, List<int>? mfstpNobinoutSite = null)
        {
            bool isStepFail = false;
            //ex. 1590259  0     SocTd_MS001_DM_TTR_GRP1_100MHz_BV PP_CMNB1_S_PL00_SC_CCC0_TDF_COM_AUT_ALLFV_DM_1_B1_1603262249.PAT (F) N/A               N/A
            //ex. 1590259  1     SocTd_MS001_DM_TTR_GRP1_100MHz_BV PP_CMNB1_S_PL00_SC_CCC0_TDF_COM_AUT_ALLFV_DM_1_B1_1603262249.PAT (F) N/A               N/A
            foreach (PatternRow oneStepPatternRow in patternRows)
            {
                string[] spt = oneStepPatternRow.PatternLine.Line.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);
                int site = oneStepPatternRow.Site;
                bool isFail = oneStepPatternRow.IsFail;
                if (isFail)
                {
                    isStepFail = true;
                    if (!isNoBinOut && !(mfstpNobinoutSite != null && mfstpNobinoutSite.Contains(site)))
                    {
                        siteInfoArray[site].AllPowers[bvName.Index].SearchStatus = EnumSearchStatus.BinOut;
                        siteInfoArray[site].IsShutDown = true;
                    }
                    string[] patLineSpilt = oneStepPatternRow.PatternLine.Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                    bool isSameInstance = patLineSpilt[2] == curInstanceName;
                    if (BinCutData.GoodBins.Contains(siteInfoArray[site].SortBin.ToString()) && !isNoBinOut && !(mfstpNobinoutSite != null && mfstpNobinoutSite.Contains(site)) && isSameInstance)
                    {
                        if (spt.Length >= 6 && !double.TryParse(spt[5], out _))
                        {
                            if (!(BinCutConfig.DebugBinCutCofStored || BinCutConfig.IsDoAll))
                            {
                                BinCutPrint.PrintBinoutError(streamWriter, oneStepPatternRow.PatternLine);
                            }
                        }
                    }
                }
            }
            return isStepFail;
        }

        private static bool IsPatPf(ref SiteInfo[] siteInfoArray, List<PatternRow> patternRows, int powerIdx, int searchStep, bool isinitSkip)
        {
            bool isStepFail = false;
            //ex. 1590259  0     SocTd_MS001_DM_TTR_GRP1_100MHz_BV PP_CMNB1_S_PL00_SC_CCC0_TDF_COM_AUT_ALLFV_DM_1_B1_1603262249.PAT (F) N/A               N/A
            //ex. 1590259  1     SocTd_MS001_DM_TTR_GRP1_100MHz_BV PP_CMNB1_S_PL00_SC_CCC0_TDF_COM_AUT_ALLFV_DM_1_B1_1603262249.PAT (F) N/A               N/A
            var testedSite = new List<int>();
            var failSite = new List<int>();
            foreach (PatternRow patternRow in patternRows)
            {
                int site = patternRow.Site;
                string patternName = patternRow.PatternName;
                int step = siteInfoArray[site].AllPowers[powerIdx].Step;
                if (step == -1 || step >= siteInfoArray[site].AllPowers[powerIdx].PossibleSteps.Count)
                {
                }
                else
                {
                    PowerStep powerInfo = siteInfoArray[site].AllPowers[powerIdx].PossibleSteps[step];
                    bool isFail = patternRow.IsFail;
                    if (isFail)
                    {
                        isStepFail = true;
                    }

                    SetPatResultList(ref siteInfoArray, ref testedSite, ref failSite, searchStep, isFail, site, patternName, powerInfo);
                }
            }
            SetPatternFail(siteInfoArray, powerIdx, isinitSkip, testedSite, failSite);
            return isStepFail;
        }

        private static bool CheckStepPf(ref SiteInfo[] siteInfoArray, List<BinCutLineBase> binCutLineBases, int patternIdx, int powerIdx, int searchStep, bool isinitSkip)
        {
            // IsIlbPf
            //Site:0 Instance: ILB_MD001_cacs_ck_BV    Lo_limit: 0    Value: 0  Hi_limit: 0
            //Site:1 Instance: ILB_MD001_cacs_ck_BV    Lo_limit: 0    Value: 0  Hi_limit: 0

            // IsElbPf
            //1508662  0     MD001_PP_TURA0_S_FULP_AN_AM00_BST_JTG_ELB_MD001_SI_CACKHOF_T109_CK_BV                                                                         -1       32.0000        0.0000         (F) 127.0000       0.0000         0
            //1508662  1     MD001_PP_TURA0_S_FULP_AN_AM00_BST_JTG_ELB_MD001_SI_CACKHOF_T109_CK_BV                                                                         -1       32.0000        0.0000         (F) 127.0000       0.0000         0

            // IsLimitPf
            //14202755 4     HAC_CalcC_L_ILBHS_D2M_X_X_x_x_0      -1       50.0000 %      0.0000 %             (F) 95.0000 %      0.0000         0       
            //14202755 5     HAC_CalcC_L_ILBHS_D2M_X_X_x_x_0      -1       50.0000 %      85.8000 %                95.0000 %      0.0000         0       

            // IsMeasurePf
            //50449003 3     HAC_MeasC_L_INTSMD5CA_AMPLP5P_X_eye-rddata_DDR2-LP5P-CA_x_6    JTAG_TDO    9.ab414    128            244                  256            0              0       
            //50449004 3     HAC_MeasC_L_INTSMD5CA_AMPLP5P_X_eye-rddata_DDR3-LP5P-CA_x_7    JTAG_TDO    9.ab414    128            0                (F) 256            0              0       

            bool isStepFail = false;
            var testedSite = new List<int>();
            var failSite = new List<int>();

            foreach (BinCutLineBase lineObj in binCutLineBases)
            {
                string line = lineObj.Line;
                string[] spt = line.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);
                if (!int.TryParse(spt[1], out int site))
                {
                    continue;
                }

                string patternName = spt[patternIdx];
                int step = siteInfoArray[site].AllPowers[powerIdx].Step;
                PowerStep powerInfo = siteInfoArray[site].AllPowers[powerIdx].PossibleSteps[step];

                bool isFail = line.Contains("(F)") || line.Contains("(A)");
                if (isFail)
                {
                    isStepFail = true;
                }

                SetPatResultList(ref siteInfoArray, ref testedSite, ref failSite, searchStep, isFail, site, patternName, powerInfo);
            }

            SetPatternFail(siteInfoArray, powerIdx, isinitSkip, testedSite, failSite);
            return isStepFail;
        }

        private static void SetPatternFail(SiteInfo[] siteInfoArray, int powerIdx, bool isinitSkip, List<int> testedSite, List<int> failSite)
        {
            if (!isinitSkip)
            {
                testedSite = [.. testedSite.Distinct()];
                failSite = [.. failSite.Distinct()];
                var passSite = testedSite.Except(failSite).ToList();
                foreach (int site in failSite)
                {
                    int step = siteInfoArray[site].AllPowers[powerIdx].Step;
                    siteInfoArray[site].AllPowers[powerIdx].PossibleSteps[step].IsPatternFail = true;
                    siteInfoArray[site].AllPowers[powerIdx].PossibleSteps[step].IsLastPatternPassForCof = false;
                }

                foreach (int site in passSite)
                {
                    int step = siteInfoArray[site].AllPowers[powerIdx].Step;
                    siteInfoArray[site].AllPowers[powerIdx].PossibleSteps[step].IsLastPatternPassForCof = true;
                }
            }
        }

        private static void SetPatResultList(ref SiteInfo[] siteInfoArray, ref List<int> testedSite, ref List<int> failSite, int searchStep, bool isFail, int site, string patternName, PowerStep powerStep)
        {
            var patternInfo = new PatternInfo
            {
                PatternName = patternName,
                Bin = powerStep.Bin,
                EqName = powerStep.EqName,
                IsFail = isFail,
                StepNum = searchStep
            };
            siteInfoArray[site].PatResultList.Add(patternInfo);

            if (isFail)
            {
                testedSite.Add(site);
                failSite.Add(site);
                if (siteInfoArray[site].PatternList.Find(x => x.PatternName == patternName) != null)
                {
                    siteInfoArray[site].PatternList.FindLast(x => x.PatternName == patternName)!.IsFail = true;
                }
            }
            else
            {
                testedSite.Add(site);
            }
        }
    }
}
