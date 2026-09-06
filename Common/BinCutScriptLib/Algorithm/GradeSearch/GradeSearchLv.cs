using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

namespace BinCutScriptLib.Algorithm.GradeSearch
{
    internal class GradeSearchLv(EnumJob enumJob, StreamWriter streamWriter, SearchLine searchLine) : GradeSearchBase(enumJob, streamWriter, searchLine)
    {
        public override bool GetInstances(ref OneTouchDown oneTouchDown, out OneGradeSearch oneGradeSearch, ref string tempName, out bool isSearch)
        {
            try
            {
                bool isSelSram = false;
                oneGradeSearch = new OneGradeSearch();
                var oneStep = new OneStep();
                isSearch = false;
                if (!SearchLine.GetReturn(oneTouchDown))
                {
                    return false;
                }

                //Search until found BV_
                bool? returnFlag = null;
                int oneTouchIndex = GetStartIndex(oneTouchDown, ref oneGradeSearch, ref oneStep, ref returnFlag, out isSearch);

                if (returnFlag != null)
                {
                    return (bool)returnFlag;
                }

                bool flag = isSearch ?
                    LvSearch(ref oneTouchDown, ref oneGradeSearch, ref oneStep, ref oneTouchIndex, ref isSelSram) :
                    HvSearch(ref oneTouchDown, ref oneGradeSearch, ref oneStep, ref oneTouchIndex, ref isSelSram);

                PostPorcess(oneTouchDown, oneGradeSearch, oneTouchIndex, ref tempName, ref isSelSram, isSearch);
                return flag;
            }

            catch (Exception ex)
            {
                throw new Exception(string.Format(ex.ToString()));
            }
        }

        public bool GetInstancesCs(ref OneTouchDown oneTouchDown, out OneGradeSearch oneGradeSearch, ref string tempName, out bool isSearch)
        {
            try
            {
                bool isSelSram = false;
                oneGradeSearch = new OneGradeSearch();
                var oneStep = new OneStep();
                isSearch = true;

                if (!SearchLine.GetReturn(oneTouchDown))
                {
                    return false;
                }

                int oneTouchIndex = GetStartIndexCs(oneTouchDown, ref oneGradeSearch, out bool returnFlag);

                if (returnFlag)
                {
                    return !returnFlag;
                }

                bool flag = isSearch ?
                    LvSearchCs(ref oneTouchDown, ref oneGradeSearch, ref oneStep, ref oneTouchIndex, ref isSelSram) :
                    HvSearchCs(ref oneTouchDown, ref oneGradeSearch, ref oneStep, ref oneTouchIndex, ref isSelSram);

                PostPorcess(oneTouchDown, oneGradeSearch, oneTouchIndex, ref tempName, ref isSelSram, isSearch);
                return flag;
            }

            catch (Exception ex)
            {
                throw new Exception(string.Format(ex.ToString()));
            }
        }

        private bool LvSearch(ref OneTouchDown oneTouchDown, ref OneGradeSearch oneGradeSearch, ref OneStep oneStep, ref int oneTouchIndex, ref bool isSelSram)
        {
            StartLine = oneTouchDown.Lines[oneTouchIndex].LineNo;

            string lastType = "";
            bool debugLineStart = false;
            int debugLineIndex = 0;
            bool vmainPins = false;
            bool valtPins = false;
            string previousType = "";
            if (oneGradeSearch.SearchResultWoTestLine.Count != 0)
            {
                oneTouchDown.Lines.RemoveRange(0, 1);
                return true;
            }
            GetAllGradeSearch(oneTouchDown, oneGradeSearch, ref oneStep, ref oneTouchIndex, ref isSelSram, ref lastType, ref debugLineStart, ref debugLineIndex, ref vmainPins, ref valtPins, ref previousType);

            oneTouchIndex = GetAllTestLimit(oneTouchDown, oneGradeSearch, oneTouchIndex);

            return true;
        }

        private int GetAllTestLimit(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, int oneTouchIndex)
        {
            //STEP3. Get all test limit and save to printOutLines
            for (; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (IsCp1LvStopPoint(oneTouchDown, oneTouchIndex))
                {
                    break;
                }

                //skip Number   Site  Test Name       Pin
                if (!oneTouchDown.Lines[oneTouchIndex].IsCp1LvStepStopPoint())
                {
                    if (oneTouchDown.Lines[oneTouchIndex].IsSetSearchResultWoTestLine())
                    {
                        oneGradeSearch.SearchResultWoTestLine.Add(oneTouchDown.Lines[oneTouchIndex].NewSearchResultWoTestLine());
                        continue;
                    }
                    string testName = "";
                    string[] arr = oneTouchDown.Lines[oneTouchIndex].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                    if (oneTouchDown.Lines[oneTouchIndex].IsHarvFlagSetTrueLine())
                    {
                        oneGradeSearch.IsHarvAssignedTrue = true;
                        continue;
                    }
                    if (arr.Length > 2)
                    {
                        testName = arr[2];
                    }

                    if (testName == "IDS" ||
                        testName == Job.ToString() || testName == "CP" ||
                        testName == "OFFSET" || testName == "EQN" ||
                        testName == "PASSBIN" || testName == "SELSRAM_DSSC")
                    {
                        oneGradeSearch.JudgeResultLines.Add(oneTouchDown.Lines[oneTouchIndex].NewJudgeResultLine());
                    }
                    else
                    {
                        oneGradeSearch.PatternResultLines.Add(oneTouchDown.Lines[oneTouchIndex].NewPatternResultLine());
                    }
                }
            }

            return oneTouchIndex;
        }

        private void GetAllGradeSearch(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, ref OneStep oneStep, ref int oneTouchIndex, ref bool isSelSram, ref string lastType, ref bool debugLineStart, ref int debugLineIndex, ref bool vmainPins, ref bool valtPins, ref string previousType)
        {
            //Iterate to get all grade search block until binning area, eg. "Number   Site  Test Name....."
            for (; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (BinCutConfig.DebugPrintFlag)
                {
                    if (oneTouchDown.Lines[oneTouchIndex].Line == "--------------------Bincut Voltage Output start--------------------")
                    {
                        debugLineStart = true;
                        debugLineIndex = oneTouchIndex - 1;
                        continue;
                    }
                    if (debugLineStart)
                    {
                        HandleDebugLineStart(oneTouchDown, oneStep, oneTouchIndex, ref debugLineStart, ref debugLineIndex, ref vmainPins, ref valtPins);
                        continue;
                    }
                }

                bool flowControl = HandleLine(oneTouchDown, oneStep, oneTouchIndex);
                if (!flowControl)
                {
                    continue;
                }

                if (oneTouchDown.Lines[oneTouchIndex].IsHarvestBinningFlag())
                {
                    string harvFlag = oneTouchDown.Lines[oneTouchIndex].Line.Split(',').ToList().Find(x => x.Contains("HarvestBinningFlag"))!.Split(':').Last().Trim();
                    if (harvFlag.Length != 0)
                    {
                        oneGradeSearch.HarvestBinningFlag = harvFlag;
                    }
                }
                //Get BV_ or Initial_Voltage_ to last pattern
                bool isPattern = oneTouchDown.Lines[oneTouchIndex].IsPatternLine();
                //Site:0,F_GFX_CORE0=False, Src Bits = 34, HarvestSourceCode [ First(L) ==> Last(R) ] :0010010011000011001001001100100100
                if (GetHarvestSourceCode(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }

                if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("str_harv_true :"))
                {
                    continue;
                }

                oneStep = HandleNotPattern(oneTouchDown, oneGradeSearch, oneStep, oneTouchIndex, lastType, isPattern);

                //1590258  5     SocTd_MS001_DM_TTR_GRP1_100MHz_BV PP_CMNB1_S_IN00_SC_CCC0_TDF_COM_AUT_MS001_DM_1_B1_1603272348.PAT     N/A               N/A                  
                //1590259  0     SocTd_MS001_DM_TTR_GRP1_100MHz_BV PP_CMNB1_S_PL00_SC_CCC0_TDF_COM_AUT_ALLFV_DM_1_B1_1603262249.PAT (F) N/A               N/A
                //1590259  0     SocTd_MS001_DM_TTR_GRP1_100MHz_BV PP_CMNB1_S_PL00_SC_CCC0_TDF_COM_AUT_ALLFV_DM_1_B1_1603262249.PAT (F) 2               61
                if (isPattern)
                {
                    PatternLine patternLine = oneTouchDown.Lines[oneTouchIndex].NewPatternLine();
                    oneStep.OneStepPatternRows.Add(patternLine.GetDatalogPatternRow());
                    EndLine = oneTouchDown.Lines[oneTouchIndex].LineNo;
                    lastType = "Pattern";
                    continue;
                }

                if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWithIgnoreCase("BV_"))
                {
                    HandleBvLine(oneTouchDown, oneGradeSearch, ref oneStep, oneTouchIndex, ref lastType, ref previousType);
                    continue;
                }
                lastType = "NoPattern";
                if (GetInitialVoltageLine(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }

                if (oneTouchDown.Lines[oneTouchIndex].IsInstanceLine())
                {
                    oneGradeSearch.InstanceLine = oneTouchDown.Lines[oneTouchIndex];
                }

                if (IsCp1LvStepStopPoint(oneTouchDown, oneTouchIndex, oneGradeSearch.IsCallInstance))
                {
                    if (oneGradeSearch.IsCallInstance && oneStep.OneStepBvLineInfo.Count != 0)
                    {
                        oneGradeSearch.Steps.Add(oneStep.Copy());
                    }
                    break;
                }

                if (GetOffSetLine(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }

                //DSSC_SELSRAM_BV_Str,0,1111101
                if (GetDsscSelSram(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    isSelSram = true;
                    continue;
                }

                //Site: 5 TMPS Trim Code 5  lowlimit = 20'C  27.70'C  (F) hilimit = 30'C
                //Site: 2 TMPS UnTrim Code  Data_4  lowlimit = 2804   2883  HiLimit = 2900
                //Site: 3 TMPS UnTrim Code  Data_4  lowlimit = 2804   2903 (F) Hilimit = 2900
                if (GetTmpsLine(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }

                ////Site:0 Instance: ILB_MD001_cacs_ck_BV    Lo_limit: 1    Value: 1  Hi_limit: 1
                if (GetIlbLine(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }
            }
        }

        private static void HandleBvLine(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, ref OneStep oneStep, int oneTouchIndex, ref string lastType, ref string previousType)
        {
            if (oneGradeSearch.IsCallInstance && oneStep.OneStepBvLineInfo.Count != 0 && lastType != "BV")
            {
                oneGradeSearch.Steps.Add(oneStep.Copy());
                oneStep = new OneStep();
            }
            if (!oneTouchDown.Lines[oneTouchIndex].Line.Contains("EQN ="))
            {
                BvLineInfo bvLineInfo = oneTouchDown.Lines[oneTouchIndex].NewBvLine().GetBvLineInfo();
                if (!(string.IsNullOrEmpty(previousType) || previousType == bvLineInfo.Type))
                {
                    if (oneStep.OneStepBvLineInfo.Count != 0)
                    {
                        oneGradeSearch.Steps.Add(oneStep.Copy());
                        oneStep = new OneStep();
                    }
                }
                oneStep.OneStepBvLineInfo.Add(bvLineInfo);
                previousType = bvLineInfo.Type;
            }
            else
            {
                BvLineInfo bvEqLine = oneTouchDown.Lines[oneTouchIndex].NewBvLine().GetBvLineInfo();
                BvLineInfo bvLineInfo =
                    oneStep.OneStepBvLineInfo.Find(
                        x => x.BvName == bvEqLine.BvName && x.Site.Equals(bvEqLine.Site))!;
                bvLineInfo.Eqn = bvEqLine.Eqn;
                bvLineInfo.Passbin = bvEqLine.Passbin;
                bvLineInfo.C = bvEqLine.C;
                bvLineInfo.M = bvEqLine.M;
                bvLineInfo.Cpgb = bvEqLine.Cpgb;
                bvLineInfo.Ids = bvEqLine.Ids;
            }

            lastType = "BV";
        }

        private static OneStep HandleNotPattern(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, OneStep oneStep, int oneTouchIndex, string lastType, bool isPattern)
        {
            if (!isPattern)
            {
                if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("(F)") && oneTouchDown.Lines[oneTouchIndex].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries).Length == 13)
                {
                    oneStep.OneStepLimits.Add(oneTouchDown.Lines[oneTouchIndex]);
                }

                if (oneGradeSearch.IsCallInstance)
                {
                    string[] arr = oneTouchDown.Lines[oneTouchIndex].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                    if (arr.Length >= 10 && arr[2].StartsWithIgnoreCase("HAC_"))
                    {
                        oneStep.OneStepMeasures.Add(oneTouchDown.Lines[oneTouchIndex]);
                    }
                }
                if (lastType == "Pattern" && !oneGradeSearch.IsCallInstance)
                {
                    oneGradeSearch.Steps.Add(oneStep.Copy());
                    oneStep = new OneStep();
                }
            }

            return oneStep;
        }

        private static bool HandleLine(OneTouchDown oneTouchDown, OneStep oneStep, int oneTouchIndex)
        {
            if (oneTouchDown.Lines[oneTouchIndex].IsNoBinOutLine())
            {
                oneStep.OneStepNoBinOut = true;
                return false;
            }
            if (oneTouchDown.Lines[oneTouchIndex].IsSyncUpLine())
            {
                oneStep.OneStepSyncUpLine.Add(oneTouchDown.Lines[oneTouchIndex]);
                return false;
            }
            if (oneTouchDown.Lines[oneTouchIndex].IsTotalFailCycleLine())
            {
                return false;
            }

            //Get pin disabled message eg. "Site:0, HAC_HARV_X_MASKED_SocTD_Partition_cc06_pl00_X_H_GROUP_R0042-0" or "Site: 1, disabled PinGroup: C_CORE0".
            if (oneTouchDown.Lines[oneTouchIndex].IsPinGroupDisabledLine())
            {
                oneStep.OneStepDisabledPinGroupLine.Add(oneTouchDown.Lines[oneTouchIndex]);
                return false;
            }
            // Ignore Total disabled PinGroup count message eg. "Total disabled PinGroup count for all site: 2"
            if (oneTouchDown.Lines[oneTouchIndex].IsTotalDisabledPinGroupCountLine())
            {
                return false;
            }

            // Ignore Pin mask feature message eg. "--(Enable Pin mask Feature)--site = 0, PinGroup = H_GROUP_L2943"
            if (oneTouchDown.Lines[oneTouchIndex].IsPinMaskFeatureLine())
            {
                return false;
            }

            if (oneTouchDown.Lines[oneTouchIndex].IsHarvFailLine())
            {
                oneStep.OneStepHarvFailLine.Add(oneTouchDown.Lines[oneTouchIndex]);
                return false;
            }
            if (oneTouchDown.Lines[oneTouchIndex].IsHarMandHarvFailLine())
            {
                return false;
            }

            //Ignore ALARM message, eg. " ALARM DCVS:0022 : ..."
            if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWithIgnoreCase("ALARM"))
            {
                return false;
            }

            //Ignore ERROR DigSrcMaximusPI:0276 : DSSC Source on pin(s) 'JTAG_TDI' has Start label '' that has not been defined. [FlowSheet/Row:Flow_MG001_TD_Multi
            if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWithIgnoreCase("Error"))
            {
                return false;
            }

            return true;
        }

        private bool LvSearchCs(ref OneTouchDown oneTouchDown, ref OneGradeSearch oneGradeSearch, ref OneStep oneStep, ref int i, ref bool isSelSram)
        {
            StartLine = oneTouchDown.Lines[i].LineNo;

            string lastType = "";
            bool hasDigSrcStart = false;
            string bvName = "";
            int startIndex = -1;
            bool isFoundPayload = false;

            for (; i < oneTouchDown.Lines.Count; i++)
            {
                //[INFO]  Starting BinCut test for Pmode='MS001' Rail='VDD_SOC'
                if (oneTouchDown.Lines[i].Line.StartsWith(RailConst) && string.IsNullOrEmpty(oneGradeSearch.InstanceLine.Line))
                {
                    bvName = HandleInfoLine(oneTouchDown, oneGradeSearch, i);
                    continue;
                }

                if (oneTouchDown.Lines[i].Line.Contains("EQN ="))
                {
                    oneStep.EqLineCs.Add(oneTouchDown.Lines[i]);
                    continue;
                }

                #region LogLevel Debug print payload
                if (oneTouchDown.Lines[i].Line.Contains("[INFO]  Creating overlay "))
                {
                    isFoundPayload = true;
                    startIndex = i;
                    continue;
                }

                if (isFoundPayload)
                {
                    if (oneTouchDown.Lines[i].Line.Contains("[INFO]  ==============================================="))
                    {
                        isFoundPayload = HandleInfo(oneTouchDown, oneGradeSearch, ref oneStep, i, ref lastType, bvName, startIndex);
                        continue;
                    }
                }
                #endregion

                if (oneTouchDown.Lines[i].Line.Contains("Bincut payload voltage:"))
                {
                    HandlePayloadVoltage(oneTouchDown, oneGradeSearch, ref oneStep, i, ref lastType, bvName);
                    continue;
                }

                if (oneTouchDown.Lines[i].Line.Contains("Bincut safe voltage:"))
                {
                    lastType = HandleSafeVoltage(oneTouchDown, oneStep, i, bvName);
                    continue;
                }

                if (oneTouchDown.Lines[i].Line.Contains("Setup Dig Src Test Start"))
                {
                    hasDigSrcStart = true;
                    startIndex = i;
                    continue;
                }

                if (hasDigSrcStart)
                {
                    if (oneTouchDown.Lines[i].Line.Contains("Setup Dig Src Test End"))
                    {
                        hasDigSrcStart = HandleSetupDigSrcTestEnd(oneTouchDown, oneStep, i, ref isSelSram, ref startIndex);
                        continue;
                    }
                }
                else
                {
                    // For MP log
                    //[INFO]  [Site 2] 0101(Selsram()) + 1111110(DSSC(X10GR1))
                    //[INFO]  [Site 2] 0101(Selsram())
                    //[INFO]  [Site 2] 1111110(DSSC(X10GR1))
                    //[INFO]  [Site 0] 1(disp_sram_rail),1(gpu_sram_rail),1(dcs_sram_rail),1(soc_sram_rail)
                    //[INFO]  [Site 0] 0101(Selsram(:MS001))
                    if (Reg.RegexSelSram1.IsMatch(oneTouchDown.Lines[i].Line) || Reg.RegexSelSram2.IsMatch(oneTouchDown.Lines[i].Line))
                    {
                        isSelSram = HandleSelSram(oneTouchDown, oneStep, i, isSelSram);
                        continue;
                    }
                }

                if (oneTouchDown.Lines[i].IsInstanceLine())
                {
                    oneGradeSearch.InstanceLine = oneTouchDown.Lines[i];
                    continue;
                }

                #region isPattern
                bool isPattern = oneTouchDown.Lines[i].IsPatternLine();
                //1590258  5     SocTd_MS001_DM_TTR_GRP1_100MHz_BV PP_CMNB1_S_IN00_SC_CCC0_TDF_COM_AUT_MS001_DM_1_B1_1603272348.PAT     N/A               N/A                  
                //1590259  0     SocTd_MS001_DM_TTR_GRP1_100MHz_BV PP_CMNB1_S_PL00_SC_CCC0_TDF_COM_AUT_ALLFV_DM_1_B1_1603262249.PAT (F) N/A               N/A
                //1590259  0     SocTd_MS001_DM_TTR_GRP1_100MHz_BV PP_CMNB1_S_PL00_SC_CCC0_TDF_COM_AUT_ALLFV_DM_1_B1_1603262249.PAT (F) 2               61
                if (isPattern)
                {
                    lastType = HandlePattern(oneTouchDown, oneStep, i);
                    continue;
                }
                lastType = "NotPattern";
                #endregion

                if (IsCp1LvStepStopPointCsharp(oneTouchDown, i))
                {
                    if (oneStep.OneStepBvLineInfo.Count != 0 || oneStep.SafeVoltageCs.Count != 0)
                    {
                        oneGradeSearch.Steps.Add(oneStep.Copy());
                    }
                    break;
                }
            }

            i = HandlePostLines(oneTouchDown, oneGradeSearch, i);

            return true;
        }

        private string HandlePattern(OneTouchDown oneTouchDown, OneStep oneStep, int i)
        {
            string lastType;
            PatternLine patternLine = oneTouchDown.Lines[i].NewPatternLine();
            oneStep.OneStepPatternRows.Add(patternLine.GetDatalogPatternRow());
            EndLine = oneTouchDown.Lines[i].LineNo;
            lastType = "Pattern";
            return lastType;
        }

        private static int HandlePostLines(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, int i)
        {
            for (; i < oneTouchDown.Lines.Count; i++)
            {
                if (IsCp1LvStopPointCsharp(oneTouchDown, i))
                {
                    break;
                }

                //skip Number   Site  Test Name       Pin
                if (!oneTouchDown.Lines[i].IsCp1LvStepStopPoint())
                {
                    if (oneTouchDown.Lines[i].IsSetSearchResultWoTestLine())
                    {

                        oneGradeSearch.SearchResultWoTestLine.Add(oneTouchDown.Lines[i].NewSearchResultWoTestLine());
                        continue;
                    }
                    HandleBinLine(oneTouchDown, oneGradeSearch, i);
                }
            }

            return i;
        }

        private static string HandleSafeVoltage(OneTouchDown oneTouchDown, OneStep oneStep, int i, string bvName)
        {
            string lastType;
            BvLineInfo bvLineInfo = oneTouchDown.Lines[i].NewBvLine().GetBvLineInfoCsharp();
            bvLineInfo.IsHardip = false;
            bvLineInfo.BvName = bvName;
            oneStep.SafeVoltageCs.Add(bvLineInfo);
            lastType = "safe voltage";
            return lastType;
        }

        private static void HandleBinLine(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, int i)
        {
            string testName = "";
            if (!oneTouchDown.Lines[i].Line.Contains('['))
            {
                string[] arr = oneTouchDown.Lines[i].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (arr.Length > 2 && !oneTouchDown.Lines[i].Line.Contains('['))
                {
                    testName = arr[2];
                }

                if (testName == "EQN" ||
                    testName == "BV_Search_Voltage")
                {
                    oneGradeSearch.JudgeResultLines.Add(oneTouchDown.Lines[i].NewJudgeResultLine());
                }
                else
                {
                    oneGradeSearch.PatternResultLines.Add(oneTouchDown.Lines[i].NewPatternResultLine());
                }
            }
        }

        private static string HandleInfoLine(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, int i)
        {
            string bvName;
            Dictionary<string, string> dic = oneTouchDown.Lines[i].Line[RailConst.Length..].GetDict(' ', '=');
            dic.TryGetValue("Pmode", out string? pmode);
            pmode = pmode?.Trim('\'');

            dic.TryGetValue("Rail", out string? rail);
            rail = rail?.Trim('\'');

            if (rail != null && rail.Contains(','))
            {
                rail = rail.Split(',').First();
            }

            oneGradeSearch.Pmode = pmode!.Trim('\'');
            oneGradeSearch.Rail = string.IsNullOrEmpty(rail) ? "" : rail.Trim('\'');
            BinCutAlgorithmService.GetModeByName(pmode);
            bvName = string.IsNullOrEmpty(rail) ? $"BV_{pmode}" : $"BV_{rail}_{pmode}";
            oneGradeSearch.BvMode = bvName;
            return bvName;
        }

        private static bool IsCp1LvStopPoint(OneTouchDown oneTouchDown, int i)
        {
            if (oneTouchDown.Lines[i].Line.StartsWithIgnoreCase("Initial_Voltage"))
            {
                return true;
            }

            if (oneTouchDown.Lines[i].Line.Length == 0)
            {
                return true;
            }

            if (oneTouchDown.Lines[i].Line.Contains("It uses GradeSearch_CallInstance_VT to call instance:"))
            {
                return true;
            }

            if (oneTouchDown.Lines[i].Line.StartsWithIgnoreCase("Flow") &&
                oneTouchDown.Lines[i].Line.Trim().EndsWithIgnoreCase("Stop"))
            {
                return true;
            }

            if (oneTouchDown.Lines[i].Line.Contains("Relay On"))
            {
                return true;
            }

            if (oneTouchDown.Lines[i].Line.StartsWith("***** Start of calculation for Vx"))
            {
                return true;
            }

            if (oneTouchDown.Lines[i].Line.StartsWith("*************"))
            {
                return true;
            }

            if (oneTouchDown.Lines[i].Line.Contains("<PrintOutVddBinning>"))
            {
                return true;
            }

            if (oneTouchDown.Lines[i].Line.Contains("===    PwrBin Harvest_Bin"))
            {
                return true;
            }

            if (oneTouchDown.Lines[i].Line.Contains("PwrBin Sheet :"))
            {
                return true;
            }

            if (oneTouchDown.Lines[i].Line.Contains("Set_VBinResult_without_Test"))
            {
                return true;
            }

            if (oneTouchDown.Lines[i].Line.Contains("Alg=") &&
                oneTouchDown.Lines[i].Line.Split(',')[2].StartsWithIgnoreCase("Alg="))
            {
                return true;
            }

            return false;
        }

        private static bool IsCp1LvStopPointCsharp(OneTouchDown oneTouchDown, int i)
        {
            if (oneTouchDown.Lines[i].Line.StartsWith(RailConst))
            {
                return true;
            }

            if (oneTouchDown.Lines[i].Line.StartsWith("Flow") && (oneTouchDown.Lines[i].Line.Contains("BV_Csharp Stop") || oneTouchDown.Lines[i].Line.Contains("BV Stop")))
            {
                return true;
            }

            return IsCp1LvStopPoint(oneTouchDown, i);
        }

        private static bool IsCp1LvStepStopPoint(OneTouchDown oneTouchDown, int oneTouchIndex, bool isCallInstance)
        {
            if (isCallInstance)
            {
                string[] arr = oneTouchDown.Lines[oneTouchIndex].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (arr.Length > 3 && arr[2].EqualsIgnoreCase("IDS"))
                {
                    return true;
                }
            }
            else
            {
                if (oneTouchDown.Lines[oneTouchIndex].IsCp1LvStepStopPoint())
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsCp1LvStepStopPointCsharp(OneTouchDown oneTouchDown, int i)
        {
            if (oneTouchDown.Lines[i].Line.StartsWith(RailConst))
            {
                return true;
            }

            if (oneTouchDown.Lines[i].IsCp1LvStepStopPoint())
            {
                if (i + 1 < oneTouchDown.Lines.Count)
                {
                    string[] arr = oneTouchDown.Lines[i + 1].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                    if (arr.Length > 3 && arr[2].Trim() == "EQN")
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                return true;
            }

            if (oneTouchDown.Lines[i].IsCp1LvStopPointCsharp())
            {
                return true;
            }
            return false;
        }

        private int GetStartIndex(OneTouchDown oneTouchDown, ref OneGradeSearch oneGradeSearch, ref OneStep oneStep, ref bool? returnFlag, out bool isSearch)
        {
            int oneTouchIndex;
            isSearch = false;
            for (oneTouchIndex = 0; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (oneTouchDown.Lines[oneTouchIndex].IsAlgLine())
                {
                    isSearch = true;
                    continue;
                }
                if (oneTouchDown.Lines[oneTouchIndex].IsFstpHeader())
                {
                    oneGradeSearch.FstpStatus = EnumFstpStatus.Start;
                    continue;
                }
                if (oneTouchDown.Lines[oneTouchIndex].IsFstpFooter())
                {
                    oneGradeSearch.FstpStatus = EnumFstpStatus.Stop;
                    continue;
                }
                if (oneTouchDown.Lines[oneTouchIndex].IsPreVddHeader())
                {
                    oneGradeSearch.PreVddStatus = EnumPreVddStatus.Start;
                    continue;
                }
                if (oneTouchDown.Lines[oneTouchIndex].IsPreVddFooter())
                {
                    oneGradeSearch.PreVddStatus = EnumPreVddStatus.Stop;
                    continue;
                }
                if (oneTouchDown.Lines[oneTouchIndex].IsReJudgeHeader())
                {
                    oneGradeSearch.IsReJudge = true;
                    continue;
                }
                if (IsCallInstanceLine(oneTouchDown, oneTouchIndex))
                {
                    oneGradeSearch.IsCallInstance = true;
                }

                if (SearchLine.IsStopPoint(oneTouchDown.Lines[oneTouchIndex]))
                {
                    RemoveLines(oneTouchDown, oneTouchIndex);
                    returnFlag = false;
                    return oneTouchIndex;
                }

                if (SearchLine.IsStartPoint(oneTouchDown.Lines[oneTouchIndex]))
                {
                    return oneTouchIndex;
                }

                if (GetEnLines(oneTouchDown, oneGradeSearch, ref oneTouchIndex))
                {
                    continue;
                }

                if (GetSetVBinResult(oneTouchDown, oneGradeSearch, ref oneTouchIndex))
                {
                    returnFlag = false;
                    return oneTouchIndex;
                }

                if (GetSetSearchResultWithOutTest(oneTouchDown, oneGradeSearch, ref oneTouchIndex))
                {
                    isSearch = true;
                    return oneTouchIndex;
                }

                if (!isSearch)
                {
                    #region HV
                    if (GetBvLine(oneTouchDown, ref oneStep, oneTouchIndex))
                    {
                        continue;
                    }

                    if (GetOffSetLine(oneTouchDown, ref oneStep, oneTouchIndex))
                    {
                        continue;
                    }

                    #endregion
                }
            }
            returnFlag = false;
            return oneTouchIndex;
        }

        private static bool GetSetVBinResult(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, ref int oneTouchIndex)
        {
            if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("Start of Set_VBinResult_without_Test"))
            {
                oneTouchIndex++;
                for (; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
                {
                    if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("End of Set_VBinResult_without_Test"))
                    {
                        return true;
                    }

                    oneGradeSearch.VBinResultLines.Add(oneTouchDown.Lines[oneTouchIndex].NewVBinResultWoTestLine());
                }
            }
            return false;
        }

        private static bool GetEnLines(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, ref int oneTouchIndex)
        {
            if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("Equation_N_Start"))
            {
                oneTouchIndex++;
                var bin4CandSites = new List<string>();
                //switch to next line, ex: Site:0, VDD_CPU_MC602 ,The En Voltage: 546.875, The lowest Performance Power: MC601 , The highest Performace Power: MC605
                for (; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
                {
                    if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("Equation_N_End"))
                    {
                        return true;
                    }

                    if (oneTouchDown.Lines[oneTouchIndex].Line.EndsWithIgnoreCase("Error!!!"))
                    {
                        return true;
                    }

                    if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("BIN4_CAND"))
                    {
                        bin4CandSites.Add(oneTouchDown.Lines[oneTouchIndex].Line.Split(',').First().Replace("Site:", ""));
                    }
                    InterpolationRow enRow = oneTouchDown.Lines[oneTouchIndex].NewEnLine().GetEnRow();
                    enRow.Bin4Cand = bin4CandSites.Exists(x => x == enRow.Site.ToString(CultureInfo.InvariantCulture));
                    if (!string.IsNullOrEmpty(enRow.Mode) &&
                        !string.IsNullOrEmpty(enRow.HighMode1) &&
                        !string.IsNullOrEmpty(enRow.LowMode))
                    {
                        oneTouchDown.EnRows.Add(enRow);
                        oneGradeSearch.EnRows.Add(enRow);
                    }
                }
            }
            return false;
        }

        private static bool GetSetSearchResultWithOutTest(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, ref int oneTouchIndex)
        {
            if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("_mode is not tested, set VBIN_RESULT of p_mode for Set_E1_Voltage_ForPmode. EQN="))
            {
                oneGradeSearch.SearchResultWoTestLine.Add(oneTouchDown.Lines[oneTouchIndex].NewSearchResultWoTestLine());
                return true;
            }
            return false;
        }
    }
}
