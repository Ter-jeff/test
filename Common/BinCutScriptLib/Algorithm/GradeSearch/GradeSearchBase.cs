using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;

namespace BinCutScriptLib.Algorithm.GradeSearch
{
    internal abstract class GradeSearchBase(EnumJob enumJob, StreamWriter streamWriter, SearchLine searchLine)
    {
        protected const string RailConst = BinCutLineBase.RailConst;
        protected EnumJob Job = enumJob;
        protected readonly StreamWriter Sw = streamWriter;
        protected readonly SearchLine SearchLine = searchLine;
        protected int StartLine;
        protected int EndLine;

        public abstract bool GetInstances(ref OneTouchDown oneTouchDown, out OneGradeSearch oneGradeSearch, ref string tempName, out bool isSearch);

        protected void PostPorcess(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, int oneTouchIndex, ref string tempName, ref bool isSelSram, bool isSearch)
        {
            oneGradeSearch.InstanceBinCut = GetPatternInfo(Sw, oneGradeSearch, ref tempName, StartLine, EndLine, isSelSram, isSearch);

            RemoveLines(oneTouchDown, oneTouchIndex);

            CheckRunPatternError(oneGradeSearch);
        }

        protected static void RemoveLines(OneTouchDown oneTouchDown, int oneTouchIndex)
        {
            //delete original data to speed up parse
            oneTouchDown.Lines.RemoveRange(0, oneTouchIndex < oneTouchDown.Lines.Count ? oneTouchIndex : oneTouchDown.Lines.Count);
        }

        protected void CheckRunPatternError(OneGradeSearch oneGradeSearch)
        {
            foreach (OneStep oneStepData in oneGradeSearch.Steps)
            {
                if (oneStepData.OneRunPattOnly.Count == 0)
                {
                    return;
                }

                int lineNo = oneStepData.OneRunPattOnly.First().LineNo;

                if (oneStepData.OneStepPatternRows.Count == 0 ||
                    oneStepData.OneStepPatternRows.All(x => x.PatternLine.LineNo < lineNo))
                {
                    string msg = $"Fail line:{lineNo}";
                    Sw.WriteLine(msg);
                    msg = "Please check if there any pattern at Fail line !!!";
                    Sw.WriteLine(msg);
                    Sw.WriteLine("");
                }
            }
        }

        protected InstanceBinCut GetPatternInfo(StreamWriter streamWriter, OneGradeSearch oneGradeSearch, ref string tempName, int startLine, int endLine, bool isSelSram, bool isSearch)
        {
            string curInstanceName = oneGradeSearch.InstanceLine.Line.Trim().Replace("<", "").Replace(">", "");
            InstanceRow? instanceRow = null;
            InstanceBinCut instanceBinCut;
            if (BinCutData.TestInstanceSheet!.Rows.Exists(x => x.TestName.EqualsIgnoreCase(curInstanceName)))
            {
                instanceRow = BinCutData.TestInstanceSheet.Rows.Find(x => x.TestName.EqualsIgnoreCase(curInstanceName));
            }

            if (curInstanceName.Length == 0 && oneGradeSearch.SearchResultWoTestLine.Count == 0) //Miss GetMergePat
            {
                string msg = $"Line <{startLine} ~ {endLine}> Missing Instance Name {curInstanceName}";
                streamWriter.WriteLine(msg);
                instanceBinCut = new InstanceBinCut(Job, oneGradeSearch, instanceRow!, isSelSram);
            }
            else
            {
                instanceBinCut = new InstanceBinCut(Job, oneGradeSearch, instanceRow!, isSelSram);
            }

            instanceBinCut.IsSearch = isSearch;
            tempName = curInstanceName;
            return instanceBinCut;
        }

        protected bool GetPatternLine(OneTouchDown oneTouchDown, ref OneStep oneStep, int oneTouchIndex)
        {
            if (oneTouchDown.Lines[oneTouchIndex].IsPatternLine())
            {
                PatternLine patternLine = oneTouchDown.Lines[oneTouchIndex].NewPatternLine();
                oneStep.OneStepPatternRows.Add(patternLine.GetDatalogPatternRow());
                EndLine = oneTouchDown.Lines[oneTouchIndex].LineNo;
                return true;
            }
            return false;
        }

        protected static bool GetRunPattOnly(OneTouchDown oneTouchDown, ref OneStep oneStep, int oneTouchIndex)
        {
            //run_patt_only : SocTd_MS001_DSSC_pp_ceba0_s_pl00_sc_ccc0_tdf_com_aut_allfv_DM, => To check if there are pattern after run pattern
            if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWithIgnoreCase("run_patt_only :"))
            {
                oneStep.OneRunPattOnly.Add(oneTouchDown.Lines[oneTouchIndex]);
                return true;
            }
            return false;
        }

        protected static bool GetDsscSelSram(OneTouchDown oneTouchDown, ref OneStep oneStep, int oneTouchIndex)
        {
            //DSSC_SELSRAM_BV_Str,0,1111101
            if (oneTouchDown.Lines[oneTouchIndex].IsSelsram())
            {
                oneStep.OneStepDssc.Add(oneTouchDown.Lines[oneTouchIndex]);
                return true;
            }
            return false;
        }

        protected static bool GetHarvestSourceCode(OneTouchDown oneTouchDown, ref OneStep oneStep, int oneTouchIndex)
        {
            //Site:0,F_GFX_CORE0=False, Src Bits = 34, HarvestSourceCode [ First(L) ==> Last(R) ] :0010010011000011001001001100100100
            if (oneTouchDown.Lines[oneTouchIndex].IsHarvestSourceCode())
            {
                HarvestSourceCodetLine line = oneTouchDown.Lines[oneTouchIndex].NewHarvestSourceCodetLine();
                oneStep.OneHarvestSourceCodetRows.Add(line.GetHarvestSourceCodetRow());
                return true;
            }
            return false;
        }

        protected static bool GetTmpsLine(OneTouchDown oneTouchDown, ref OneStep oneStep, int oneTouchIndex)
        {
            //Site: 2 TMPS UnTrim Code  Data_4  lowlimit = 2804   2883  HiLimit = 2900
            //Site: 3 TMPS UnTrim Code  Data_4  lowlimit = 2804   2903 (F) Hilimit = 2900
            if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("TMPS UNTRIM CODE", StringComparison.OrdinalIgnoreCase) &&
                     oneTouchDown.Lines[oneTouchIndex].Line.Contains("(F)"))
            {
                oneStep.OneStepTmps.Add(oneTouchDown.Lines[oneTouchIndex]);
                return true;
            }
            return false;
        }

        protected static bool GetIlbLine(OneTouchDown oneTouchDown, ref OneStep oneStep, int oneTouchIndex)
        {
            //2017/01/16: search ILB item key word, add for skyE
            //Site:0 Instance: ILB_MD001_cacs_ck_BV    Lo_limit: 0    Value: 0  Hi_limit: 0
            //Site:1 Instance: ILB_MD001_cacs_ck_BV    Lo_limit: 0    Value: 0  Hi_limit: 0
            if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWithIgnoreCase("INSTANCE:"))
            {
                oneStep.OneStepIlbs.Add(oneTouchDown.Lines[oneTouchIndex]);
                return true;
            }
            ////error DSP:1044 : Internal Error: An unknown error occurred while fetching a DSP result.  Line = 1166, hr = 0x8000000a
            ////error DSP:1044 : Internal Error: An unknown error occurred while fetching a DSP result.  Line = 1166, hr = 0x8000000a
            //if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWith("error", StringComparison.OrdinalIgnoreCase))
            //{
            //    oneStep.OneStepIlbs.Add(oneTouchDown.Lines[oneTouchIndex]);
            //    return true;
            //}
            return false;
        }

        protected static bool GetInitialVoltageLine(OneTouchDown oneTouchDown, ref OneStep oneStep, int oneTouchIndex)
        {
            if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWithIgnoreCase("Initial_Voltage"))
            {
                oneStep.OneStepInitLines.Add(oneTouchDown.Lines[oneTouchIndex]);
                return true;
            }
            return false;
        }

        protected static bool GetBvLine(OneTouchDown oneTouchDown, ref OneStep oneStep, int oneTouchIndex)
        {
            if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWithIgnoreCase("BV_"))
            {
                if (!oneTouchDown.Lines[oneTouchIndex].Line.Contains("EQN ="))
                {
                    oneStep.OneStepBvLineInfo.Add(oneTouchDown.Lines[oneTouchIndex].NewBvLine().GetBvLineInfo());
                    return true;
                }
            }
            return false;
        }

        protected static bool GetCurrentPassBinCutNum(OneTouchDown oneTouchDown, int oneTouchIndex)
        {
            //print: Site(0), ECIDBlankChk_Var = False, CurrentPassBinCutNum = 1, Fuse CFG & UDR = 0
            if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWithIgnoreCase("print:") &&
                oneTouchDown.Lines[oneTouchIndex].Line.Contains("CurrentPassBinCutNum"))
            {
                string[] arr = oneTouchDown.Lines[oneTouchIndex].Line.Split(',');
                if (Reg.RegexSite3.IsMatch(arr.First()))
                {
                    string[] spt = arr.First().Split([' ', '(', ')'], StringSplitOptions.RemoveEmptyEntries);
                    if (int.TryParse(spt[2], out int site))
                    {
                        oneTouchDown.Bins[site] = (int)double.Parse(spt[3].Split('=').Last());
                        //allDice[site].Bin = oneTouchDown.Bins[site];
                        return true;
                    }
                }
            }
            return false;
        }

        protected static bool GetOffSetLine(OneTouchDown oneTouchDown, ref OneStep oneStep, int oneTouchIndex)
        {
            //Site: 0, VDD_SOC_OFFSET Value: 3, VDD_SOC = 528
            if ((oneTouchDown.Lines[oneTouchIndex].Line.Contains("OFFSET VALUE", StringComparison.OrdinalIgnoreCase) ||
                oneTouchDown.Lines[oneTouchIndex].Line.Contains("DYNAMIC OFFSET", StringComparison.OrdinalIgnoreCase)) && oneTouchDown.Lines[oneTouchIndex].Line.Contains(',', StringComparison.OrdinalIgnoreCase))
            {
                if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("MFSTP", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                oneStep.OneStepOffset.Add(oneTouchDown.Lines[oneTouchIndex].NewOffsetLine());
                return true;
            }
            return false;
        }

        protected static bool IsCallInstanceLine(OneTouchDown oneTouchDown, int oneTouchIndex)
        {
            if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("to call instance:"))
            {
                return true;
            }

            return false;
        }

        protected static bool IsCallInstanceLineCs(OneTouchDown oneTouchDown, int oneTouchIndex)
        {
            if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("[INFO]  Creating overlay"))
            {
                return true;
            }

            return false;
        }

        protected bool HvSearch(ref OneTouchDown oneTouchDown, ref OneGradeSearch oneGradeSearch, ref OneStep oneStep, ref int oneTouchIndex, ref bool isSelSram)
        {
            StartLine = oneTouchDown.Lines[oneTouchIndex].LineNo;
            bool debugLineStart = false;
            int debugLineIndex = 0;
            bool vmainPins = false;
            bool valtPins = false;
            (bool flowControl, bool value) = HandleBeforeInstanceName(oneTouchDown, oneGradeSearch, ref oneStep, ref oneTouchIndex, ref isSelSram);
            if (!flowControl)
            {
                return value;
            }

            HandleAfterInstanceName(oneTouchDown, oneGradeSearch, ref oneStep, ref oneTouchIndex, ref isSelSram, ref debugLineStart, ref debugLineIndex, ref vmainPins, ref valtPins);
            return true;
        }

        private void HandleAfterInstanceName(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, ref OneStep oneStep, ref int oneTouchIndex, ref bool isSelSram, ref bool debugLineStart, ref int debugLineIndex, ref bool vmainPins, ref bool valtPins)
        {
            //Get bv parts after instanceName
            string lastType = "";
            for (; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                bool isPattern = oneTouchDown.Lines[oneTouchIndex].IsPatternLine();
                bool isEnd = isPattern;
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

                CheckIfEnd(oneTouchDown, oneGradeSearch, ref oneStep, oneTouchIndex, lastType, ref isEnd);

                //Site:0,F_GFX_CORE0=False, Src Bits = 34, HarvestSourceCode [ First(L) ==> Last(R) ] :0010010011000011001001001100100100
                if (GetHarvestSourceCode(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }

                if (isPattern)
                {
                    HandlePattern(oneTouchDown, oneGradeSearch, oneStep, oneTouchIndex);
                    lastType = "Pattern";
                    continue;
                }
                if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWithIgnoreCase("BV_"))
                {
                    if (oneGradeSearch.IsCallInstance && oneStep.OneStepBvLineInfo.Count != 0 && lastType != "BV")
                    {
                        oneGradeSearch.Steps.Add(oneStep.Copy());
                        oneStep = new OneStep();
                    }
                }
                if (GetBvLine(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    lastType = "BV";
                    continue;
                }
                lastType = "NotPattern";

                if (GetDsscSelSram(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    isSelSram = true;
                    continue;
                }

                if (SearchLine.IsStartPoint(oneTouchDown.Lines[oneTouchIndex]) ||
                    SearchLine.IsStopPoint(oneTouchDown.Lines[oneTouchIndex]) ||
                    oneTouchDown.Lines[oneTouchIndex].IsAlgLine() ||
                    oneTouchDown.Lines[oneTouchIndex].IsEqnNStart() ||
                    oneTouchDown.Lines[oneTouchIndex].IsVbinResultLine())
                {
                    if (oneStep.OneStepBvLineInfo.Count != 0)
                    {
                        oneGradeSearch.Steps.Add(oneStep.Copy());
                    }

                    break;
                }

                if (GetCurrentPassBinCutNum(oneTouchDown, oneTouchIndex))
                {
                    continue;
                }
            }
        }

        private static void CheckIfEnd(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, ref OneStep oneStep, int oneTouchIndex, string lastType, ref bool isEnd)
        {
            if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWith("*************************************************") ||
                                oneTouchDown.Lines[oneTouchIndex].Line.StartsWith("*print:"))
            {
                isEnd = true;
            }

            if (!isEnd && lastType == "Pattern" && !oneGradeSearch.IsCallInstance)
            {
                oneGradeSearch.Steps.Add(oneStep.Copy());
                oneStep = new OneStep();
            }
        }

        private void HandlePattern(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, OneStep oneStep, int oneTouchIndex)
        {
            PatternLine patternLine = oneTouchDown.Lines[oneTouchIndex].NewPatternLine();
            PatternRow patternRow = patternLine.GetDatalogPatternRow();
            oneStep.OneStepPatternRows.Add(patternRow);
            if (oneGradeSearch.FailSite.Count != 0 && oneGradeSearch.FailSite.ContainsKey(patternRow.Site))
            {
                oneGradeSearch.IsSiteMismatch = true;
            }

            if (patternRow.IsFail)
            {
                if (!oneGradeSearch.FailSite.ContainsKey(patternRow.Site))
                {
                    oneGradeSearch.FailSite.Add(patternRow.Site, oneTouchDown.Lines[oneTouchIndex]);
                }
            }
            EndLine = oneTouchDown.Lines[oneTouchIndex].LineNo;
        }

        private static bool HandleLine(OneTouchDown oneTouchDown, OneStep oneStep, int oneTouchIndex)
        {
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

            if (oneTouchDown.Lines[oneTouchIndex].IsHarMandHarvFailLine())
            {
                return false;
            }

            //Multi-FSTP bypass Binout line (Special for T-Lob) 
            if (oneTouchDown.Lines[oneTouchIndex].IsMfstpBypassBinoutLine())
            {
                string siteStr = oneTouchDown.Lines[oneTouchIndex].Line.Split(',').First().Split(':').Last();

                _ = int.TryParse(siteStr, out int siteNum);
                if (!oneStep.OneStepMfstpNoBinOut.Contains(siteNum))
                {
                    oneStep.OneStepMfstpNoBinOut.Add(siteNum);
                }

                return false;
            }
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

        protected static void HandleDebugLineStart(OneTouchDown oneTouchDown, OneStep oneStep, int oneTouchIndex, ref bool debugLineStart, ref int debugLineIndex, ref bool vmainPins, ref bool valtPins)
        {
            if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWith("---Vmain Pin:"))
            {
                if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("VDD_"))
                {
                    vmainPins = true;
                }
            }
            if (oneTouchDown.Lines[oneTouchIndex].Line.StartsWith("---Valt Pin:"))
            {
                if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("VDD_"))
                {
                    valtPins = true;
                }
            }
            if (
                oneTouchDown.Lines[oneTouchIndex].Line == "-------------------- Bincut Voltage Output end --------------------")
            {
                if (!valtPins ^ vmainPins)
                {
                    oneStep.OneStepDebugErrorLine.Add(oneTouchDown.Lines[debugLineIndex]);
                }
                vmainPins = false;
                valtPins = false;
                debugLineStart = false;
                debugLineIndex = 0;
            }
        }

        private (bool flowControl, bool value) HandleBeforeInstanceName(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, ref OneStep oneStep, ref int oneTouchIndex, ref bool isSelSram)
        {
            //Get bv parts before instanceName
            for (; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (SearchLine.IsStopPoint(oneTouchDown.Lines[oneTouchIndex]))
                {
                    if (oneStep.OneStepBvLineInfo.Count != 0)
                    {
                        oneGradeSearch.Steps.Add(oneStep.Copy());
                    }

                    return (flowControl: false, value: true);
                }

                if (oneTouchDown.Lines[oneTouchIndex].IsInstanceLine())
                {
                    oneGradeSearch.InstanceLine = oneTouchDown.Lines[oneTouchIndex];
                    break;
                }
                if (oneTouchDown.Lines[oneTouchIndex].IsPatternLine())
                {
                    break;
                }

                if (GetBvLine(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }

                if (GetDsscSelSram(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    isSelSram = true;
                    continue;
                }

                if (GetHarvestSourceCode(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }

                if (GetCurrentPassBinCutNum(oneTouchDown, oneTouchIndex))
                {
                    continue;
                }

                if (GetRunPattOnly(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }

                if (GetOffSetLine(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }
            }

            return (flowControl: true, value: default);
        }

        protected bool HvSearchCs(ref OneTouchDown oneTouchDown, ref OneGradeSearch oneGradeSearch, ref OneStep oneStep, ref int i, ref bool isSelSram)
        {
            StartLine = oneTouchDown.Lines[i].LineNo;

            string lastType = "";
            bool hasDigSrcStart = false;
            string bvName = "";
            int startIndex = -1;
            bool isFoundPayload = false;
            for (; i < oneTouchDown.Lines.Count; i++)
            {
                if (oneTouchDown.Lines[i].Line.StartsWith(RailConst) && string.IsNullOrEmpty(oneGradeSearch.InstanceLine.Line))
                {
                    Dictionary<string, string> dic = oneTouchDown.Lines[i].Line[RailConst.Length..].GetDict(' ', '=');
                    dic.TryGetValue("Pmode", out string? pmode);
                    pmode = pmode?.Trim('\'');

                    dic.TryGetValue("Rail", out string? rail);
                    rail = rail?.Trim('\'');

                    oneGradeSearch.Pmode = pmode!.Trim('\'');
                    oneGradeSearch.Rail = string.IsNullOrEmpty(rail) ? "" : rail.Trim('\'');
                    BinCutAlgorithmService.GetModeByName(pmode);
                    bvName = string.IsNullOrEmpty(rail) ? $"BV_{pmode}" : $"BV_{rail}_{pmode}";
                    oneGradeSearch.BvMode = bvName;
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
                    BvLineInfo bvLineInfo = oneTouchDown.Lines[i].NewBvLine().GetBvLineInfoCsharp();
                    bvLineInfo.BvName = bvName;
                    oneStep.SafeVoltageCs.Add(bvLineInfo);
                    lastType = "safe voltage";
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
                    PatternLine patternLine = oneTouchDown.Lines[i].NewPatternLine();
                    oneStep.OneStepPatternRows.Add(patternLine.GetDatalogPatternRow());
                    EndLine = oneTouchDown.Lines[i].LineNo;
                    lastType = "Pattern";
                    continue;
                }
                lastType = "NotPattern";
                #endregion

                if (SearchLine.IsStepStartCs(oneTouchDown.Lines[i]) ||
                   SearchLine.IsStepEndCs(oneTouchDown.Lines[i]))
                {
                    if (oneStep.OneStepBvLineInfo.Count != 0 || oneStep.SafeVoltageCs.Count != 0)
                    {
                        oneGradeSearch.Steps.Add(oneStep.Copy());
                    }
                    break;
                }
            }

            return true;
        }

        protected static bool HandleInfo(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, ref OneStep oneStep, int i, ref string lastType, string bvName, int startIndex)
        {
            bool isFoundPayload;
            if (lastType != "payload voltage" && !string.IsNullOrEmpty(oneGradeSearch.InstanceLine.Line))
            {
                oneGradeSearch.Steps.Add(oneStep.Copy());
                oneStep = new OneStep();
            }

            isFoundPayload = false;
            List<BinCutLineBase> lines = oneTouchDown.Lines.GetRange(startIndex + 1, i - startIndex - 1);

            for (int j = 0; j < lines.Count; j++)
            {
                if (lines[j].Line.Contains("Spec voltage"))
                {
                    _ = int.TryParse(Reg.RegexSite1.Match(lines[j].Line).Groups["site"].ToString(), out int _);
                    List<string> arr = [.. lines[j].Line.Split(',')];
                    var processList = arr.Select(item =>
                    {
                        string trimmed = item.Trim();
                        if (trimmed.EndsWith('V'))
                        {
                            trimmed = trimmed[..^1];
                        }
                        return trimmed;
                    }).ToList();

                    string combinedString = string.Join(", ", processList);
                    combinedString = combinedString.Replace("Spec voltages", "Bincut payload voltage");
                    int lineno = lines[j].LineNo;
                    int findLineno = oneTouchDown.Lines.FindIndex(line => line.LineNo == lineno);
                    oneTouchDown.Lines[findLineno].Line = combinedString;

                    BvLineInfo bvLineInfo = oneTouchDown.Lines[findLineno].NewBvLine().GetBvLineInfoCombinePayloadCs();
                    bvLineInfo.IsHardip = true;
                    bvLineInfo.BvName = bvName;
                    oneStep.OneStepBvLineInfo.Add(bvLineInfo);
                    lastType = "payload voltage";
                }
            }

            return isFoundPayload;
        }

        protected static bool HandleSelSram(OneTouchDown oneTouchDown, OneStep oneStep, int i, bool isSelSram)
        {
            string[] arr = [.. oneTouchDown.Lines[i].Line.Split([' ', '(', ')'], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim())];
            int site = oneTouchDown.Lines[i].GetSite();
            if (oneTouchDown.Lines[i].Line.Contains("DSSC"))
            {
                List<BinCutLineBase> lines = oneTouchDown.Lines.GetRange(i, 1);
                var harvestSrc = new DigSrcsCs(lines);
                List<HarvestSourceCodetRow> harvestSourceCodeRows = harvestSrc.GetDsscsCs();
                oneStep.OneHarvestSourceCodetRows.AddRange(harvestSourceCodeRows);
            }
            else if (oneTouchDown.Lines[i].Line.Contains("Selsram"))
            {
                string selSramBit = arr[3];
                var selSram = new BinCutLineBase
                {
                    Line = $"[INFO]  [Site {site}] {selSramBit}",
                    LineNo = oneTouchDown.Lines[i].LineNo
                };
                oneStep.OneStepDssc.Add(selSram);
                isSelSram = true;
            }
            else
            {
                string selSramBit = arr[3];
                var selSram = new BinCutLineBase
                {
                    Line = $"[INFO]  [Site {site}] {selSramBit}",
                    LineNo = oneTouchDown.Lines[i].LineNo
                };
                oneStep.OneStepDssc.Add(selSram);
                isSelSram = true;
            }

            return isSelSram;
        }

        protected static bool HandleSetupDigSrcTestEnd(OneTouchDown oneTouchDown, OneStep oneStep, int i, ref bool isSelSram, ref int startIndex)
        {
            bool hasDigSrcStart = false;
            List<BinCutLineBase> lines = oneTouchDown.Lines.GetRange(startIndex + 1, i - startIndex - 1);
            var digSrc = new DigSrcsCs(lines);
            List<HarvestSourceCodetRow> harvestSourceCodeRows = digSrc.GetDsscsCsForHarv();
            oneStep.OneHarvestSourceCodetRows.AddRange(harvestSourceCodeRows);
            List<BinCutLineBase> selsrams = digSrc.GetSelsramsCs();
            if (selsrams.Count != 0)
            {
                isSelSram = true;
            }

            oneStep.OneStepDssc.AddRange(selsrams);
            startIndex = -1;
            return hasDigSrcStart;
        }

        protected static void HandlePayloadVoltage(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, ref OneStep oneStep, int i, ref string lastType, string bvName)
        {
            if (lastType != "payload voltage" && !string.IsNullOrEmpty(oneGradeSearch.InstanceLine.Line))
            {
                oneGradeSearch.Steps.Add(oneStep.Copy());
                oneStep = new OneStep();
            }

            BvLineInfo bvLineInfo = oneTouchDown.Lines[i].NewBvLine().GetBvLineInfoCsharp();
            bvLineInfo.BvName = bvName;
            oneStep.OneStepBvLineInfo.Add(bvLineInfo);
            lastType = "payload voltage";
        }

        public int GetStartIndexCs(OneTouchDown oneTouchDown, ref OneGradeSearch oneGradeSearch, out bool returnFlag)
        {
            int i;
            returnFlag = false;
            for (i = 0; i < oneTouchDown.Lines.Count; i++)
            {
                if (SearchLine.IsStepEndCs(oneTouchDown.Lines[i]))
                {
                    RemoveLines(oneTouchDown, i);
                    returnFlag = true;
                    return i;
                }

                if (IsCallInstanceLineCs(oneTouchDown, i))
                {
                    oneGradeSearch.IsCallInstance = true;
                }

                if (SearchLine.IsStepStartCs(oneTouchDown.Lines[i]))
                {
                    return i;
                }

                if (GetEnLinesCs(oneTouchDown, oneGradeSearch, ref i))
                {
                    continue;
                }

                if (GetSetVBinResultCs(oneTouchDown, oneGradeSearch, ref i))
                {
                    continue;
                }
            }
            returnFlag = true;
            return i;
        }

        private static bool GetSetVBinResultCs(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, ref int i)
        {
            if (oneTouchDown.Lines[i].Line.Contains("[INFO]  ===== Start of SetVoltageWithoutTest ====="))
            {
                i++;
                for (; i < oneTouchDown.Lines.Count; i++)
                {
                    if (oneTouchDown.Lines[i].Line.Contains("[INFO]  ===== End of SetVoltageWithoutTest ====="))
                    {
                        return true;
                    }

                    oneGradeSearch.VBinResultLines.Add(oneTouchDown.Lines[i].NewVBinResultWoTestLine());
                }
            }

            return false;
        }

        private static bool GetEnLinesCs(OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch, ref int start)
        {
            if (oneTouchDown.Lines[start].Line.StartsWith("****Start of calculation for"))
            {
                string text = oneTouchDown.Lines[start].Line.Trim(' ').Trim('*').Split(' ').Last();
                string mode = BinCutAlgorithmService.GetModeByName(text);
                //****Start of calculation for Vx_SOC_MS002 * ***
                //[INFO][Site 0] Interpolated MS001 bin result 575mV and MS003 bin result 710mV to MS002 635mV
                //[INFO][Site 0] Rounded interpolated voltage 635 up to EQN 4 voltage 640mV
                //< IPL_VDD_SOC_MS002_BV >
                // 41896000         0     VDD_SOC_MS002 Interpolated_EQN                                                                                 VDD_SOC                             9.e246   1              4                    9              0              0
                // 41896001         0     VDD_SOC_MS002 Interpolated_BV                                                                                  VDD_SOC                             9.e246   530.00 mV      640.00 mV            705.00 mV      0.00           0
                // 41896002         0     VDD_SOC_MS002 Interpolated_Product                                                                             VDD_SOC                             9.e246   595.00 mV      705.00 mV            770.00 mV      0.00           0
                //****End of calculation for Vx_SOC_MS002 * ***
                var rows = new List<InterpolationRow>();
                for (int i = start; i < oneTouchDown.Lines.Count; i++)
                {
                    if (oneTouchDown.Lines[i].Line.StartsWith("****End of calculation for"))
                    {
                        //Merge by site
                        IEnumerable<IGrouping<int, InterpolationRow>> groups = rows.GroupBy(x => x.Site);
                        foreach (IGrouping<int, InterpolationRow> group in groups)
                        {
                            var interpolationRow = new InterpolationRow
                            {
                                Site = group.Key,
                                Mode = mode,
                                SkipTest = true
                            };
                            foreach (InterpolationRow item in group)
                            {
                                if (item.Eqidx != -1)
                                {
                                    interpolationRow.Eqidx = item.Eqidx;
                                }

                                if (item.EnLine != null)
                                {
                                    interpolationRow.EnLine = item.EnLine;
                                }

                                if (item.EnValue != -1)
                                {
                                    interpolationRow.EnValue = item.EnValue;
                                }
                            }
                            oneTouchDown.EnRows.Add(interpolationRow);
                            oneGradeSearch.EnRows.Add(interpolationRow);
                        }

                        start = i;
                        return true;
                    }
                    //var match = Reg.regexInterpolation1.Match(oneTouchDown.Lines[i].Line);
                    //if (match.Success)
                    //{
                    //    interpolationRow = new InterpolationRow();
                    //    interpolationRow.EnLine = oneTouchDown.Lines[i].NewEnLine();
                    //    interpolationRow.Mode = match.Groups[5].Value;
                    //    interpolationRow.LowMode = match.Groups[1].Value;
                    //    interpolationRow.HighMode1 = match.Groups[3].Value;
                    //    interpolationRow.HighMode2 = "";
                    //    double.TryParse(match.Groups[6].Value, out interpolationRow.EnValue);
                    //    interpolationRow.Site = oneTouchDown.Lines[i].GetSite();
                    //    interpolationRow.SkipTest = true;
                    //    interpolationRow.SelectedLvcc = 0;
                    //    interpolationRow.Bin4Cand = false;
                    //    continue;
                    //}

                    //var match1 = Reg.regexInterpolation2.Match(oneTouchDown.Lines[i].Line);
                    //if (match1.Success)
                    //{
                    //    double.TryParse(match1.Groups[3].Value, out interpolationRow.EnValue);
                    //    int.TryParse(match1.Groups[2].Value, out interpolationRow.Eqidx);
                    //    if (!string.IsNullOrEmpty(interpolationRow.Mode) &&
                    //        !string.IsNullOrEmpty(interpolationRow.HighMode1) &&
                    //        !string.IsNullOrEmpty(interpolationRow.LowMode))
                    //    {
                    //        oneTouchDown.EnRows.Add(interpolationRow);
                    //        oneGradeSearch.EnRows.Add(interpolationRow);
                    //    }
                    //    continue;
                    //}

                    BinCutLineBase line = oneTouchDown.Lines[i];
                    if (line.Line.StartsWith("[INFO]"))
                    {
                        continue;
                    }

                    if (!line.Line.Contains("Interpolated_EQN ") && !line.Line.Contains("Interpolated_BV ") && !line.Line.Contains("Interpolated_Product "))
                    {
                        continue;
                    }

                    string[] arr = line.Line.Split([" K ", " "], StringSplitOptions.RemoveEmptyEntries);
                    _ = int.TryParse(arr[1], out int site);
                    string type = arr[3];
                    int channelIndex = CommonLib.Datalog.LineBase.GetChannelIndex(arr);
                    int measureIndex = CommonLib.Datalog.LineBase.GetMeasureIndex(channelIndex, arr);
                    _ = double.TryParse(arr[measureIndex], out double value);
                    var row = new InterpolationRow
                    {
                        Site = site
                    };
                    switch (type)
                    {
                        case "Interpolated_EQN":
                            {
                                row.Eqidx = (int)value;
                                row.EnLine = line.NewEnLine();
                                rows.Add(row);
                                break;
                            }
                        case "Interpolated_BV":
                            {
                                row.EnValue = value;
                                row.EnLine = line.NewEnLine();
                                rows.Add(row);
                                break;
                            }
                        case "Interpolated_Product":
                            {
                                break;
                            }
                    }
                }
            }
            return false;
        }
    }
}
