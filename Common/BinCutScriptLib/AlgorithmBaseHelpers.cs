using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Reader;
using BinCutScriptLib.Static;

using CommonLib.Datalog;
using CommonLib.Extension;

using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib
{
    internal static class AlgorithmBaseHelpers
    {
        public static void SetPatternRow(List<LimitLine>? limitLines, ref InstanceData instanceData)
        {
            if (limitLines != null)
            {
                foreach (LimitLine line in limitLines)
                {
                    LimitRow? rowTmp = line?.ToRow();
                    if (rowTmp == null)
                    {
                        continue;
                    }

                    instanceData.PatternResultRows.Add(rowTmp);
                }
            }
        }

        public static List<LimitLine> GetLimitLinesBySite(OneGradeSearch oneGradeSearch, int site)
        {
            var lineTmp = new List<LimitLine>();
            foreach (LimitLine limitLine in oneGradeSearch.PatternResultLines)
            {
                if (!(limitLine.Line.Contains("_EQN") || limitLine.Line.Contains("_PASSBIN") || limitLine.Line.Contains("_CP")))
                {
                    continue;
                }

                LimitRow? limitRow = limitLine.ToRow();
                if (limitRow == null || limitRow.Site != site)
                {
                    continue;
                }

                lineTmp.Add(limitLine);
            }
            return lineTmp;
        }

        internal static int GetDevice(OneTouchDown oneTouchDown, string line, int lineNoCounter)
        {
            #region Device#
            int activeSiteCount;
            oneTouchDown.Lines.Add(new BinCutLineBase { Line = line, LineNo = lineNoCounter });
            string[] spt = line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
            //Case-1:Device#: 1,2
            //Case-2:Device#: 3
            //Case-3:Device#: 29-33
            if (spt[1].Contains(','))
            {
                activeSiteCount = spt[1].Split([','], StringSplitOptions.RemoveEmptyEntries).Length;
            }
            else if (spt[1].Contains('-'))
            {
                string[] spt2 = spt[1].Split(['-'], StringSplitOptions.RemoveEmptyEntries);
                int op1 = int.Parse(spt2[0]);
                int op2 = int.Parse(spt2[1]);
                activeSiteCount = op2 - op1 + 1;
            }
            else
            {
                activeSiteCount = 1;
            }
            #endregion
            return activeSiteCount;
        }

        internal static void SetActiveSite(OneGradeSearch oneGradeSearch, SiteInfo[] siteInfoArray)
        {
            foreach (SiteInfo allDic in siteInfoArray)
            {
                allDic.IsActiveSite = false;
            }

            foreach (OneStep step in oneGradeSearch.Steps)
            {
                foreach (BvLineInfo stepBvLine in step.OneStepBvLineInfo)
                {
                    siteInfoArray[stepBvLine.Site].IsActiveSite = true;
                }
            }
        }

        internal static void GetCoor(StreamReader streamReader, OneTouchDown oneTouchDown, SiteInfo[] siteInfoArray, ref int lineNoCounter)
        {
            //STEP3. get rest coor and rslt
            string? line;
            while ((line = streamReader.ReadLine()) != null)
            {
                lineNoCounter++;
                string[] spt = line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (spt.Length != 3)
                {
                    break;
                }

                if (spt[1] == "N/A" || spt[2] == "N/A")
                {
                    continue;
                }

                int siteTmp = int.Parse(spt[0]);
                //allDice[siteTmp].site = siteTmp;  //site value assign on set sort/sortBin area
                if (spt[1] == "N/A" || spt[2] == "N/A")
                {
                    siteInfoArray[siteTmp].XCoor = -1;
                    siteInfoArray[siteTmp].YCoor = -1;
                }
                else
                {
                    siteInfoArray[siteTmp].XCoor = int.Parse(spt[1]);
                    siteInfoArray[siteTmp].YCoor = int.Parse(spt[2]);
                }

                if (line.Contains('='))  //site info end
                {
                    break;
                }

                oneTouchDown.Lines.Add(new BinCutLineBase { Line = line, LineNo = lineNoCounter });
            }
        }

        internal static void GetSortBin(StreamReader streamReader, OneTouchDown oneTouchDown, SiteInfo[] siteInfoArray, ref int lineNoCounter)
        {
            //STEP2a. Get Sort/SortBin
            // Site    Sort     Bin
            //------------------------------------
            //    0       1       1
            //    1       1       1
            //    2      49       4
            string? line = streamReader.ReadLine() ?? "";
            lineNoCounter++;
            oneTouchDown.Lines.Add(new BinCutLineBase { Line = line, LineNo = lineNoCounter });
            while ((line = streamReader.ReadLine()) != null)
            {
                lineNoCounter++;
                oneTouchDown.Lines.Add(new BinCutLineBase { Line = line, LineNo = lineNoCounter });
                string[] spt = line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (spt.Length < 3 || line.StartsWith("============="))
                {
                    break;
                }

                if (!int.TryParse(spt[0], out _))
                {
                    break;
                }

                int siteTmp = int.Parse(spt[0]);
                siteInfoArray[siteTmp].Site = siteTmp;
                if (spt[1] == "N/A" || spt[2] == "N/A")
                {
                    siteInfoArray[siteTmp].Sort = -1;
                    siteInfoArray[siteTmp].SortBin = -1;
                }
                else
                {
                    siteInfoArray[siteTmp].Sort = int.Parse(spt[1]);
                    siteInfoArray[siteTmp].SortBin = int.Parse(spt[2]);
                    siteInfoArray[siteTmp].SortLineNo = lineNoCounter;
                }
            }

            //read 2 lines
            // Site    X_Coord     Y_Coord
            //------------------------------------
            if (line != " Site    X_Coord     Y_Coord")
            {
                string line1 = streamReader.ReadLine() ?? "";
                lineNoCounter++;
                oneTouchDown.Lines.Add(new BinCutLineBase { Line = line1, LineNo = lineNoCounter });
            }
        }

        public static void SetEqFromBin1Eq1(SiteInfo siteInfo)
        {
            foreach (PowerZone power in siteInfo.AllPowers)
            {
                power.SetBin4CandStep();
            }
        }

        internal static void SetHarvResult(SiteInfo[] siteInfoArray, Dictionary<string, Dictionary<string, bool>> multiPinResult)
        {
            for (int i = 0; i < siteInfoArray.Length; i++)
            {
                if (!multiPinResult.ContainsKey(i.ToString()))
                {
                    continue;
                }

                foreach (KeyValuePair<string, bool> harvResult in multiPinResult[i.ToString()])
                {
                    if (!siteInfoArray[i].HarvesFlags.ContainsKey(harvResult.Key))
                    {
                        siteInfoArray[i].HarvesFlags.Add(harvResult.Key, harvResult.Value ? "T" : "F");
                    }
                }
            }
        }

        public static void SetEqFromBin1Eq1ByMode(SiteInfo siteInfo, string mode)
        {
            foreach (PowerZone power in siteInfo.AllPowers)
            {
                if (power.Mode != mode)
                {
                    continue;
                }

                power.SetBin4CandStep();
                break;
            }
        }

        public static void SetTargetBin(SiteInfo siteInfo, int bin)
        {
            foreach (PowerZone power in siteInfo.AllPowers)
            {
                SetTargetBinStep(bin, power);
            }
        }

        public static void SetTargetBinStep(int bin, PowerZone powerZone)
        {
            int binStep = powerZone.PossibleSteps.FindIndex(x => x.Bin == bin && x.EqName == 1);
            if (binStep != -1)
            {
                powerZone.Bin = bin;
                powerZone.Step = binStep;
                powerZone.StartStep = binStep;
                powerZone.FinalStep = binStep;
                powerZone.StopStep = powerZone.AllSteps.Count - 1;
                powerZone.SearchStatus = EnumSearchStatus.Search;
            }
        }

        internal static void SyncUpStepForAffiliatePin(SiteInfo[] siteInfoArray)
        {
            #region Sync up step for affiliate pin
            foreach (KeyValuePair<string, int> affiliatedPin in BinCutData.BinCutFlowTables.First().AffiliatedPin)
            {
                if (BinCutData.BinCutFlowTables.First().PowerPins.ContainsValue(affiliatedPin.Value))
                {
                    string mainPinName = BinCutData.BinCutFlowTables.First().PowerPins.First(x => x.Value == affiliatedPin.Value).Key;
                    foreach (SiteInfo site in siteInfoArray)
                    {
                        if (!site.IsActiveSite || site.AllPowers.Count == 0)
                        {
                            continue;
                        }

                        List<PowerZone> affiliatedPinPmode = site.AllPowers.FindAll(x => x.Pin.EqualsIgnoreCase(affiliatedPin.Key));
                        List<PowerZone> mainPinPmode = site.AllPowers.FindAll(x => x.Pin.EqualsIgnoreCase(mainPinName));
                        for (int i = 0; i < affiliatedPinPmode.Count; i++)
                        {
                            string a = affiliatedPinPmode[i].Pin;
                            string b = affiliatedPinPmode[i].PinMode;
                            string c = affiliatedPinPmode[i].Mode;
                            int powerIdx = site.AllPowers.FindIndex(x => x.Mode == c);
                            site.AllPowers[powerIdx] = mainPinPmode[i].Copy();
                            site.AllPowers[powerIdx].Pin = a;
                            site.AllPowers[powerIdx].PinMode = b;
                            site.AllPowers[powerIdx].Mode = c;
                        }
                    }
                }
            }
            #endregion
        }

        internal static void SetBinForFBin4XPass(ref SiteInfo[] siteInfoArray, Dictionary<string, Dictionary<string, bool>> multiPinResult)
        {
            foreach (SiteInfo dice in siteInfoArray)
            {
                if (!multiPinResult.ContainsKey(dice.Site.ToString(CultureInfo.InvariantCulture)))
                {
                    continue;
                }

                if (multiPinResult[dice.Site.ToString(CultureInfo.InvariantCulture)].TryGetValue("F_Bin4X_Pass", out bool value) &&
value.Equals(true))
                {
                    foreach (PowerZone power in dice.AllPowers)
                    {
                        int binStep = power.PossibleSteps.FindIndex(x => x.Bin == 2 && x.EqName == 1);
                        if (binStep != -1)
                        {
                            power.Bin = 2;
                            power.Step = binStep;
                            power.StartStep = binStep;
                            power.FinalStep = binStep;
                            power.StopStep = power.GetPosCount() - 1;
                            power.SearchStatus = EnumSearchStatus.Search;
                        }
                    }
                }
            }
        }

        internal static List<BinCutLineBase> GotoVddbinningstart(ref OneTouchDown oneTouchDown)
        {
            //Remove before "[INFO]  ----- BinCut Config start -----"
            int bincutConfigStart = oneTouchDown.Lines.FindIndex(x => x.Line.StartsWithIgnoreCase(BinCutDatalogConfigReader.BcStart));
            if (bincutConfigStart != -1)
            {
                oneTouchDown.Lines.RemoveRange(0, bincutConfigStart - 1);
                int hvEnd = oneTouchDown.Lines.FindIndex(x => x.Line.StartsWith("Flow VddBinning_HVCC Stop"));
                int postEnd = oneTouchDown.Lines.FindIndex(x => x.Line.StartsWith("Flow PostBincut Stop"));
                int lvEnd = oneTouchDown.Lines.FindIndex(x => x.Line.StartsWith("print: Flow_Vddbinning end"));
                if (postEnd != -1)
                {
                    return oneTouchDown.Lines.GetRange(1, postEnd);
                }
                else if (hvEnd != -1)
                {
                    return oneTouchDown.Lines.GetRange(1, hvEnd);
                }
                else if (lvEnd != -1)
                {
                    return oneTouchDown.Lines.GetRange(1, lvEnd);
                }
            }
            return [];
        }

        internal static List<SiteInfo> AddDiceInfo(SiteInfo[] siteInfoArray)
        {
            #region Add diceInfo
            List<SiteInfo> currentDiceInfos = [];
            for (int i = 0; i < siteInfoArray.Length; i++)
            {
                if (siteInfoArray[i].Site != -1)
                {
                    currentDiceInfos.Add(siteInfoArray[i]);
                }
            }
            return currentDiceInfos;
            #endregion
        }

        internal static bool CheckCsharpLogFormat(string dataLogFile)
        {
            if (!File.Exists(dataLogFile))
            {
                return false;
            }
            using var sr = new StreamReader(dataLogFile);
            string? line;
            while ((line = sr.ReadLine()) != null)
            {
                Match csharpMatchFlag = Reg.RegexcsharpMatchFlag.Match(line);
                Match vbtMatchFlag = Reg.RegexvbtMatchFlag.Match(line);
                if (csharpMatchFlag.Success)
                {
                    return true;
                }
                if (vbtMatchFlag.Success)
                {
                    return false;
                }
                if (line.Length >= 1000)
                {
                    return false;
                }

            }
            return false;
        }

        internal static void InitSramPwrStep(SiteInfo[] siteInfoArray, List<PinInfo> pinInfos)
        {
            foreach (SiteInfo diceInfo in siteInfoArray)
            {
                foreach (PinInfo pinInfo in pinInfos.Where(pinInfo => pinInfo.Binned.EqualsIgnoreCase("TRUE") && pinInfo.Pin.Contains("SRAM", StringComparison.OrdinalIgnoreCase)))
                {
                    if (diceInfo.AllPowers.Select(x => x.PinMode).Contains(pinInfo.PinMode))
                    {
                        diceInfo.AllPowers.Find(x => x.PinMode == pinInfo.PinMode)!.FinalStep
                            = diceInfo.AllPowers.Find(x => x.PinMode == pinInfo.PinMode)!.FinalStep == -1 ?
                                0 : diceInfo.AllPowers.Find(x => x.PinMode == pinInfo.PinMode)!.FinalStep;
                    }
                }
            }
        }

        internal static void InitNonCorePwrStep(SiteInfo[] siteInfoArray, List<PinInfo> pinInfos)
        {
            foreach (SiteInfo diceInfo in siteInfoArray)
            {
                foreach (PinInfo pinInfo in pinInfos.Where(pinInfo => !pinInfo.Binned.EqualsIgnoreCase("TRUE")))
                {
                    if (diceInfo.AllPowers.Select(x => x.PinMode).Contains(pinInfo.PinMode))
                    {
                        diceInfo.AllPowers.Find(x => x.PinMode == pinInfo.PinMode)!.FinalStep = 0;
                    }
                }
            }
        }

        internal static void AddCofPattern(bool cofMode, bool isInitSkipMoveStep, OneGradeSearch oneGradeSearch, SiteInfo[] siteInfoArray, BvName bvName)
        {
            if (cofMode && !isInitSkipMoveStep)
            {
                foreach (LimitLine line in oneGradeSearch.PatternResultLines)
                {
                    if (line.Line.Contains("_EQN"))
                    {
                        LimitRow? row = line.ToRow();
                        if (row != null)
                        {
                            int site = row.Site;
                            string name = row.TestName.Replace('_' + row.TestName.Split('_').Last(), "");
                            int step = siteInfoArray[site].AllPowers[bvName.Index].Step;
                            PowerStep powerInfo = siteInfoArray[site].AllPowers[bvName.Index].PossibleSteps[step];
                            var patTemp = new PatternInfo
                            {
                                PatternName = name,
                                Bin = powerInfo.Bin,
                                Cp = powerInfo.Lvcc,
                                EqName = powerInfo.EqName,
                                IsFail = false
                            };
                            siteInfoArray[site].PatternList.Add(patTemp);
                        }
                    }
                }
            }
        }

        internal static void ModifyStepByEfuseValue(SiteInfo[] siteInfoArray, bool cmdMode, Action<string, Color> richTextBoxAppend)
        {
            foreach (SiteInfo allDic in siteInfoArray)
            {
                foreach (EFuseRow eFuse in allDic.EFuseValues)
                {
                    if (eFuse.Value == 0)
                    {
                        continue;
                    }

                    if (allDic.AllPowers.Exists(x => x.Mode.EqualsIgnoreCase(eFuse.Name)))
                    {
                        PowerZone powerZone = allDic.AllPowers.Find(x => x.Mode.EqualsIgnoreCase(eFuse.Name))!;
                        if (powerZone.IdsValue == 0)
                        {
                            continue;
                        }

                        for (int i = 0; i < powerZone.PossibleSteps.Count; i++)
                        {
                            PowerStep step = powerZone.PossibleSteps[i];
                            if (step.BinningProduct == eFuse.Value)
                            {
                                powerZone.StartStep = i;
                                powerZone.Step = i;
                                powerZone.StopStep = powerZone.PossibleSteps.Count - 1;
                                powerZone.FinalStep = i;
                                powerZone.SearchStatus = EnumSearchStatus.Search;
                                break;
                            }
                            if (step.ProductValue > eFuse.Value && powerZone.IdsValue != 0)
                            {
                                string errormsg = $"The efuse value {eFuse.Value} is not on the EQ step !!!";
                                if (cmdMode)
                                {
                                    richTextBoxAppend("The error message is:" + errormsg, Color.Red);
                                }

                                throw new Exception(errormsg);
                            }
                        }
                    }
                }
            }
        }

        internal static void SetAllDiceByLog(SiteInfo[] siteInfoArray, OneTouchDown oneTouchDown, OneGradeSearch oneGradeSearch)
        {
            foreach (KeyValuePair<int, int> item in oneTouchDown.Bins)
            {
                siteInfoArray[item.Key].Bin = item.Value;
            }

            SetActiveSite(oneGradeSearch, siteInfoArray);
        }

        internal static int SuperposRtDice(ref List<SiteInfo> siteInfos)
        {
            int duplicateCnt = 0;
            int totalDice = siteInfos.Count;
            bool[] isDuplicatelist = new bool[totalDice];
            for (int diceIdx = totalDice - 1; diceIdx > -1; diceIdx--)
            {
                int baseX = siteInfos[diceIdx].XCoor;
                int baseY = siteInfos[diceIdx].YCoor;
                for (int nDiceIdx = diceIdx - 1; nDiceIdx > -1; nDiceIdx--)
                {
                    int nextX = siteInfos[nDiceIdx].XCoor;
                    int nextY = siteInfos[nDiceIdx].YCoor;
                    if (baseX == nextX && baseY == nextY && baseX != -1 && baseY != -1)
                    {
                        isDuplicatelist[nDiceIdx] = true;
                        duplicateCnt++;
                        break;
                    }
                }
            }
            for (int diceIdx = totalDice - 1; diceIdx > -1; diceIdx--)
            {
                if (isDuplicatelist[diceIdx])
                {
                    siteInfos.RemoveAt(diceIdx);
                }
            }

            return duplicateCnt;
        }

        internal static void ReadUntilEndOfOneTouch(StreamReader streamReader, OneTouchDown oneTouchDown, ref bool isFoundSortBin, ref bool beforeBVfalg, ref bool beforeoutsideBVfalg, ref bool debugLines, ref int lineNoCounter, ref List<Alarm> alarms)
        {
            string? line;
            //STEP2. Read Until end of one touch
            while ((line = streamReader.ReadLine()) != null)
            {
                lineNoCounter++;
                //Ignore debug lines
                if (line == "================debug print start==================")
                {
                    debugLines = true;
                    continue;
                }
                if (line == "================debug print end  ==================")
                {
                    debugLines = false;
                    continue;
                }
                if (debugLines)
                {
                    continue;
                }

                //Read alarm message (From Device# to Site    Sort     Bin)
                var oneline = new BinCutLineBase { Line = line, LineNo = lineNoCounter };
                if (line.StartsWithIgnoreCase("alarm") || line.StartsWithIgnoreCase("Error ") || line.EndsWithIgnoreCase("warning!!!") || line.EndsWithIgnoreCase("Error!!!") || line.StartsWithIgnoreCase("<warning")
                    )
                {
                    alarms.Add(new Alarm { _alarmMessage = oneline, Type = "", IsBeforeBv = beforeBVfalg });
                }

                if (line.StartsWithIgnoreCase("BV_"))
                {
                    beforeBVfalg = false;
                }

                if (line.Contains("print: VddBinning_Outside_BV start"))
                {
                    beforeoutsideBVfalg = false;
                }

                if (beforeoutsideBVfalg)
                {
                    oneTouchDown.Lines.Add(oneline);
                }

                //touch end string
                if (line.Contains("Site    Sort     Bin"))
                {
                    isFoundSortBin = true;
                    break;
                }
            }
        }

        internal static void SetInstanceInfoLv(string curInstanceName, List<List<PatternRow>> patternRows, ref SiteInfo[] siteInfoArray, int powerIdx, int searchStep, bool cofMode, OneGradeSearch oneGradeSearch)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (!siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                {
                    continue;
                }
                List<LimitLine>? limitLines = cofMode ? GetLimitLinesBySite(oneGradeSearch, site) : null;

                PowerZone pwrRef = siteInfoArray[site].AllPowers[powerIdx];
                int stepDiffTmp = pwrRef.Step - pwrRef.FinalStep;
                if (stepDiffTmp < 0)
                {
                    stepDiffTmp = Math.Abs(stepDiffTmp) - 1;
                }

                if (pwrRef.SearchStatus == EnumSearchStatus.Search)
                {
                    HandleStatusSearchLv(curInstanceName, patternRows, siteInfoArray, powerIdx, searchStep, oneGradeSearch, site, limitLines, pwrRef, stepDiffTmp);
                }
                else
                {
                    HandleStatusOtherLv(curInstanceName, patternRows, siteInfoArray, powerIdx, searchStep, oneGradeSearch, site, limitLines, pwrRef, stepDiffTmp);
                }
            }
        }

        private static List<PatternInfo> GetFailPatternData(SiteInfo[] siteInfoArray, int site)
        {
            List<PatternInfo> failPatternData = [];
            if (BinCutConfig.IsDoAll)
            {
                foreach (PatternInfo pattern in siteInfoArray[site].PatResultList)
                {
                    if (pattern.IsFail)
                    {
                        failPatternData.Add(pattern);
                    }
                }
            }
            return failPatternData;
        }

        private static void HandleStatusOtherLv(string curInstanceName, List<List<PatternRow>> patternRows, SiteInfo[] siteInfoArray, int powerIdx, int searchStep, OneGradeSearch oneGradeSearch, int site, List<LimitLine>? limitLines, PowerZone powerZone, int stepDiffTmp)
        {
            int actualStepsOffset = powerZone.FinalStep == -1 ? 0 : stepDiffTmp;

            // 2. Initialize instance via unified helper
            InstanceData oneInstanceData = CreateBaseInstanceData(curInstanceName, patternRows, powerIdx, searchStep, oneGradeSearch, site, siteInfoArray, limitLines, powerZone, 0.0, 0.0, -1, -1, actualStepsOffset);

            // 3. Add to target collection
            siteInfoArray[site].InstanceList.Add(oneInstanceData);
        }

        private static void HandleStatusSearchLv(string curInstanceName, List<List<PatternRow>> patternRows, SiteInfo[] siteInfoArray, int powerIdx, int searchStep, OneGradeSearch oneGradeSearch, int site, List<LimitLine>? limitLines, PowerZone powerZone, int stepDiffTmp)
        {
            int actualStepsOffset = powerZone.SearchStatus == EnumSearchStatus.BinOut ? 0 : stepDiffTmp;
            int bin = powerZone.GetFinalBin();

            InstanceData oneInstanceData = CreateBaseInstanceData(curInstanceName, patternRows, powerIdx, searchStep, oneGradeSearch, site, siteInfoArray, limitLines, powerZone, powerZone.IdsValue, powerZone.GetFinalLvcc(), powerZone.GetFinalEqName(), bin, actualStepsOffset);

            powerZone.Step = powerZone.FinalStep;
            powerZone.Bin = bin;
            siteInfoArray[site].Bin = bin;
            siteInfoArray[site].InstanceList.Add(oneInstanceData);
        }

        private static InstanceData CreateBaseInstanceData(string curInstanceName, List<List<PatternRow>> patternRows, int powerIdx, int searchStep, OneGradeSearch oneGradeSearch, int site, SiteInfo[] siteInfoArray, List<LimitLine>? limitLines, PowerZone powerZone, double ids, double lvcc, int eqns, int bin, int actualStepsOffset)
        {
            var instanceData = new InstanceData
            {
                InstanceName = curInstanceName,
                PowersIdx = powerIdx,
                Ids = ids,
                Lvcc = lvcc,
                Eqns = eqns,
                Bin = bin,
                IsCheckPassByInstance = oneGradeSearch.InstanceBinCut!.IsBvPass,
                IsSearch = oneGradeSearch.InstanceBinCut.IsSearch,
                FinalStep = powerZone.FinalStep,
                UsedSteps = searchStep,
                ActualSteps = searchStep - actualStepsOffset,
                PatternRows = [.. patternRows.Select(y => y.Where(x => x.Site == site).ToList())]
            };

            instanceData.FailPatternData.AddRange(GetFailPatternData(siteInfoArray, site));
            SetPatternRow(limitLines, ref instanceData);

            return instanceData;
        }

        internal static void PrintCheckedMessage(Action<string, Color> richTextBoxAppend, List<List<SiteInfo>> allDiceInfos, CheckManager checkManager, int errorCount)
        {
            if (checkManager.TotalHarvestSourceCodeCnt == checkManager.TotalHarvestSourceCodeCheckCnt)
            {
                richTextBoxAppend("=> Totally HarvestSourceCode lines are " + checkManager.TotalHarvestSourceCodeCnt + ", and " + checkManager.TotalHarvestSourceCodeCheckCnt + " BV lines had be checked.", Color.Blue);
            }
            else
            {
                richTextBoxAppend("=> Checked HarvestSourceCode lines (" + checkManager.TotalHarvestSourceCodeCheckCnt + ") don't match with datalog (" + checkManager.TotalHarvestSourceCodeCnt + ").", Color.Red);
            }

            if (checkManager.TotalBvCnt == checkManager.TotalBvCheckCnt)
            {
                richTextBoxAppend("=> Totally BV voltage lines are " + checkManager.TotalBvCnt + ", and " + checkManager.TotalBvCheckCnt + " BV lines had be checked.", Color.Blue);
            }
            else
            {
                richTextBoxAppend("=> Checked BV voltage lines (" + checkManager.TotalBvCheckCnt + ") don't match with datalog (" + checkManager.TotalBvCnt + ").", Color.Red);
            }

            if (checkManager.TotalDsscCnt == checkManager.TotalDsscCheckCnt)
            {
                richTextBoxAppend("=> Totally DSSC lines are " + checkManager.TotalDsscCnt + ", and " + checkManager.TotalDsscCheckCnt + " BV lines had be checked.", Color.Blue);
            }
            else
            {
                richTextBoxAppend("=> Checked DSSC lines (" + checkManager.TotalDsscCheckCnt + ") don't match with datalog (" + checkManager.TotalDsscCnt + ").", Color.Red);
            }

            if (checkManager.TotalPowerBinningCnt == checkManager.TotalPowerBinningCheckCnt)
            {
                richTextBoxAppend("=> Totally Power Binning lines are " + checkManager.TotalPowerBinningCnt + ", and " + checkManager.TotalPowerBinningCheckCnt + " BV lines had be checked.", Color.Blue);
            }
            else
            {
                richTextBoxAppend("=> Checked Power Binning lines (" + checkManager.TotalPowerBinningCheckCnt + ") don't match with datalog (" + checkManager.TotalPowerBinningCnt + ").", Color.Red);
            }

            if (errorCount == 0)
            {
                richTextBoxAppend("=> Totally 0 errors had be found.", Color.Blue);
            }
            else
            {
                richTextBoxAppend("=> Totally " + errorCount + " errors had be found.", Color.Orange);
            }

            int bvPassDice = allDiceInfos.SelectMany(x => x).Count(y => y.InstanceList.Any(x => !x.IsCheckPassByInstance));
            if (bvPassDice != 0)
            {
                richTextBoxAppend("=> Totally BV different dice are " + bvPassDice + ".", Color.Red);
            }

            int lvResultDice = allDiceInfos.SelectMany(x => x).Select(x => x.CheckResult.IsLvResultPass).Count(y => !y);
            if (lvResultDice != 0)
            {
                richTextBoxAppend("=> Totally LV result different dice are " + lvResultDice + ".", Color.Red);
            }

            int dsscPassDice = allDiceInfos.SelectMany(x => x).Select(x => x.CheckResult.IsDsscPass).Count(y => !y);
            if (dsscPassDice != 0)
            {
                richTextBoxAppend("=> Totally dssc different dice are " + dsscPassDice + ".", Color.Red);
            }

            int interpolationPassDice = allDiceInfos.SelectMany(x => x).Select(x => x.CheckResult.IsInterpolationPass).Count(y => !y);
            if (interpolationPassDice != 0)
            {
                richTextBoxAppend("=> Totally interpolation different dice are " + interpolationPassDice + ".", Color.Red);
            }

            int powerBinningPassDice = allDiceInfos.SelectMany(x => x).Select(x => x.CheckResult.IsPowerBinningPass).Count(y => !y);
            if (powerBinningPassDice != 0)
            {
                richTextBoxAppend("=> Totally power binning different dice are " + powerBinningPassDice + ".", Color.Red);
            }

            int binoutStatusPassDice = allDiceInfos.SelectMany(x => x).Select(x => x.CheckResult.IsBinoutStatusPass).Count(y => !y);
            if (binoutStatusPassDice != 0)
            {
                richTextBoxAppend("=> Totally binout status different dice are " + binoutStatusPassDice + ".", Color.Red);
            }

            int totalDiffDice = allDiceInfos.SelectMany(x => x).Select(x => x.IsTotalCurrCheckPass).Count(y => !y);
            if (totalDiffDice == 0)
            {
                richTextBoxAppend("=> Totally different dice are 0.", Color.Blue);
            }
            else
            {
                richTextBoxAppend("=> Totally different dice are " + totalDiffDice + ".", Color.Red);
            }
        }
    }
}
