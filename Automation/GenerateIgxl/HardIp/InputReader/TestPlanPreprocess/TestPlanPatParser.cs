using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.Wireless.DVDC.InputObject;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.Enums;
using CommonLib.Utility;

using LogLib.Utility;

namespace Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess
{
    internal class TestPlanPatParser
    {
        //private const string ErrorMsgWrongMeas = "Wrong measure data in 'Meas' column";
        //private const string ErrorMsgWrongForceFormat = "Wrong format of force condition in Test Plan";
        //private const string ErrorMsgWrongLimitValue = "Unrecognied limit value";
        private readonly List<string> _chipletList = new List<string>();

        private static readonly Regex _regexChiplet = new Regex(@"\w+_(?<chiplet>[A-Za-z]\d+$)", RegexOptions.Compiled);

        private readonly TestPlanSheet _planSheet;
        private readonly SweepVoltageResolver _sweepVoltageResolver;
        private readonly CalcEqnResolver _calcEqnResolver;
        private readonly ForceConditionResolver _forceConditionResolver;
        private readonly MeasPinResolver _measPinResolver;

        public TestPlanPatParser(TestPlanSheet planSheet)
        {
            _planSheet = planSheet;
            _chipletList = MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.GetChipletList(TestPlanStatic.PowerInfoSheet);
            _sweepVoltageResolver = new SweepVoltageResolver(planSheet);
            _calcEqnResolver = new CalcEqnResolver(planSheet);
            _forceConditionResolver = new ForceConditionResolver(planSheet, _calcEqnResolver);
            _measPinResolver = new MeasPinResolver(planSheet, _forceConditionResolver);
        }

        /// <summary>
        /// Convert test plan pattern data
        /// </summary>
        public void ConvertTpPatterns()
        {
            _forceConditionResolver.ResetForceConditionLabels();
            //Convert test plan raw data to HardIp patterns
            HardIpPattern pattern = null;
            foreach (PatternRow patternRow in _planSheet.PatternRows)
            {
                try
                {
                    //Replace force condition labels with actual values
                    _forceConditionResolver.ConvertForceLableForPatternNew(patternRow, pattern);
                    if (!string.IsNullOrEmpty(_planSheet.ForceStr))
                    {
                        patternRow.ForceCondition.ForceCondition += ";" + _planSheet.ForceStr;
                    }

                    //Convert test plan raw data to hardIp pattern data
                    pattern = ConvertTpPattern(patternRow);
                    _planSheet.PatternItems.Add(pattern);

                }
                catch (IndexOutOfRangeException ex)
                {
                    ErrorMessageBox.Show($"The execution result is: ConvertTpPatterns IndexOutOfRangeException in {patternRow.Pattern.GetLastPayload()}.\n{ex}");
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="patternRow"></param>
        /// <returns></returns>
#nullable enable
        internal HardIpPattern? ConvertTpPattern(PatternRow? patternRow)
#nullable restore
        {
            if (patternRow == null)
            {
                return null;
            }

            var pattern = new HardIpPattern
            {
                SheetName = _planSheet.SheetName,
                RowNum = patternRow.RowNum,
                ColumnNum = patternRow.PatternColumnNum,
                Failflag = patternRow.FailFlag,
                SiteFlag = patternRow.SiteFlag,
                Enable = patternRow.Enable,
                Pattern = patternRow.Pattern,
                TtrStr = patternRow.TtrStr,
                PartStr = patternRow.PartStr,
                NoBinoutStr = new HashSet<string>(patternRow.NoBinOutStr.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()), StringComparer.OrdinalIgnoreCase),
                DupIndex = patternRow.DupIndex
            };

            string newMisc = "";
            if (patternRow.Pattern.IsMultiTimeDomain())
            {
                int refMtDnum = patternRow.Pattern.RealPatternName.Split('#').ToList().FindIndex(pat => pat.Equals(patternRow.Pattern.TestPlanPatternName));
                foreach (string misc in patternRow.MiscInfo.Split(';'))
                {
                    if (!misc.Any())
                    {
                        continue;
                    }

                    if (Regex.IsMatch(misc, "ref_subblock", RegexOptions.IgnoreCase) && refMtDnum != -1)
                    {
                        pattern.SubBlock = misc.Split(':')[1];
                        string refPatterns = misc.Split(':')[1].Split('#')[refMtDnum];
                        newMisc += misc.Split(':')[0] + ":" + refPatterns + ";";
                    }
                    else
                    {
                        newMisc += misc + ";";
                    }
                }
            }
            else
            {
                foreach (string misc in patternRow.MiscInfo.Split(';'))
                {
                    if (!misc.Any())
                    {
                        continue;
                    }

                    if (Regex.IsMatch(misc, "subblock", RegexOptions.IgnoreCase))
                    {
                        pattern.SubBlock = misc.Split(':')[1];
                        break;
                    }
                }
                newMisc = patternRow.MiscInfo;
            }
            pattern.MiscInfo = newMisc;
            pattern.ForceCondition = patternRow.ForceCondition;
            pattern.RegisterAssignment = patternRow.RegisterAssignment.Trim(';');
            pattern.Chiplet = GetChiplet(_planSheet.SheetName, _chipletList);

            //Read measure sequence info from test plan
            pattern.TestPlanSequences = ReadSequenceInfoNew(patternRow, pattern);
            //Post test force condition
            pattern.InterposePostTest = patternRow.PostPatForceCondition;

            #region Convert force condition str which assigned in the same row as pattern
            //Get MCG mode setting
            pattern.TimeSetUsed.McgSetting = patternRow.ForceCondition.GetMcgSetting();
            //Get interpose pre pattern force condition
            pattern.ForceConditionList = _forceConditionResolver.DivideForceCondition(patternRow.ForceCondition.GetPrePatForceCondition(), patternRow.RowNum);
            #endregion
            //Convert measure pins' information
            pattern.TestPlanSequencesRf = ReadSequenceInfoRf(patternRow, _measPinResolver.GetMeasPins(patternRow, pattern), pattern);
            pattern.OriMeasPins.AddRange(_measPinResolver.GetMeasPins(patternRow, pattern));
            pattern.MeasPins.AddRange(pattern.OriMeasPins);
            //Calc_Eqn
            pattern.CalcEqn = _calcEqnResolver.GetPatternCalcEqn(pattern);
            //RF Interpose
            pattern.RfInterPose = patternRow.RfInterPose;

            //Loop Flow loop
            pattern.SweepVoltage = _sweepVoltageResolver.GetSweepVoltage(patternRow);
            //Shmoo
            pattern.Shmoo = patternRow.ForceCondition.GetShmoo(_planSheet, pattern, patternRow.ForceCondition.ForceCondition, CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, ""));

            pattern.RegAssignName = CommonGenerator.GetRegAssignName(pattern);

            pattern.TestName = patternRow.SpecifyTestName;

            string subblock = CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, "");
            pattern.SubBlockCopy = subblock;
            pattern.SheetSubBlockName = Combination.CombineByUnderLine(pattern.SheetName, subblock);

            (string ateTestCondition, string overlayName) binCutOverlay = UpdateBinCutOverlayName(pattern.MiscInfoDict);
            pattern.BinCutOverlay = binCutOverlay;
            return pattern;
        }

        public void ConvertWirelessTpPatterns()
        {
            int cnt = 0;
            foreach (PatternRow patternRows in _planSheet.PatternRows)
            {
                //    string vbt = SearchInfo.GetVbtNameByMiscInfo(patternRows.MiscInfo);
                //    _planSheet.PatternItems[cnt].FunctionName = string.IsNullOrEmpty(vbt) ? VbtFunctionLib.DcdvTrim : vbt;
                var wirelessPatternRow = (WirelessPatternRow)patternRows;
                _planSheet.PatternItems[cnt].WirelessData = wirelessPatternRow.WirelessData;
                cnt++;
            }
        }

        private (string ateTestCondition, string overlayName) UpdateBinCutOverlayName(Dictionary<string, string> miscInfo)
        {
            bool funcUsing = miscInfo.Any(x => x.Key.Equals("Func", StringComparison.OrdinalIgnoreCase) && x.Value.Equals("CreateOverlay", StringComparison.OrdinalIgnoreCase));
            if (!funcUsing || !miscInfo.ContainsKey("ateTestCondition") || !miscInfo.ContainsKey("overlayName"))
            {
                return (null, null);
            }

            string ateTestCondition = miscInfo["ateTestCondition"];
            string overlayName = miscInfo["overlayName"];
            if (!string.IsNullOrEmpty(ateTestCondition) && !string.IsNullOrEmpty(overlayName))
            {
                return (ateTestCondition, overlayName);
            }
            return (null, null);
        }

        private List<TestPlanSequence> ReadSequenceInfoRf(PatternRow patternRow, List<MeasPin> measpins, HardIpPattern pattern)
        {
            var nonSeqItems = new List<string> { MeasType.MeasLimit, MeasType.MeasCalc, MeasType.MeasC };
            var testPlanSequence = new List<TestPlanSequence>();
            _ = HardIpService.GetHardIpInfo(pattern).SeqInfo.Count > 0;
            int sequenceIndex = 1;
            string measStr, forceStr;
            int startRow, endRow;
            if (!measpins.Exists(p => p.MeasType.Equals(MeasType.WiMeas) || p.MeasType.Equals(MeasType.WiSrc)))
            {
                return testPlanSequence;
            }

            foreach (PatChildRow childRow in patternRow.PatChildRows)
            {
                startRow = ((PatSubChildRow)childRow).TpRows.First().RowNum;
                endRow = ((PatSubChildRow)childRow).TpRows.Last().RowNum;
                var sequence = new TestPlanSequence(startRow, endRow, sequenceIndex++);
                bool isSeqItems = false;
                foreach (TestPlanRow subChildRow in ((PatSubChildRow)childRow).TpRows)
                {
                    measStr = subChildRow.Meas;
                    forceStr = subChildRow.ForceCondition;
                    if (nonSeqItems.All(p => Regex.IsMatch(measStr, p, RegexOptions.IgnoreCase)))
                    {
                        continue;
                    }

                    isSeqItems = true;
                    if (!string.IsNullOrEmpty(forceStr))
                    {
                        sequence.ForceCondition.AddRange(forceStr.Split(';').ToList());
                    }
                }
                if (isSeqItems)
                {
                    testPlanSequence.Add(sequence);
                }
            }

            return testPlanSequence;
        }

        private List<TestPlanSequence> ReadSequenceInfoNew(PatternRow patternRow, HardIpPattern pattern)
        {
            List<string> nonSeqItems = LocalSpecs.Options.Device != EnumDevice.LCD /*&& LocalSpecs.Optionals.Device != DeviceEnum.RF*/
                ?
                new List<string> { MeasType.MeasLimit, MeasType.MeasCalc, MeasType.MeasC } :
                new List<string> { MeasType.MeasC };
            var testPlanSequence = new List<TestPlanSequence>();
            try
            {
                HardIpService.GetHardIpInfo(pattern);
                _ = HardIpService.GetHardIpInfo(pattern).SeqInfo.Count > 0;
                int sequenceIndex = 1;
                string measStr, forceStr;
                int startRow, endRow;
                foreach (PatChildRow childRow in patternRow.PatChildRows)
                {
                    startRow = ((PatSubChildRow)childRow).TpRows.First().RowNum;
                    endRow = ((PatSubChildRow)childRow).TpRows.Last().RowNum;
                    var sequence = new TestPlanSequence(startRow, endRow, 0);
                    bool isSeqItems = false;
                    foreach (TestPlanRow subChildRow in ((PatSubChildRow)childRow).TpRows)
                    {
                        measStr = subChildRow.Meas;
                        forceStr = subChildRow.ForceCondition;
                        if (nonSeqItems.Any(p => Regex.IsMatch(measStr.Trim().Split(' ')[0], p, RegexOptions.IgnoreCase)))
                        {
                            continue;
                        }

                        isSeqItems = true;
                        if (!string.IsNullOrEmpty(forceStr))
                        {
                            if (Regex.IsMatch(measStr, MeasType.MeasR2, RegexOptions.IgnoreCase))
                            {
                                sequence.ForceCondition.AddRange(forceStr.Split(';').Select(x => x.Replace(",", "&")).ToList());
                            }
                            else
                            {
                                sequence.ForceCondition.AddRange(forceStr.Split(';').ToList());
                            }
                        }
                    }
                    sequence.ForceCondition = sequence.ForceCondition.Distinct().ToList();
                    if (isSeqItems)
                    {
                        sequence.SeqIndex = sequenceIndex++;
                        testPlanSequence.Add(sequence);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }

            return testPlanSequence;
        }

        internal string GetChiplet(string sheetName, List<string> chipletList)
        {
            Match match = _regexChiplet.Match(sheetName);
            if (match.Success)
            {
                return match.Groups["chiplet"].Value;
            }

            return chipletList?.FirstOrDefault() ?? string.Empty;
        }

    }
}
