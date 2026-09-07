using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTable;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlow;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlSheets;

using LogLib.Utility;

using TestPlanLib.Scan;
using TestPlanLib.Static;

namespace Automation.Reader.TestPlan.ClockOut
{
    public class ClockOutMeasGenerator
    {
        private readonly List<ClockMeasRow> _clockRows = new List<ClockMeasRow>();
        private readonly BinCutFinalInstanceRows _binCutInstanceRows;

        public List<string> InstanceClockNonUsedItems = new List<string>();

        private int _noPatDupIndex = 1;
        private readonly List<string> _jobList = new List<string>();
        private readonly List<PatSet> _generatedPatSets = new List<PatSet>();
        private readonly string _sheetName = "";

        public ClockOutMeasGenerator(string sheetName, List<ClockMeasRow> clockRows, BinCutFinalInstanceRows binCutInstanceRows)
        {
            _sheetName = sheetName.Replace("_", "");
            _clockRows = clockRows;
            _binCutInstanceRows = binCutInstanceRows;
            _jobList = LocalSpecs.AllJobs;
        }

        public void WorkFlow()
        {
            try
            {
                var clockCheckInstanceRows = new InstanceRows();
                var hipPatterns = new List<HardIpPattern>();
                // setup ClockOutMeas Pattern sets
                if (_binCutInstanceRows == null || _binCutInstanceRows.Count == 0)
                {
                    return;
                }

                var clockGroups = _binCutInstanceRows.GroupBy(p => p.BinCutInstanceRow.FlowName).ToDictionary(p => p.Key, p => p.ToList());
                InstanceClockNonUsedItems = _binCutInstanceRows.Select(p => p.BinCutInstanceRow.FlowName).Distinct().ToList();

                foreach (ClockMeasRow clockRow in _clockRows)
                {
                    if (clockRow.Type.Equals("MSX004DEBUG_GPU_BIST_CLOCK", StringComparison.OrdinalIgnoreCase))
                    {
                    }
                    if (clockRow.PinInfo.Count == 0)
                    {
                        continue;
                    }
                    if (!clockGroups.TryGetValue(clockRow.Type, out List<BinCutFinalInstanceRow> group))
                    {
                        clockRow.ErrMsg = $"{clockRow.Type} Not Found in Instance Clock sheet";
                    }
                    else
                    {
                        foreach (BinCutFinalInstanceRow matchItem in group)
                        {
                            (InstanceRows instanceRows, List<HardIpPattern> measSet) = ConvertToHipPatternBincut(clockRow, matchItem, _sheetName);
                            if (measSet == null)
                            {
                                continue;
                            }

                            clockCheckInstanceRows.AddRange(instanceRows);
                            hipPatterns.AddRange(measSet);
                        }
                        InstanceClockNonUsedItems.Remove(clockRow.Type);
                    }
                }

                WriteErrorReport(_clockRows, InstanceClockNonUsedItems, clockGroups);
                if (hipPatterns.Count == 0)
                {
                    return;
                }


                #region GenPatSet

                if (_generatedPatSets.Count > 0)
                {
                    string patSetSheetName = "PatSets_ClockOut";
                    var patSetSheet = new PatSetSheet(patSetSheetName);
                    if (
                        TestProgram.IgxlWorkBk.PatSetSheets.Values.FirstOrDefault(
                            p => p.IgxlSheetName.Equals(patSetSheetName)) == null)
                    {
                        TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirPatSetsAll, patSetSheet);
                    }
                    else
                    {
                        patSetSheet =
                            TestProgram.IgxlWorkBk.PatSetSheets.Values.FirstOrDefault(
                                p => p.IgxlSheetName.Equals(patSetSheetName));
                    }
                    _generatedPatSets.ForEach(p => AddPatSet(patSetSheet, p));
                }

                #endregion

                //Generate Instance
                var instSheet = new InstanceSheet("TestInst_" + CommonGenerator.GetHardipSheetName(_sheetName).ToUpper())
                {
                    Rows = clockCheckInstanceRows
                };

                //Generate Flow
                var hardIpInputData = new HardIpInputData(new HardIpParaData(EnumBlock.HardIp));
                var flowSheetGenerator = new FreqFlowSheetGenerator(hardIpInputData, _sheetName, hipPatterns);

                List<SubFlowSheet> subFlowSheets = flowSheetGenerator.GenerateFlowSheet();

                foreach (SubFlowSheet flowSheet in subFlowSheets)
                {
                    if (flowSheet.Rows.Count > 0)
                    {
                        TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirHardIp, flowSheet);
                    }
                }

                if (instSheet.Rows.Count > 0)
                {
                    TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirHardIp, instSheet);
                }

                const string hardIpBinTableSheetName = "Bin_Table_ClockOut";
                var binTableSheet = new BinTableSheet(hardIpBinTableSheetName);
                var errorBinNums = new List<string>();
                var duplicateParameter = new List<string>();
                BlockBinTableGeneratorBase blockBinGenerator = null;

                blockBinGenerator = new FreqBinTableGenerator(hardIpInputData, binTableSheet, _sheetName, hipPatterns, duplicateParameter, errorBinNums);
                blockBinGenerator.GenerateBinTableRows();
                if (binTableSheet.Rows.Count > 0)
                {
                    TestProgram.IgxlWorkBk.AddBinTblSheet(FolderStructure.DirCommon, binTableSheet);
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        private (InstanceRows, List<HardIpPattern>) ConvertToHipPatternBincut(ClockMeasRow specRow, BinCutFinalInstanceRow row, string sheet)
        {
            var hardIpInputData = new HardIpInputData(new HardIpParaData(EnumBlock.HardIp))
            {
                HardIpDcSheet = TestPlanStatic.HardIpDcSheet
            };
            var instanceRows = new InstanceRows();
            var hardIpPatterns = new List<HardIpPattern>();
            string voltage = row.GetVoltageType();
            voltage = voltage == "UnknowType" ? "NV" : voltage;

            //setup init patterns
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass(string.Join("+", row.PatternList)),
                SheetName = sheet,
                BlockType = specRow.PMode.Equals("bist", StringComparison.OrdinalIgnoreCase) ? "MBist" : "Scan",
                FunctionName = VbtFunctionLibShared.FunctionalName,
                MiscInfo = $"{voltage}Only"
            };
            pattern.ForceCondition.ForceCondition += string.Format($"DC:{row.GetDcCategory()}");
            GenPatSet(new List<HardIpPattern> { pattern }, specRow, row);
            var hardIpSheet = new HardIpSheet
            {
                SheetName = _sheetName,
                Rows = new List<HardIpPattern> { pattern }
            };
            var clockCheckHipInstanceGenerator = new FreqPllBlockInsGenerator(hardIpInputData, _sheetName, hardIpSheet);
            var firstInstance = clockCheckHipInstanceGenerator.GenBlockInsRows(voltage).SelectMany(x => x.Rows).ToList();
            instanceRows.AddRange(firstInstance);
            hardIpPatterns.Add(pattern);

            //setup payload pattern
            pattern = new HardIpPattern();
            if (!string.IsNullOrEmpty(specRow.Pattern))
            {
                pattern.Pattern = new PatternClass(specRow.Pattern);
                pattern.DupIndex = ++_noPatDupIndex;
            }
            else
            {
                pattern.Pattern = new PatternClass(string.Format($"Instance:{specRow.PMode}_Meas_{++_noPatDupIndex}_{voltage}"));
            }

            pattern.SheetName = sheet;
            pattern.BlockType = specRow.PMode.Equals("bist", StringComparison.OrdinalIgnoreCase) ? "MBist" : "Scan";
            pattern.TestPlanSequences = new List<TestPlanSequence> { new TestPlanSequence(1, 1, 1) };
            pattern.MeasPins = GetMeasPins(specRow);
            pattern.MiscInfo = $"{voltage}Only";
            pattern.ForceCondition.ForceCondition += string.Format($"DC:{row.GetDcCategory()}");
            GenPatSet(new List<HardIpPattern> { pattern }, specRow, row);
            hardIpSheet = new HardIpSheet
            {
                SheetName = _sheetName,
                Rows = new List<HardIpPattern> { pattern }
            };
            clockCheckHipInstanceGenerator = new FreqPllBlockInsGenerator(hardIpInputData, _sheetName, hardIpSheet);
            var secondInstance = clockCheckHipInstanceGenerator.GenBlockInsRows(voltage).SelectMany(x => x.Rows).ToList();
            UpdateDigSrc(ref secondInstance, specRow);
            if (string.IsNullOrEmpty(specRow.Pattern) && firstInstance.Count > 0)
            {
                secondInstance.ForEach(x =>
                {
                    x.AcCategory = firstInstance.FirstOrDefault().AcCategory;
                    x.AcSelector = firstInstance.FirstOrDefault().AcSelector;
                    x.TimeSets = firstInstance.FirstOrDefault().TimeSets;
                    x.DcCategory = firstInstance.FirstOrDefault().DcCategory;
                    x.DcSelector = firstInstance.FirstOrDefault().DcSelector;
                    x.PinLevels = firstInstance.FirstOrDefault().PinLevels;
                });
            }
            instanceRows.AddRange(secondInstance);
            hardIpPatterns.Add(pattern);

            return (instanceRows, hardIpPatterns);
        }
        private void UpdateDigSrc(ref List<InstanceRow> instanceRows, ClockMeasRow specRow)
        {
            string digSrc = specRow.Digsrc;
            string pattern = specRow.Pattern;
            if (string.IsNullOrEmpty(digSrc) || string.IsNullOrEmpty(pattern))
            {
                return;
            }

            foreach (InstanceRow instanceRow in instanceRows)
            {
                string digSrcPin = SearchInfo.GetSrcPin(HardIpService.GetHardIpInfo(pattern));
                instanceRow.SetArgument("digSrcPin", string.IsNullOrEmpty(digSrcPin) ? "JTAG_TDI" : digSrcPin);
                instanceRow.SetArgument("digSrcEquation", "C");
                instanceRow.SetArgument("digSrcAssignment", $"C={digSrc}");
            }
        }
        private List<MeasPin> GetMeasPins(ClockMeasRow row)
        {
            var result = new List<MeasPin>();
            foreach (KeyValuePair<string, List<string>> pinInfo in row.PinInfo)
            {
                var measPin = new MeasPin
                {
                    PinName = pinInfo.Key,
                    MeasType = MeasType.MeasF,
                    SequenceIndex = 1
                };

                string hilimit = "";
                string lolimit = "";
                for (int i = 0; i < pinInfo.Value.Count; i++)
                {
                    if (i == 0)
                    {
                        lolimit = pinInfo.Value[i];
                    }
                    else
                    {
                        hilimit = pinInfo.Value[i];
                    }
                }
                List<List<MeasLimit>> limits = GetLimits(lolimit, hilimit);
                measPin.MeasLimitsH = limits[0];
                measPin.MeasLimitsL = limits[1];
                measPin.MeasLimitsN = limits[2];
                measPin.TestName = $"{row.Type}_{row.RowNum}";
                result.Add(measPin);
            }
            return result;
        }

        private List<List<MeasLimit>> GetLimits(string lo, string hi)
        {
            var limits = new List<List<MeasLimit>>();
            var limitsH = new List<MeasLimit>();
            var limitsL = new List<MeasLimit>();
            var limitsN = new List<MeasLimit>();
            limits.Add(limitsH);
            limits.Add(limitsL);
            limits.Add(limitsN);

            foreach (string jobName in _jobList)
            {
                var limitH = new MeasLimit(jobName);
                var limitL = new MeasLimit(jobName);
                var limitN = new MeasLimit(jobName);
                limitsH.Add(limitH);
                limitsL.Add(limitL);
                limitsN.Add(limitN);

                //limit valus contains HV,LV,NV

                //Set hi limit for Hv,Lv,Nv

                limitH.HiLimit = hi;
                limitL.HiLimit = hi;
                limitN.HiLimit = hi;
                //Set lo limit for Hv,Lv,Nv

                limitH.LoLimit = lo;
                limitL.LoLimit = lo;
                limitN.LoLimit = lo;

            }

            return limits;
        }

        public void GenPatSet(List<HardIpPattern> patterns, ClockMeasRow clockrow, BinCutFinalInstanceRow bincutInstance)
        {
            var pllFreqSet = new List<HardIpPattern>();
            var newPatSet = new PatSet();
            try
            {
                foreach (HardIpPattern pattern in patterns)
                {
                    pllFreqSet.Add(pattern);
                    if (pattern.MeasPins.Count == 0)
                    {
                        string regDSelSrm = "_[D]*SRA*M";
                        HardIpPattern burstPattern = pllFreqSet.FirstOrDefault(p => Regex.IsMatch(p.Pattern.GetLastPayload(), regDSelSrm, RegexOptions.IgnoreCase)) ?? pllFreqSet[0];
                        burstPattern.Pattern.InstancePatternName = new List<string>();

                        newPatSet.PatSetName = GenPllPatSetName(pllFreqSet.Where(p => !p.Pattern.GetLastPayload().Contains(":")).ToList(), bincutInstance);
                        burstPattern.Pattern.InstancePatternName.Add(newPatSet.PatSetName);
                        burstPattern.Pattern.InstancePayloadName.Add(newPatSet.PatSetName);
                        burstPattern.TestName = newPatSet.PatSetName;
                        foreach (List<string> pats in pattern.Pattern.PatternSetList)
                        {
                            foreach (string pat in pats)
                            {
                                var patSet = new PatSetRow
                                {
                                    Burst = "No",
                                    File = pat
                                };
                                newPatSet.AddRow(patSet);
                            }

                        }
                    }
                }
                if (newPatSet.PatSetRows.Count > 0)
                {
                    _generatedPatSets.Add(newPatSet);
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        private void AddPatSet(PatSetSheet newSheet, PatSet newPatSet)
        {
            if (newSheet.IsExistTheSamePatSet(newPatSet, out _))
            {
                return;
            }

            if (newPatSet.PatSetRows.Count == 1 &&
                newPatSet.PatSetRows[0].File.Equals(newPatSet.PatSetName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            newSheet.AddRow(newPatSet);
        }

        private string GenPllPatSetName(List<HardIpPattern> patterns, BinCutFinalInstanceRow bincutInstance)
        {
            string module = GetModuleName(patterns.Last().Pattern.GetLastPayload());
            foreach (HardIpPattern pattern in patterns)
            {
                if (string.IsNullOrEmpty(module))
                {
                    module = GetModuleName(pattern.Pattern.GetLastPayload());
                }
            }
            if (string.IsNullOrEmpty(module))
            {
            }
            var nameList = new List<string> { bincutInstance.PatSetName };

            return CheckRepeatName(nameList);
        }

        private string CheckRepeatName(List<string> nameList)
        {
            string name = string.Join("_", nameList);
            int index = 1;
            while (_generatedPatSets.Exists(p => p.PatSetName.Equals(name)))
            {
                string tmpName = string.Join("_", nameList);
                name = tmpName + "_" + index;
                index++;
            }
            return name;
        }

        private string GetModuleName(string pattern, string module = "")
        {
            string[] sgmts = pattern.Split('_');
            if (sgmts.Length <= 9)
            {
                return "";
            }

            string result = module;
            switch (sgmts[2].ToUpper())
            {
                case "L":
                    result = "Gfx";
                    break;
                case "C":
                    result = "Cpu";
                    break;
                case "S":
                    result = "Soc";
                    break;
            }

            return result;
        }

        private void WriteErrorReport(List<ClockMeasRow> clockRows, List<string> nonUsedItems, Dictionary<string, List<BinCutFinalInstanceRow>> bincutGroups)
        {
            foreach (ClockMeasRow notFoundItem in clockRows.Where(p => !string.IsNullOrEmpty(p.ErrMsg)))
            {
                ErrorReportManager.AddError(ClockCheckErrorType.E_MissingFlow_01, NeededSheets.ClockPllMeas, notFoundItem.RowNum, 1, [notFoundItem.Type]);
            }

            foreach (string notUsedItem in nonUsedItems)
            {
                if (bincutGroups.TryGetValue(notUsedItem, out List<BinCutFinalInstanceRow> group))
                {
                    foreach (BinCutFinalInstanceRow member in group)
                    {
                        ErrorReportManager.AddError(ClockCheckErrorType.E_MissingSetting_01, member.BinCutInstanceRow.SheetName, member.BinCutInstanceRow.RowNum, 1,
                            [NeededSheets.ClockPllMeas]);
                    }
                }
            }
        }
    }
}
