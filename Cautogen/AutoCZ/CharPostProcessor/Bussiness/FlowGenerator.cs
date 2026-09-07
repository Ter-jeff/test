using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.Controller;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility.TestNumManager;
using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;
using Cautogen.AutoCZ.CharPostProcessor.Utility.VbtModuleManager;
using Cautogen.Utility;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class FlowGenerator
    {
        private readonly bool _genCharNoUsed;
        private readonly bool _genPatNoUsed;
        private readonly bool _genTNum;
        private readonly bool _genTmpsOnFlow;
        private readonly bool _genReadEcid;
        private readonly bool _ignoreHfLimits;
        private readonly bool _genCSharp;
        private readonly string _enableWords;
        private readonly string _outputFolder;
        private readonly string _flowTmpsName;
        private readonly string _job;
        private bool _isNeedApplyBincutVoltage = false;
        private readonly FlowRow _binShmooAlarmFlowRow = new FlowRow { Opcode = "BinTable", Parameter = "Bin_initFlow_BinOut" };

        public FlowGenerator(InputParam inputParam)
        {
            _genCharNoUsed = inputParam.GenCharNotUse;
            _genPatNoUsed = inputParam.GenPatNotUse;
            _genTNum = inputParam.GenTNum;
            _genTmpsOnFlow = inputParam.GenTmpsOnFlow;
            _enableWords = inputParam.EnableWords;
            _outputFolder = inputParam.GenTxtOnly ? inputParam.OutputFolder : Path.Combine(inputParam.OutputFolder, ConstData.CzFolder);
            _flowTmpsName = inputParam.FlowTmpsName;
            _genReadEcid = inputParam.GenReadEcid;
            _ignoreHfLimits = inputParam.IgnoreHfLimits;
            _genCSharp = inputParam.GenCSharp;
            _job = inputParam.JobName;
        }

        public SubFlowSheet Generate(List<CharPlanSheet> charPlanSheets)
        {
            GeneralFunc.WriteMessage("Generating flow sheets... ");
            var preScanList = new List<string>();
            var testFlowList = new List<string>();
            var hipFlowList = new List<string>();
            _isNeedApplyBincutVoltage = charPlanSheets.SelectMany(p => p.CharList).Any(p => !string.IsNullOrEmpty(p.ApplyVoltageFromBinCut));
            _ProcessTMPS_Flow(charPlanSheets);
            _ProcessNwireFlow(charPlanSheets);
            if (BasMain.VbtFunctionLib.IsFunctionExist("onProgramStartedBinOutFunction"))
            {
                _ProcessMainInitFlow();
                _ProcessMainInitEnableWd();
            }

            foreach (CharPlanSheet planSheet in charPlanSheets)
            {
                TestNumMain.NextBlock();  //Set TNum for each block

                if (planSheet.IsHardIp)
                {
                    hipFlowList.Add(_HipHandler(planSheet));
                }
                else
                {
                    _AtpgHandler(planSheet, preScanList, testFlowList);
                }
            }
            return WriteFlowChar(preScanList, testFlowList, hipFlowList);
        }

        private SubFlowSheet WriteFlowChar(IEnumerable<string> preScanList, IEnumerable<string> testFlowList, IEnumerable<string> hipFlowList)
        {
            SubFlowSheet mainFlow = LocalSpecs.TestProgram.GetMainFlowSheet();
            FlowRow charPlanFromTp = mainFlow.Rows.FirstOrDefault(p => !string.IsNullOrEmpty(p.ColumnA));
            SubFlowSheet mainChar = charPlanFromTp == null ? new SubFlowSheet("Flow_Char") : new SubFlowSheet($"{charPlanFromTp.Parameter}");

            if (mainFlow == null)
            {
                return mainChar;
            }

            string seterrbin = "999";
            foreach (FlowRow row in mainFlow.Rows.Where(row => row.Opcode.Equals("set-error-bin", StringComparison.CurrentCultureIgnoreCase)))
            {
                seterrbin = row.BinFail;
                break;
            }
            if (string.IsNullOrEmpty(seterrbin))
            {
                BinTableSheet bintable = LocalSpecs.TestProgram.BintableSheets.FirstOrDefault(p => Regex.IsMatch(p.Name, "BinTable|Bin_Table", RegexOptions.IgnoreCase));
                if (bintable != null)
                {
                    foreach (BinTableRow binrow in bintable.Rows.Where(binrow => binrow.Name.Equals("Bin_DC_open", StringComparison.CurrentCulture)))
                    {
                        seterrbin = binrow.Bin;
                        break;
                    }
                }
            }

            mainChar.AddRow(new FlowRow { Opcode = ConstData.SetErrorOpCode, BinFail = seterrbin, SortFail = "999" }); // error bin

            if (!string.IsNullOrEmpty(_enableWords))
            {
                //mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = "SetEnableWords" });
                mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = "PrintEnableWords_Header" });
                mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = "PrintEnableWords" });
                mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = "PrintEnableWords_Footer" });
            }

            mainChar = GenerateShmooSetupInit(mainChar); //generate shmoo init flag 2017/7/13
            if (!string.IsNullOrEmpty(LocalSpecs.InputParam.PatListFile))
            {
                mainChar.AddRow(new FlowRow
                {
                    Opcode = "Print",
                    Parameter = "\"" + Path.GetFileName(LocalSpecs.InputParam.PatListFile) + "\""
                });
            }
            mainChar.AddRow(new FlowRow
            {
                Opcode = "Print",
                Parameter = "\"" + Path.GetFileName(LocalSpecs.InputParam.CharFile) + "\""
            });
            mainChar.AddRow(new FlowRow { Opcode = "Print", Parameter = "\"" + ConstData.DefaultSigsrcValue + "\"" });
            if (_genReadEcid)
            {
                mainChar.AddRow(new FlowRow { Opcode = "Call", Parameter = "Flow_eFuse_ECID", Enable = "!Pgm2file&&!Char_CP", Job = "!CP1,!CP2" });
            }

            mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = "Relay_ON_Default" });
            mainChar.AddRow(new FlowRow { Opcode = "Call", Parameter = "Flow_nWire" });
            mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = "Datalog_Setup_Char", Env = "X" });
            mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = "ReadWaferData_Char", Env = "X" });
            mainChar.AddRow(new FlowRow { Opcode = "create-site-var", Parameter = "Function_Result" });

            // add const flow 2017/6/27 Raze
            mainChar.AddRow(new FlowRow { Opcode = "flag-clear", Parameter = "F_Check_Shmoo_Hole_Ratio_Within_Spec" });
            mainChar.AddRow(new FlowRow { Opcode = "flag-clear", Parameter = "F_Check_Shmoo_Allfail_Ratio_Within_Spec" });
            mainChar.AddRow(new FlowRow { Opcode = "flag-clear", Parameter = "F_Check_Shmoo_Alarm_Ratio_Within_Spec" });
            mainChar.AddRow(new FlowRow { Opcode = "flag-clear", Parameter = "F_Dummy" });
            mainChar.AddRow(new FlowRow { Opcode = "flag-clear", Parameter = "F_Shmoo_Alarm" });
            mainChar.AddRow(new FlowRow { Opcode = "flag-clear", Parameter = "F_initFlow_Binout" });
            mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = "Char_Setup_Gating", FailAction = "F_initFlow_Binout" });
            mainChar.AddRow(new FlowRow { Opcode = "BinTable", Parameter = "Bin_initFlow_BinOut" });
            mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = ConstData.EnableCzMode });
            // init enable words for cz
            mainChar.AddRow(new FlowRow { Opcode = "nop", Enable = "Char_TMPS" });
            mainChar.AddRow(new FlowRow { Opcode = "disable-flow-word", Parameter = "Char_TMPS" });
            mainChar.AddRow(new FlowRow { Opcode = "enable-flow-word", Parameter = "Char_TMPS", Enable = "Char_TMPS" });
            mainChar.AddRow(new FlowRow { Opcode = "nop", Enable = "Disable_MemoryClr" });
            mainChar.AddRow(new FlowRow { Opcode = "disable-flow-word", Parameter = "Disable_MemoryClr" });
            mainChar.AddRow(new FlowRow { Opcode = "enable-flow-word", Parameter = "Disable_MemoryClr", Enable = "Disable_MemoryClr" });
            mainChar.AddRow(new FlowRow { Opcode = "nop", Enable = "Enable_HAC" });
            mainChar.AddRow(new FlowRow { Opcode = "nop", Enable = "Enable_HF" });
            mainChar.AddRow(new FlowRow { Opcode = "nop", Enable = "Enable_HIO" });
            mainChar.AddRow(new FlowRow { Opcode = "enable-flow-word", Parameter = "Enable_HAC", Enable = "Char_HIP" });
            mainChar.AddRow(new FlowRow { Opcode = "enable-flow-word", Parameter = "Enable_HF", Enable = "Char_HIP" });
            mainChar.AddRow(new FlowRow { Opcode = "enable-flow-word", Parameter = "Enable_HIO", Enable = "Char_HIP" });
            if (_isNeedApplyBincutVoltage)
            {
                mainChar.AddRow(new FlowRow { Opcode = "Call", Parameter = "Flow_eFuse_BankRead", Job = "!CP1", Enable = "Enable_ApplyBVToCZ" });
                mainChar.AddRow(new FlowRow { Opcode = "Call", Parameter = "Flow_DCTEST_IDS", Job = "CP1", Enable = "Enable_ApplyBVToCZ" });
                mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = "Judge_store_IDS", Enable = "Enable_ApplyBVToCZ" });
                mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = "Read_DVFM_To_GradeVDD", Job = "!CP1", Enable = "Enable_ApplyBVToCZ" });
            }
            foreach (string scanSheet in preScanList)
            {
                mainChar.AddRow(new FlowRow { Opcode = "Call", Parameter = scanSheet, Enable = "PreScan" });
            }

            foreach (string testSheet in testFlowList)
            {
                string sheetName = Regex.Match(testSheet, @"_CZ_(?<sheet>\w+)$").Groups["sheet"].ToString();
                mainChar.AddRow(new FlowRow { Opcode = "Call", Parameter = testSheet, Enable = "Enable_" + sheetName });
            }
            foreach (string hipSheet in hipFlowList.Where(x => x.EndsWith("_IDS", StringComparison.OrdinalIgnoreCase)).ToList())
            {
                mainChar.AddRow(new FlowRow { Opcode = "Call", Parameter = hipSheet });
            }

            if (hipFlowList.Any(x => !x.EndsWith("_IDS", StringComparison.OrdinalIgnoreCase)))
            {
                mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = "Relay_On_HARDIP" });
            }

            foreach (string hipSheet in hipFlowList.Where(x => !x.EndsWith("_IDS", StringComparison.OrdinalIgnoreCase)).ToList())
            {
                mainChar.AddRow(new FlowRow { Opcode = "Call", Parameter = hipSheet });
            }

            // edit DC_Conti flow for request 2017/7/4
            //GenFlowDcConti2Nd(mainChar, _outputFolder); // turn off conti 2nd 2022/12/26
            //insert htol and ttr option 2017/8/16
            foreach (string htolSheet in LocalSpecs.HtolAndTtr[ConstData.Htol])
            {
                mainChar.AddRow(new FlowRow { Opcode = "Call", Parameter = htolSheet, Enable = "Enable_" + ConstData.Htol });
            }

            foreach (string ttrSheet in LocalSpecs.HtolAndTtr[ConstData.Ttr])
            {
                mainChar.AddRow(new FlowRow { Opcode = "Call", Parameter = ttrSheet, Enable = "Enable_" + ConstData.Ttr });
            }
            AddShmooAbnormalBinOutFlowRows(mainChar);


            // set device
            //mainChar.AddRow(new FlowRow { Opcode = "set-device", SortPass = "1", BinPass = "1", Result = "Pass" });
            mainChar.AddRow(new FlowRow { Opcode = "return" });

            // export Flow_Char
            string mainCharFile = Path.Combine(_outputFolder, mainChar.Name + ".txt");
            mainChar.Write(mainCharFile, LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0");
            LocalSpecs.GenSheets.Add(mainChar);
            return mainChar;
        }

        private string _HipHandler(CharPlanSheet planSheet)
        {
            /* return Flow_CZ_<charPlanSheetName> */
            bool existHtol = planSheet.CharList.Exists(a => a.Htol);
            bool existTtr = planSheet.CharList.Exists(a => a.Ttr);
            var flowSheet = new SubFlowSheet("Flow_CZ_" + planSheet.SheetName);
            var ttrHardIpFlow = new SubFlowSheet("Flow_TTR_CZ_" + planSheet.SheetName);
            var htolHardIpFlow = new SubFlowSheet("Flow_HTOL_CZ_" + planSheet.SheetName);
            SubFlowSheet tempFlowSheet = HardIpTempFlowGenerator.GetHardIpTempSubFlow(planSheet, ProdProg.AllTestInstancesDict);

            foreach (FlowRow flowStep in tempFlowSheet.Rows)
            {
                ProcessTempFlowStep(flowStep, planSheet, flowSheet, ttrHardIpFlow, htolHardIpFlow);
            }

            // not exist in flow but in prod inst sheet
            foreach (CharPlanItem charRow in planSheet.CharList.Where(row => !row.InProgFlow && row.InProgInstance).ToList())
            {
                List<FlowRow> resultFlowRows = GetHipFlowRows(charRow, TestNumMain.GetTestNum());

                if (charRow.InstanceName.Contains("SweepVoltage"))
                {
                    resultFlowRows = GenerateShmooForLoop(charRow, resultFlowRows);
                }

                if (_HasSweepCode(charRow))
                {
                    resultFlowRows.Insert(0, GenSweepCodeForRow(charRow));
                    resultFlowRows.Add(new FlowRow { Opcode = "next" });
                }

                AppendToHipFlows(resultFlowRows, charRow.Ttr, charRow.Htol, flowSheet, ttrHardIpFlow, htolHardIpFlow);
            }

            flowSheet.AddRow(new FlowRow { Opcode = "return" });

            if (existTtr)
            {
                ttrHardIpFlow.AddRow(new FlowRow { Opcode = "return" });
            }

            if (existHtol)
            {
                htolHardIpFlow.AddRow(new FlowRow { Opcode = "return" });
            }

            // add blank row for flow sheet
            flowSheet.AddRow(new FlowRow());
            ttrHardIpFlow.AddRow(new FlowRow());
            htolHardIpFlow.AddRow(new FlowRow());

            // neither exist in flow sheet, nor in instance sheet
            foreach (CharPlanItem charItem in planSheet.CharList.Where(row => !row.InProgFlow && !row.InProgInstance).ToList())
            {
                List<FlowRow> resultFlowRows = GetHipFlowRows(charItem, TestNumMain.GetTestNum());
                AppendToHipFlows(resultFlowRows, charItem.Ttr, charItem.Htol, flowSheet, ttrHardIpFlow, htolHardIpFlow);
            }

            WriteHipFlowOutputs(flowSheet, ttrHardIpFlow, htolHardIpFlow, existTtr, existHtol);

            return flowSheet.Name;
        }

        private void ProcessTempFlowStep(FlowRow flowStep, CharPlanSheet planSheet, SubFlowSheet flowSheet, SubFlowSheet ttrHardIpFlow, SubFlowSheet htolHardIpFlow)
        {
            flowStep.TNum = "";
            // handle flowStep exists in production flow
            // directly push into flows sheet for non XX_XX_NV test_name
            if (!Regex.IsMatch(flowStep.Parameter, "_HV$|_LV$|_NV$"))
            {
                flowSheet.AddRow(flowStep);
                return;
            }

            string progVoltage = flowStep.Parameter.Substring(flowStep.Parameter.Length - 2, 2);
            string subInstanceName = flowStep.Parameter.Replace("_" + progVoltage, "").ToLower();
            if (!LocalSpecs.PatternDic.TryGetValue(subInstanceName, out string pattern))
            {
                //else
                //{
                //    flowStep.Env = "NotCharItem";
                //    flowSheet.AddRow(flowStep);
                //}
                return;
            }

            // select same payload1 and voltage
            var charRowList =
                planSheet.CharList.Where(a => a.Payload1.ToLower() == pattern
                    && string.Equals(a.Voltage, progVoltage, StringComparison.CurrentCultureIgnoreCase)
                   ).ToList();

            // update select field
            charRowList.ForEach(row => row.Select = true);

            // If no releated voltage in CharPlan, use the first matched pattern item
            if (charRowList.Count == 0 && planSheet.CharList.Exists(a => a.Payload1.ToLower() == pattern))
            {
                charRowList.AddRange(planSheet.CharList.Where(a => a.Payload1.ToLower() == pattern).ToList());
            }

            foreach (CharPlanItem charRow in charRowList)
            {
                if (!ProcessHipCharRow(charRow, planSheet, flowStep, flowSheet, ttrHardIpFlow, htolHardIpFlow))
                {
                    break;
                }
            }
        }

        // Returns false to signal "break out of outer foreach over charRowList" (for the
        // pre-existing _HandleSpecialHacInstance early-break behavior).
        private bool ProcessHipCharRow(CharPlanItem charRow, CharPlanSheet planSheet, FlowRow flowStep, SubFlowSheet flowSheet, SubFlowSheet ttrHardIpFlow, SubFlowSheet htolHardIpFlow)
        {
            charRow.InProgFlow = true;
            if (charRow.Visit || string.IsNullOrEmpty(charRow.InstanceName))
            {
                return true;
            }

            //Init Pattern set to Payload1 already
            //Do not run shmoo
            if (charRow.UsedInits.Count == 1)
            {
                if (charRow.Payload1.ToUpper() == charRow.UsedInits.FirstOrDefault().ToUpper())
                {
                    return true;
                }
            }

            charRow.Visit = true;

            var resultFlowRows = new List<FlowRow>();

            if (_HandleSpecialHacInstance(flowStep, flowSheet))
            {
                return false;
            }

            var initItems = planSheet.CharList.Where(p => p.RowNum == charRow.RowNum && p != charRow).ToList();
            initItems.ForEach(p => p.InProgFlow = true);
            resultFlowRows.AddRange(GetHipFlowRows(charRow, TestNumMain.GetTestNum(), initItems));

            //resultFlowRows.AddRange(GetHipFlowRows(charRow, TestNumMain.GetTestNum()));
            try
            {
                if (charRow.InstanceName.Contains("SweepVoltage"))
                {
                    resultFlowRows = GenerateShmooForLoop(charRow, resultFlowRows);
                }
            }
            catch (Exception)
            {
            }
            if (_HasSweepCode(charRow))
            {
                resultFlowRows.Insert(0, GenSweepCodeForRow(charRow));
                resultFlowRows.Add(new FlowRow { Opcode = "next" });
            }

            AppendToHipFlows(resultFlowRows, charRow.Ttr, charRow.Htol, flowSheet, ttrHardIpFlow, htolHardIpFlow);
            return true;
        }

        private static void AppendToHipFlows(List<FlowRow> resultFlowRows, bool ttr, bool htol, SubFlowSheet flowSheet, SubFlowSheet ttrHardIpFlow, SubFlowSheet htolHardIpFlow)
        {
            flowSheet.Rows.AddRange(resultFlowRows);

            if (ttr)
            {
                ttrHardIpFlow.Rows.AddRange(resultFlowRows.ToList());
            }

            if (htol)
            {
                htolHardIpFlow.Rows.AddRange(resultFlowRows.ToList());
            }
        }

        private void WriteHipFlowOutputs(SubFlowSheet flowSheet, SubFlowSheet ttrHardIpFlow, SubFlowSheet htolHardIpFlow, bool existTtr, bool existHtol)
        {
            // export
            flowSheet.Write(Path.Combine(_outputFolder, flowSheet.Name + ".txt"), LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0");
            LocalSpecs.GenSheets.Add(flowSheet);

            if (existTtr)
            {
                ttrHardIpFlow.Write(Path.Combine(_outputFolder, ttrHardIpFlow.Name + ".txt"), LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0");
                LocalSpecs.GenSheets.Add(ttrHardIpFlow);
                LocalSpecs.HtolAndTtr[ConstData.Ttr].Add(htolHardIpFlow.Name);
            }

            if (existHtol)
            {
                htolHardIpFlow.Write(Path.Combine(_outputFolder, htolHardIpFlow.Name + ".txt"), LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0");
                LocalSpecs.GenSheets.Add(htolHardIpFlow);
                LocalSpecs.HtolAndTtr[ConstData.Htol].Add(htolHardIpFlow.Name);
            }
        }

        private bool _HandleSpecialHacInstance(FlowRow flowStep, SubFlowSheet flowSheet)
        {
            Dictionary<string, List<FlowRow>> flowItemGroups = LocalSpecs.TestProgram.FlowNameRowDict;

            string specialInstance =
                DataConvertor.SpecialHacInstanceRows.Keys.FirstOrDefault(
                    p => p.Equals(flowStep.Parameter, StringComparison.OrdinalIgnoreCase));

            if (specialInstance == null)
            {
                return false;
            }

            flowItemGroups.TryGetValue(specialInstance, out List<FlowRow> testSet);

            if (testSet == null)
            {
                return false;
            }

            testSet.ForEach(p => p.Parameter = DataConvertor.SpecialHacInstanceRows[specialInstance]);
            flowSheet.Rows.AddRange(testSet);
            return true;
        }

        private sealed class AtpgFlowSheets
        {
            public SubFlowSheet HvccFlow;
            public SubFlowSheet LvccFlow;
            public SubFlowSheet LvccWatCp1Stage;
            public SubFlowSheet LvccWatCp1AllStage;
            public SubFlowSheet LvccWatCp2Stage;
            public SubFlowSheet LvccWatCp2AllStage;
            public SubFlowSheet LvccWatFt1Stage;
            public SubFlowSheet LvccWatFt1AllStage;
            public SubFlowSheet LvccWatFt2Stage;
            public SubFlowSheet LvccWatFt2AllStage;
            public SubFlowSheet HtolHvccFlow;
            public SubFlowSheet HtolLvccFlow;
            public SubFlowSheet TtrHvccFlow;
            public SubFlowSheet TtrLvccFlow;
        }

        private static AtpgFlowSheets BuildAtpgFlowSheets(CharPlanSheet planSheet)
        {
            return new AtpgFlowSheets
            {
                HvccFlow = new SubFlowSheet("Flow_HVCC_CZ_" + planSheet.SheetName),
                LvccFlow = new SubFlowSheet("Flow_LVCC_CZ_" + planSheet.SheetName),
                LvccWatCp1Stage = new SubFlowSheet($"Flow_LVCC_CZ_CP1_WAT_{planSheet.SheetName}"),
                LvccWatCp1AllStage = new SubFlowSheet($"Flow_LVCC_CZ_CP1_All_{planSheet.SheetName}"),
                LvccWatCp2Stage = new SubFlowSheet($"Flow_LVCC_CZ_CP2_WAT_{planSheet.SheetName}"),
                LvccWatCp2AllStage = new SubFlowSheet($"Flow_LVCC_CZ_CP2_All_{planSheet.SheetName}"),
                LvccWatFt1Stage = new SubFlowSheet($"Flow_LVCC_CZ_FT1_WAT_{planSheet.SheetName}"),
                LvccWatFt1AllStage = new SubFlowSheet($"Flow_LVCC_CZ_FT1_All_{planSheet.SheetName}"),
                LvccWatFt2Stage = new SubFlowSheet($"Flow_LVCC_CZ_FT2_WAT_{planSheet.SheetName}"),
                LvccWatFt2AllStage = new SubFlowSheet($"Flow_LVCC_CZ_FT2_All_{planSheet.SheetName}"),
                HtolHvccFlow = new SubFlowSheet("Flow_HTOL_HVCC_CZ_" + planSheet.SheetName),
                HtolLvccFlow = new SubFlowSheet("Flow_HTOL_LVCC_CZ_" + planSheet.SheetName),
                TtrHvccFlow = new SubFlowSheet("Flow_TTR_HVCC_CZ_" + planSheet.SheetName),
                TtrLvccFlow = new SubFlowSheet("Flow_TTR_LVCC_CZ_" + planSheet.SheetName),
            };
        }

        private void _AtpgHandler(CharPlanSheet planSheet, ICollection<string> preScanList, ICollection<string> testFlowList)
        {
            bool existHvcc = planSheet.CharList.Exists(a => a.MeasType == "HVCC");
            bool existLvcc = planSheet.CharList.Exists(a => a.MeasType == "LVCC");


            bool existWat = planSheet.CharList.Any(a => !string.IsNullOrEmpty(a.StageCp1) || !string.IsNullOrEmpty(a.StageCp2) || !string.IsNullOrEmpty(a.StageFt1) || !string.IsNullOrEmpty(a.StageFt2));


            bool existHtol = planSheet.CharList.Exists(a => a.Htol);
            bool existTtr = planSheet.CharList.Exists(a => a.Ttr);

            AtpgFlowSheets sheets = BuildAtpgFlowSheets(planSheet);
            //var preScreenFlow = new SubFlowSheet("Flow_PreScan_" + planSheet.SheetName);
            SubFlowSheet lvccFlow = sheets.LvccFlow;
            var preScanInstanceName = new Dictionary<string, string>();

            //preScanList.Add(preScreenFlow.Name);
            WriteAtpgInitialHeaders(planSheet, sheets, existHvcc, existLvcc, existHtol, existTtr, existWat, testFlowList);

            string currentNWire = "";
            string nwireFlowName = "";
            foreach (CharPlanItem planItem in planSheet.CharList)
            {
                if ((!_genCharNoUsed && !planItem.Use) || (!_genPatNoUsed && !SearchInfo.CheckPatUsed(planItem)))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(planItem.ManualAc) &&
                    planItem.IsFreeRunClk &&
                    !planItem.ManualAc.Split(':')[1].Equals(currentNWire))
                {
                    currentNWire = planItem.ManualAc.Split(':')[1];
                    nwireFlowName = $"Flow_nWire_{currentNWire.Replace(".", "p")}MHz";
                }
                else
                {
                    nwireFlowName = "";
                }
                int testNum = TestNumMain.GetTestNum(false);
                var tmpEnableWd = new List<string>();
                if (planItem.CharShmooSetup != null)
                {
                    tmpEnableWd.Add("!ShmooOnly");
                }

                if (!string.IsNullOrEmpty(planItem.FailInfo) && !planItem.FailInfo.Equals("No", StringComparison.CurrentCultureIgnoreCase))
                {
                    tmpEnableWd.Add(planItem.FailInfo);
                }

                string nonHipEnableWord = string.Join("||", tmpEnableWd);

                string envWord = "";
                var envList = new List<string>();
                if (!string.IsNullOrEmpty(planItem.IsNeedMask))
                {
                    envList.Add(planItem.IsNeedMask);
                }
                //if (SearchInfo.CheckItemMask(charItem)) envWord += ",AutoMask";
                if (!planItem.Use)
                {
                    envList.Add("CharNoUse");
                }

                envWord = string.Join(",", envList);

                switch (planItem.MeasType)
                {
                    case "HVCC":
                        AtpgHvccCase(planItem, sheets, testNum, envWord, nonHipEnableWord, nwireFlowName, currentNWire);
                        break;

                    case "LVCC":
                        AtpgLvccCase(planItem, sheets, testNum, envWord, nonHipEnableWord, nwireFlowName, currentNWire);
                        break;

                }


                #region PreScreen

                // avoid preScan duplicate flow name
                if (preScanInstanceName.ContainsKey(planItem.TestInstanceName))
                {
                    continue;
                }

                var hvPreScreenStep = new FlowRow
                {
                    Opcode = "Test",
                    Parameter = planItem.TestInstanceName + "Pre_HV",
                    TNum = _TNumHandler(testNum.ToString(CultureInfo.InvariantCulture)),
                    Env = envWord,
                    SortFail = "9100",
                    BinFail = "5",
                    Result = "fail",
                    FailAction = ConstData.FailFlagDummy
                };

                preScanInstanceName.Add(planItem.TestInstanceName, "");
                #endregion
            }
            if (!string.IsNullOrEmpty(currentNWire))
            {
                _AddNwireSwitch(lvccFlow, ProdProg.NwireFlow.Name);
            }

            WriteAtpgFooters(planSheet, sheets, existHvcc, existLvcc, existHtol, existTtr, currentNWire);

            //var preScreenFooter = footerStep.DeepClone();
            //preScreenFooter.Parameter = planSheet.SheetName.Replace(" ", "_") + "_PreScreen_Footer";
            //preScreenFlow.AddRow(preScreenFooter);
            //preScreenFlow.AddRow(new FlowRow { Opcode = "return" });

            ////Write result
            //var preScreenFileName = Path.Combine(_outputFolder, preScreenFlow.Name + ".txt");
            //preScreenFlow.Write(preScreenFileName, LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0");
        }

        private static void WriteAtpgInitialHeaders(CharPlanSheet planSheet, AtpgFlowSheets sheets, bool existHvcc, bool existLvcc, bool existHtol, bool existTtr, bool existWat, ICollection<string> testFlowList)
        {
            //preScreenFlow.AddRow(new FlowRow { Opcode = "Test", Parameter = planSheet.SheetName + "_PreScreen_Header" });
            if (existHvcc)
            {
                var hvccFlowStep = new FlowRow { Opcode = "Print", Parameter = $"\"*print: {planSheet.SheetName.Replace(" ", "_") + "_HVCC_Header"} start*\"" };
                sheets.HvccFlow.AddRow(hvccFlowStep);
                testFlowList.Add(sheets.HvccFlow.Name);

                if (existHtol)
                {
                    sheets.HtolHvccFlow.AddRow(hvccFlowStep);
                    testFlowList.Add(sheets.HtolHvccFlow.Name);
                }
                if (existTtr)
                {
                    sheets.TtrHvccFlow.AddRow(hvccFlowStep);
                    testFlowList.Add(sheets.TtrHvccFlow.Name);
                }
            }
            if (existLvcc)
            {
                WriteLvccInitialHeader(planSheet, sheets, existHtol, existTtr, existWat, testFlowList);
            }
        }

        private static void WriteLvccInitialHeader(CharPlanSheet planSheet, AtpgFlowSheets sheets, bool existHtol, bool existTtr, bool existWat, ICollection<string> testFlowList)
        {
            var lvccFlowStep = new FlowRow { Opcode = "Print", Parameter = $"\"*print: {planSheet.SheetName.Replace(" ", "_") + "_LVCC_Header"} start*\"" };
            sheets.LvccFlow.AddRow(lvccFlowStep);
            testFlowList.Add(sheets.LvccFlow.Name);

            if (existHtol)
            {
                sheets.HtolLvccFlow.AddRow(lvccFlowStep);
                testFlowList.Add(sheets.HtolLvccFlow.Name);
            }
            if (existTtr)
            {
                sheets.TtrLvccFlow.AddRow(lvccFlowStep);
                testFlowList.Add(sheets.TtrLvccFlow.Name);
            }

            //WAT Die
            if (existWat)
            {
                sheets.LvccWatCp1Stage.AddRow(lvccFlowStep);
                sheets.LvccWatCp1AllStage.AddRow(lvccFlowStep);
                sheets.LvccWatCp2Stage.AddRow(lvccFlowStep);
                sheets.LvccWatCp2AllStage.AddRow(lvccFlowStep);
                sheets.LvccWatFt1Stage.AddRow(lvccFlowStep);
                sheets.LvccWatFt1AllStage.AddRow(lvccFlowStep);
                sheets.LvccWatFt2Stage.AddRow(lvccFlowStep);
                sheets.LvccWatFt2AllStage.AddRow(lvccFlowStep);

                testFlowList.Add(sheets.LvccWatCp1Stage.Name);
                testFlowList.Add(sheets.LvccWatCp1AllStage.Name);
                testFlowList.Add(sheets.LvccWatCp2Stage.Name);
                testFlowList.Add(sheets.LvccWatCp2AllStage.Name);
                testFlowList.Add(sheets.LvccWatFt1Stage.Name);
                testFlowList.Add(sheets.LvccWatFt1AllStage.Name);
                testFlowList.Add(sheets.LvccWatFt2Stage.Name);
                testFlowList.Add(sheets.LvccWatFt2AllStage.Name);
            }
        }

        private void WriteAtpgFooters(CharPlanSheet planSheet, AtpgFlowSheets sheets, bool existHvcc, bool existLvcc, bool existHtol, bool existTtr, string currentNWire)
        {
            // Generate footer step for each flow
            var footerStep = new FlowRow { Opcode = "Print" };
            if (existHvcc)
            {
                WriteHvccFooters(planSheet, sheets, footerStep, existHtol, existTtr, currentNWire);
            }

            if (existLvcc)
            {
                WriteLvccFooters(planSheet, sheets, footerStep, existHtol, existTtr);
            }
        }

        private void WriteHvccFooters(CharPlanSheet planSheet, AtpgFlowSheets sheets, FlowRow footerStep, bool existHtol, bool existTtr, string currentNWire)
        {
            if (!string.IsNullOrEmpty(currentNWire))
            {
                _AddNwireSwitch(sheets.HvccFlow, ProdProg.NwireFlow.Name);
            }
            FlowRow hvccFooterStep = footerStep.Copy();
            hvccFooterStep.Parameter = $"\"*print: {planSheet.SheetName.Replace(" ", "_") + "_HVCC_Footer"} end*\"";
            sheets.HvccFlow.AddRow(hvccFooterStep);
            sheets.HvccFlow.AddRow(new FlowRow { Opcode = "return" });

            //Write result
            string screenFileName = Path.Combine(_outputFolder, sheets.HvccFlow.Name + ".txt");
            sheets.HvccFlow.Write(screenFileName, LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0");
            LocalSpecs.GenSheets.Add(sheets.HvccFlow);

            if (existHtol)
            {
                WriteHvccFooterVariant(sheets.HtolHvccFlow, hvccFooterStep, currentNWire);
            }
            if (existTtr)
            {
                WriteHvccFooterVariant(sheets.TtrHvccFlow, hvccFooterStep, currentNWire);
            }
        }

        private void WriteHvccFooterVariant(SubFlowSheet flow, FlowRow hvccFooterStep, string currentNWire)
        {
            if (!string.IsNullOrEmpty(currentNWire))
            {
                _AddNwireSwitch(flow, ProdProg.NwireFlow.Name);
            }
            flow.AddRow(hvccFooterStep);
            flow.AddRow(new FlowRow { Opcode = "return" });
            //Write result
            string fileName = Path.Combine(_outputFolder, flow.Name + ".txt");
            flow.Write(fileName, LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0");
            LocalSpecs.GenSheets.Add(flow);
        }

        private void WriteLvccFooters(CharPlanSheet planSheet, AtpgFlowSheets sheets, FlowRow footerStep, bool existHtol, bool existTtr)
        {
            FlowRow lvccFooterStep = footerStep.Copy();
            lvccFooterStep.Parameter = $"\"*print: {planSheet.SheetName.Replace(" ", "_") + "_Lvcc_Footer"} end*\"";
            sheets.LvccFlow.AddRow(lvccFooterStep);
            sheets.LvccFlow.AddRow(new FlowRow { Opcode = "return" });

            //Write result
            string screenFileName = Path.Combine(_outputFolder, sheets.LvccFlow.Name + ".txt");
            sheets.LvccFlow.Write(screenFileName, LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0");
            LocalSpecs.GenSheets.Add(sheets.LvccFlow);

            if (existHtol)
            {
                WriteLvccFooterVariant(sheets.HtolLvccFlow, lvccFooterStep);
            }
            if (existTtr)
            {
                WriteLvccFooterVariant(sheets.TtrLvccFlow, lvccFooterStep);
            }

            WriteLvccWatFlows(sheets);
        }

        private void WriteLvccFooterVariant(SubFlowSheet flow, FlowRow lvccFooterStep)
        {
            flow.AddRow(lvccFooterStep);
            flow.AddRow(new FlowRow { Opcode = "return" });
            //Write result
            string fileName = Path.Combine(_outputFolder, flow.Name + ".txt");
            flow.Write(fileName, LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0");
            LocalSpecs.GenSheets.Add(flow);
        }

        private void WriteLvccWatFlows(AtpgFlowSheets sheets)
        {
            sheets.LvccWatCp1Stage.AddRow(new FlowRow { Opcode = "return" });
            sheets.LvccWatCp1AllStage.AddRow(new FlowRow { Opcode = "return" });
            sheets.LvccWatCp2Stage.AddRow(new FlowRow { Opcode = "return" });
            sheets.LvccWatCp2AllStage.AddRow(new FlowRow { Opcode = "return" });
            sheets.LvccWatFt1Stage.AddRow(new FlowRow { Opcode = "return" });
            sheets.LvccWatFt1AllStage.AddRow(new FlowRow { Opcode = "return" });
            sheets.LvccWatFt2Stage.AddRow(new FlowRow { Opcode = "return" });
            sheets.LvccWatFt2AllStage.AddRow(new FlowRow { Opcode = "return" });

            LocalSpecs.GenSheets.Add(sheets.LvccWatCp1Stage);
            LocalSpecs.GenSheets.Add(sheets.LvccWatCp1AllStage);
            LocalSpecs.GenSheets.Add(sheets.LvccWatCp2Stage);
            LocalSpecs.GenSheets.Add(sheets.LvccWatCp2AllStage);
            LocalSpecs.GenSheets.Add(sheets.LvccWatFt1Stage);
            LocalSpecs.GenSheets.Add(sheets.LvccWatFt1AllStage);
            LocalSpecs.GenSheets.Add(sheets.LvccWatFt2Stage);
            LocalSpecs.GenSheets.Add(sheets.LvccWatFt2AllStage);

            string ver = LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0";
            sheets.LvccWatCp1Stage.Write(Path.Combine(_outputFolder, sheets.LvccWatCp1Stage.Name + ".txt"), ver);
            sheets.LvccWatCp1AllStage.Write(Path.Combine(_outputFolder, sheets.LvccWatCp1AllStage.Name + ".txt"), ver);
            sheets.LvccWatCp2Stage.Write(Path.Combine(_outputFolder, sheets.LvccWatCp2Stage.Name + ".txt"), ver);
            sheets.LvccWatCp2AllStage.Write(Path.Combine(_outputFolder, sheets.LvccWatCp2AllStage.Name + ".txt"), ver);
            sheets.LvccWatFt1Stage.Write(Path.Combine(_outputFolder, sheets.LvccWatFt1Stage.Name + ".txt"), ver);
            sheets.LvccWatFt1AllStage.Write(Path.Combine(_outputFolder, sheets.LvccWatFt1AllStage.Name + ".txt"), ver);
            sheets.LvccWatFt2Stage.Write(Path.Combine(_outputFolder, sheets.LvccWatFt2Stage.Name + ".txt"), ver);
            sheets.LvccWatFt2AllStage.Write(Path.Combine(_outputFolder, sheets.LvccWatFt2AllStage.Name + ".txt"), ver);
        }

        private void AtpgHvccCase(CharPlanItem planItem, AtpgFlowSheets sheets, int testNum, string envWord, string nonHipEnableWord, string nwireFlowName, string currentNWire)
        {
            //Write nwire reset flow if needed
            if (!string.IsNullOrEmpty(nwireFlowName))
            {
                AddNwireSwitchVariants(sheets.HvccFlow, sheets.HtolHvccFlow, sheets.TtrHvccFlow, nwireFlowName, planItem.Htol, planItem.Ttr);
            }
            // write Init Variable
            if (planItem.CharShmooSetup != null && LocalSpecs.GenAssignSiteVar)
            {
                AddInitFlagVariants(sheets.HvccFlow, sheets.HtolHvccFlow, sheets.TtrHvccFlow, planItem.Htol, planItem.Ttr);
            }

            FlowRow hvccFlowStep = BuildTestFlowStep(planItem, envWord, nonHipEnableWord, testNum);
            WriteHvccTestSteps(sheets, hvccFlowStep, planItem, testNum, nonHipEnableWord);
            WriteBinTableIfFailFlag(sheets.HvccFlow, planItem);

            if (planItem.CharShmooSetup != null)
            {
                WriteHvccShmoo(sheets, planItem, testNum, envWord);
            }

            if (!string.IsNullOrEmpty(planItem.AdaptiveCooling) && ProdProg.TmpsFlow != null && _genTmpsOnFlow)
            {
                _AddTMPSFlows(sheets.HvccFlow, planItem, testNum, envWord, currentNWire);
            }
        }

        private void AddNwireSwitchVariants(SubFlowSheet mainFlow, SubFlowSheet htolFlow, SubFlowSheet ttrFlow, string nwireFlowName, bool htol, bool ttr)
        {
            _AddNwireSwitch(mainFlow, nwireFlowName);
            if (htol)
            {
                _AddNwireSwitch(htolFlow, nwireFlowName);
            }

            if (ttr)
            {
                _AddNwireSwitch(ttrFlow, nwireFlowName);
            }
        }

        private static void AddInitFlagVariants(SubFlowSheet mainFlow, SubFlowSheet htolFlow, SubFlowSheet ttrFlow, bool htol, bool ttr)
        {
            var initFlowStep = new FlowRow
            {
                Opcode = ConstData.FlagInitOpCode,
                Parameter = ConstData.FlagInit
            };
            mainFlow.AddRow(initFlowStep);
            if (htol)
            {
                htolFlow.AddRow(initFlowStep);
            }

            if (ttr)
            {
                ttrFlow.AddRow(initFlowStep);
            }
        }

        private FlowRow BuildTestFlowStep(CharPlanItem planItem, string envWord, string nonHipEnableWord, int testNum)
        {
            var step = new FlowRow
            {
                Opcode = "Test",
                Parameter = planItem.InstanceName,
                Enable = !string.IsNullOrEmpty(planItem.EnableWord) ? planItem.EnableWord : nonHipEnableWord,
                Env = envWord,
                //FailAction = LocalSpecs.GenAssignSiteVar?ConstData.FailFlagSet:"",
                FailAction = !string.IsNullOrEmpty(planItem.FailFlag) ? planItem.FailFlag : ConstData.FailFlagDummy,
                TNum = _TNumHandler((testNum + 60).ToString(CultureInfo.InvariantCulture))
            };
            if (!string.IsNullOrEmpty(planItem.SiteFlag))
            {
                step.DeviceCondition = "Flag-True";
                step.DeviceName = planItem.SiteFlag;
            }
            return step;
        }

        private void WriteHvccTestSteps(AtpgFlowSheets sheets, FlowRow hvccFlowStep, CharPlanItem planItem, int testNum, string nonHipEnableWord)
        {
            AddRtosBootRow(sheets.HvccFlow, planItem.IsUseRtosCmd, nonHipEnableWord);
            sheets.HvccFlow.AddRow(hvccFlowStep);
            if (planItem.Htol)
            {
                FlowRow htolStep = hvccFlowStep.Copy();
                htolStep.TNum = _TNumHandler((testNum + 100).ToString(CultureInfo.InvariantCulture));
                AddRtosBootRow(sheets.HtolHvccFlow, planItem.IsUseRtosCmd, nonHipEnableWord);
                sheets.HtolHvccFlow.AddRow(htolStep);
            }
            if (planItem.Ttr)
            {
                FlowRow ttrStep = hvccFlowStep.Copy();
                ttrStep.TNum = _TNumHandler((testNum + 150).ToString(CultureInfo.InvariantCulture));
                AddRtosBootRow(sheets.TtrHvccFlow, planItem.IsUseRtosCmd, nonHipEnableWord);
                sheets.TtrHvccFlow.AddRow(ttrStep);
            }
        }

        private static void WriteBinTableIfFailFlag(SubFlowSheet flow, CharPlanItem planItem)
        {
            if (!string.IsNullOrEmpty(planItem.FailFlag))
            {
                var binRow = new FlowRow { Opcode = "BinTable", Parameter = Regex.Replace(planItem.FailFlag, "^F_", "Bin_", RegexOptions.IgnoreCase) };
                flow.AddRow(binRow);
            }
        }

        private void WriteHvccShmoo(AtpgFlowSheets sheets, CharPlanItem planItem, int testNum, string envWord)
        {
            var shmooFlowStep = new FlowRow
            {
                Opcode = "characterize",
                Parameter = planItem.InstanceName + " " + planItem.CharShmooSetup.ShmooSetupName,
                TNum = _TNumHandler((testNum + 90).ToString(CultureInfo.InvariantCulture)),
                Enable = "!TestOnly",
                Env = envWord,
                FailAction = ConstData.FailFlagDummy,
                DeviceCondition = LocalSpecs.GenAssignSiteVar ? ConstData.FlagCondOpCode : "",
                DeviceName = LocalSpecs.GenAssignSiteVar ? ConstData.FlagCond : "",
            };
            sheets.HvccFlow.AddRow(shmooFlowStep);
            sheets.HvccFlow.AddRow(_binShmooAlarmFlowRow);
            if (planItem.Htol)
            {
                FlowRow htolShmooStep = shmooFlowStep.Copy();
                htolShmooStep.TNum = _TNumHandler((testNum + 130).ToString(CultureInfo.InvariantCulture));
                sheets.HtolHvccFlow.AddRow(htolShmooStep);
                sheets.HtolHvccFlow.AddRow(_binShmooAlarmFlowRow);
            }
            if (planItem.Ttr)
            {
                FlowRow ttrShmooStep = shmooFlowStep.Copy();
                ttrShmooStep.TNum = _TNumHandler((testNum + 180).ToString(CultureInfo.InvariantCulture));
                sheets.TtrHvccFlow.AddRow(ttrShmooStep);
                sheets.TtrHvccFlow.AddRow(_binShmooAlarmFlowRow);
            }
        }

        private void AtpgLvccCase(CharPlanItem planItem, AtpgFlowSheets sheets, int testNum, string envWord, string nonHipEnableWord, string nwireFlowName, string currentNWire)
        {
            if (!string.IsNullOrEmpty(nwireFlowName))
            {
                AddNwireSwitchVariants(sheets.LvccFlow, sheets.HtolLvccFlow, sheets.TtrLvccFlow, nwireFlowName, planItem.Htol, planItem.Ttr);
            }

            if (planItem.CharShmooSetup != null && LocalSpecs.GenAssignSiteVar)
            {
                AddInitFlagVariants(sheets.LvccFlow, sheets.HtolLvccFlow, sheets.TtrLvccFlow, planItem.Htol, planItem.Ttr);
            }

            FlowRow lvccFlowStep = BuildTestFlowStep(planItem, envWord, nonHipEnableWord, testNum);
            WriteLvccTestSteps(sheets, lvccFlowStep, planItem, testNum, nonHipEnableWord);
            WriteLvccWatStagesNonShmoo(sheets, lvccFlowStep, planItem, testNum, nonHipEnableWord);
            WriteBinTableIfFailFlag(sheets.LvccFlow, planItem);

            if (planItem.CharShmooSetup != null)
            {
                WriteLvccShmoo(sheets, planItem, testNum, envWord, nonHipEnableWord);
            }

            if (!string.IsNullOrEmpty(planItem.AdaptiveCooling) && ProdProg.TmpsFlow != null && _genTmpsOnFlow)
            {
                _AddTMPSFlows(sheets.LvccFlow, planItem, testNum, envWord, currentNWire);
            }
        }

        private void WriteLvccTestSteps(AtpgFlowSheets sheets, FlowRow lvccFlowStep, CharPlanItem planItem, int testNum, string nonHipEnableWord)
        {
            if (planItem.IsUseRtosCmd)
            {
                sheets.LvccFlow.AddRow
                    (new FlowRow { Opcode = "Test", Parameter = "RTOS_Boot", Enable = nonHipEnableWord });
            }

            sheets.LvccFlow.AddRow(lvccFlowStep);
            if (planItem.Htol)
            {
                FlowRow htolStep = lvccFlowStep.Copy();
                htolStep.TNum = _TNumHandler((testNum + 100).ToString(CultureInfo.InvariantCulture));
                AddRtosBootRow(sheets.HtolLvccFlow, planItem.IsUseRtosCmd, nonHipEnableWord);
                sheets.HtolLvccFlow.AddRow(htolStep);
            }
            if (planItem.Ttr)
            {
                FlowRow ttrStep = lvccFlowStep.Copy();
                ttrStep.TNum = _TNumHandler((testNum + 150).ToString(CultureInfo.InvariantCulture));
                AddRtosBootRow(sheets.TtrLvccFlow, planItem.IsUseRtosCmd, nonHipEnableWord);
                sheets.TtrLvccFlow.AddRow(ttrStep);
            }
        }

        // Switch WAT Die info to add row (non-shmoo). NOTE: pre-existing logic preserved as-is including
        // a suspected bug in the StageFt1 block where the non-All branch writes flowClone to the
        // AllStage flow instead of the Stage flow (preserved via explicit elseFlow parameter at call site).
        private void WriteLvccWatStagesNonShmoo(AtpgFlowSheets sheets, FlowRow lvccFlowStep, CharPlanItem planItem, int testNum, string nonHipEnableWord)
        {
            // StageCp1: correct mapping
            WriteLvccWatStage(planItem.StageCp1, planItem.StageCp1,
                allRtosFlow: sheets.LvccWatCp1AllStage, allFlow: sheets.LvccWatCp1AllStage,
                elseRtosFlow: sheets.LvccWatCp1Stage, elseFlow: sheets.LvccWatCp1Stage,
                baseStep: lvccFlowStep, testNum: testNum,
                isUseRtosCmd: planItem.IsUseRtosCmd, nonHipEnableWord: nonHipEnableWord);

            // StageCp2: correct mapping
            WriteLvccWatStage(planItem.StageCp2, planItem.StageCp2,
                allRtosFlow: sheets.LvccWatCp2AllStage, allFlow: sheets.LvccWatCp2AllStage,
                elseRtosFlow: sheets.LvccWatCp2Stage, elseFlow: sheets.LvccWatCp2Stage,
                baseStep: lvccFlowStep, testNum: testNum,
                isUseRtosCmd: planItem.IsUseRtosCmd, nonHipEnableWord: nonHipEnableWord);

            // StageFt1: BUG-PRESERVE — original code at the All-branch RTOS step wrote to LvccWatCp1AllStage
            // (likely should be LvccWatFt1AllStage), and the else-branch wrote flowClone to LvccWatFt1AllStage
            // (likely should be LvccWatFt1Stage). Explicit params preserve both anomalies.
            WriteLvccWatStage(planItem.StageFt1, planItem.StageFt1,
                allRtosFlow: sheets.LvccWatCp1AllStage, allFlow: sheets.LvccWatFt1AllStage,
                elseRtosFlow: sheets.LvccWatFt1Stage, elseFlow: sheets.LvccWatFt1AllStage,
                baseStep: lvccFlowStep, testNum: testNum,
                isUseRtosCmd: planItem.IsUseRtosCmd, nonHipEnableWord: nonHipEnableWord);

            // StageFt2: correct mapping
            WriteLvccWatStage(planItem.StageFt2, planItem.StageFt2,
                allRtosFlow: sheets.LvccWatFt2AllStage, allFlow: sheets.LvccWatFt2AllStage,
                elseRtosFlow: sheets.LvccWatFt2Stage, elseFlow: sheets.LvccWatFt2Stage,
                baseStep: lvccFlowStep, testNum: testNum,
                isUseRtosCmd: planItem.IsUseRtosCmd, nonHipEnableWord: nonHipEnableWord);
        }

        // Switch WAT Die info to add row (shmoo). NOTE: pre-existing logic preserved as-is including
        // multiple suspected bugs where StageCp2/StageFt1/StageFt2 blocks check planItem.StageCp1 for "ALL"
        // (likely should be StageCp2/StageFt1/StageFt2 respectively). Explicit params preserve the bugs.
        private void WriteLvccShmooWatStages(AtpgFlowSheets sheets, FlowRow shmooFlowStep, CharPlanItem planItem, int testNum, string nonHipEnableWord)
        {
            // StageCp1: correct mapping
            WriteLvccWatStage(planItem.StageCp1, planItem.StageCp1,
                allRtosFlow: sheets.LvccWatCp1AllStage, allFlow: sheets.LvccWatCp1AllStage,
                elseRtosFlow: sheets.LvccWatCp1Stage, elseFlow: sheets.LvccWatCp1Stage,
                baseStep: shmooFlowStep, testNum: testNum,
                isUseRtosCmd: planItem.IsUseRtosCmd, nonHipEnableWord: nonHipEnableWord);

            // StageCp2: BUG-PRESERVE — original checks StageCp1 for "ALL" instead of StageCp2.
            WriteLvccWatStage(planItem.StageCp2, planItem.StageCp1,
                allRtosFlow: sheets.LvccWatCp2AllStage, allFlow: sheets.LvccWatCp2AllStage,
                elseRtosFlow: sheets.LvccWatCp2Stage, elseFlow: sheets.LvccWatCp2Stage,
                baseStep: shmooFlowStep, testNum: testNum,
                isUseRtosCmd: planItem.IsUseRtosCmd, nonHipEnableWord: nonHipEnableWord);

            // StageFt1: BUG-PRESERVE — original checks StageCp1 for "ALL" instead of StageFt1; All-branch
            // RTOS step writes to LvccWatCp1AllStage; else-branch flowClone writes to LvccWatFt1AllStage.
            WriteLvccWatStage(planItem.StageFt1, planItem.StageCp1,
                allRtosFlow: sheets.LvccWatCp1AllStage, allFlow: sheets.LvccWatFt1AllStage,
                elseRtosFlow: sheets.LvccWatFt1Stage, elseFlow: sheets.LvccWatFt1AllStage,
                baseStep: shmooFlowStep, testNum: testNum,
                isUseRtosCmd: planItem.IsUseRtosCmd, nonHipEnableWord: nonHipEnableWord);

            // StageFt2: BUG-PRESERVE — original checks StageCp1 for "ALL" instead of StageFt2.
            WriteLvccWatStage(planItem.StageFt2, planItem.StageCp1,
                allRtosFlow: sheets.LvccWatFt2AllStage, allFlow: sheets.LvccWatFt2AllStage,
                elseRtosFlow: sheets.LvccWatFt2Stage, elseFlow: sheets.LvccWatFt2Stage,
                baseStep: shmooFlowStep, testNum: testNum,
                isUseRtosCmd: planItem.IsUseRtosCmd, nonHipEnableWord: nonHipEnableWord);
        }

        private void WriteLvccWatStage(string presenceStr, string allCheckStr,
            SubFlowSheet allRtosFlow, SubFlowSheet allFlow,
            SubFlowSheet elseRtosFlow, SubFlowSheet elseFlow,
            FlowRow baseStep, int testNum, bool isUseRtosCmd, string nonHipEnableWord)
        {
            if (string.IsNullOrEmpty(presenceStr))
            {
                return;
            }
            FlowRow flowClone = baseStep.Copy();
            flowClone.TNum = _TNumHandler((testNum + 150).ToString(CultureInfo.InvariantCulture));

            if (allCheckStr.Equals("ALL", StringComparison.CurrentCultureIgnoreCase))
            {
                AddRtosBootRow(allRtosFlow, isUseRtosCmd, nonHipEnableWord);
                allFlow.AddRow(flowClone);
            }
            else
            {
                AddRtosBootRow(elseRtosFlow, isUseRtosCmd, nonHipEnableWord);
                elseFlow.AddRow(flowClone);
            }
        }

        private static void AddRtosBootRow(SubFlowSheet flow, bool isUseRtosCmd, string nonHipEnableWord)
        {
            if (!isUseRtosCmd)
            {
                return;
            }
            flow.AddRow(new FlowRow { Opcode = "Test", Parameter = "RTOS_Boot", Enable = nonHipEnableWord });
        }

        private void WriteLvccShmoo(AtpgFlowSheets sheets, CharPlanItem planItem, int testNum, string envWord, string nonHipEnableWord)
        {
            var shmooFlowStep = new FlowRow
            {
                Opcode = "characterize",
                Parameter = planItem.InstanceName + " " +
                            planItem.CharShmooSetup.ShmooSetupName,
                Enable = "!TestOnly",
                Env = envWord,
                FailAction = ConstData.FailFlagDummy,
                DeviceCondition = LocalSpecs.GenAssignSiteVar ? ConstData.FlagCondOpCode : "",
                DeviceName = LocalSpecs.GenAssignSiteVar ? ConstData.FlagCond : ""
            };
            sheets.LvccFlow.AddRow(shmooFlowStep);
            sheets.LvccFlow.AddRow(_binShmooAlarmFlowRow);
            if (planItem.Htol)
            {
                FlowRow htolShmooStep = shmooFlowStep.Copy();
                htolShmooStep.TNum = _TNumHandler((testNum + 130).ToString(CultureInfo.InvariantCulture));
                sheets.HtolLvccFlow.AddRow(htolShmooStep);
                sheets.HtolLvccFlow.AddRow(_binShmooAlarmFlowRow);
            }
            if (planItem.Ttr)
            {
                FlowRow ttrShmooStep = shmooFlowStep.Copy();
                ttrShmooStep.TNum = _TNumHandler((testNum + 180).ToString(CultureInfo.InvariantCulture));
                sheets.TtrLvccFlow.AddRow(ttrShmooStep);
                sheets.TtrLvccFlow.AddRow(_binShmooAlarmFlowRow);
            }

            WriteLvccShmooWatStages(sheets, shmooFlowStep, planItem, testNum, nonHipEnableWord);
        }

        private void _AddTMPSFlows(SubFlowSheet flowSheet, CharPlanItem planItem, int testNum, string envWord, string currentNwire)
        {
            if (!string.IsNullOrEmpty(currentNwire))
            {
                var nwireDefaultFlow = new FlowRow
                {
                    Opcode = "call",
                    Parameter = "Flow_nWire_Default",
                    Enable = Regex.IsMatch(planItem.Environment, "wait", RegexOptions.IgnoreCase) ? "!Wait" : "!Char_TMPS",
                    Env = envWord,
                };
                flowSheet.AddRow(nwireDefaultFlow);
            }
            //string parameter;
            //string opCode;
            //if (DataConvertor.TryGetAdaptiveCoolingTmpsFlow(job, _flowTmpsName, out string flowName))
            //{
            //    parameter = flowName;
            //}
            //else
            //{
            //    parameter = _flowTmpsName;
            //}
            //var coolingFlowStep = new FlowRow
            //{
            //    Opcode = opCode,
            //    Parameter = Regex.IsMatch(planItem.Environment, "wait", RegexOptions.IgnoreCase) ? "Flow_Wait" : parameter,
            //    TNum = _TNumHandler((testNum + 190).ToString(CultureInfo.InvariantCulture)),
            //    Enable = Regex.IsMatch(planItem.Environment, "wait", RegexOptions.IgnoreCase) ? "!Wait" : "!Char_TMPS",
            //    Env = envWord,
            //    FailAction = ConstData.FailFlagDummy,
            //    DeviceCondition = LocalSpecs.GenAssignSiteVar ? ConstData.FlagCondOpCode : "",
            //    DeviceName = LocalSpecs.GenAssignSiteVar ? ConstData.FlagCond : "",
            //};

            var coolingFlowStep = new FlowRow
            {
                Opcode = "call",
                Parameter = Regex.IsMatch(planItem.Environment, "wait", RegexOptions.IgnoreCase) ? "Flow_Wait" :
                                    DataConvertor.GetCharAdaptiveCoolingTmpsFlow(planItem.AdaptiveCooling, _flowTmpsName, _job),
                TNum = _TNumHandler((testNum + 190).ToString(CultureInfo.InvariantCulture)),
                Enable = Regex.IsMatch(planItem.Environment, "wait", RegexOptions.IgnoreCase) ? "!Wait" : "!Char_TMPS",
                Env = envWord,
                FailAction = ConstData.FailFlagDummy,
                DeviceCondition = LocalSpecs.GenAssignSiteVar ? ConstData.FlagCondOpCode : "",
                DeviceName = LocalSpecs.GenAssignSiteVar ? ConstData.FlagCond : "",
            };
            flowSheet.AddRow(coolingFlowStep);

            if (!string.IsNullOrEmpty(currentNwire))
            {
                var nwireCurrentFlow = new FlowRow
                {
                    Opcode = "call",
                    Parameter = $"Flow_nWire_{currentNwire.Replace(".", "p")}MHz",
                    Enable = Regex.IsMatch(planItem.Environment, "wait", RegexOptions.IgnoreCase) ? "!Wait" : "!Char_TMPS",
                    Env = envWord,
                };
                flowSheet.AddRow(nwireCurrentFlow);
            }
        }

        private void _AddNwireSwitch(SubFlowSheet flowSheet, string nwireFlowName)
        {
            bool? removeLast = flowSheet.Rows.LastOrDefault()?
                .Parameter.StartsWith("Flow_nWire_", StringComparison.OrdinalIgnoreCase);
            if ((bool)removeLast)
            {
                flowSheet.Rows.RemoveAt(flowSheet.Rows.Count - 1);
            }
            flowSheet.AddRow(new FlowRow { Opcode = "Test", Parameter = "Relay_ON_Default" });
            flowSheet.AddRow(new FlowRow { Opcode = "call", Parameter = nwireFlowName });
        }

        private void AddShmooAbnormalBinOutFlowRows(SubFlowSheet mainChar)
        {
            //hardcode to add test instance 2017/6/26 
            mainChar.AddRow(new FlowRow
            {
                FailAction = "F_Check_Shmoo_Hole_Ratio_Within_Spec",
                Opcode = "test",
                Parameter = "Check_Shmoo_Hole_Ratio_Within_Spec"
            });

            mainChar.AddRow(new FlowRow
            {
                FailAction = "F_Check_Shmoo_Allfail_Ratio_Within_Spec",
                Opcode = "test",
                Parameter = "Check_Shmoo_Allfail_Ratio_Within_Spec"
            });

            mainChar.AddRow(new FlowRow
            {
                FailAction = "F_Check_Shmoo_Alarm_Ratio_Within_Spec",
                Opcode = "test",
                Parameter = "Check_Shmoo_Alarm_Ratio_Within_Spec"
            });

            mainChar.AddRow(new FlowRow
            {
                Opcode = "test",
                Parameter = ConstData.DisableCzMode
            });
            //mainChar.AddRow(new FlowRow { Opcode = "Test", Parameter = ConstData.DisableCzMode });

            // hardcode bintable 2017/6/27 for shmoo
            mainChar.AddRow(new FlowRow
            {
                Enable = "Error_Code_Bin_Out",
                Opcode = "BinTable",
                Parameter = "Bin_Char_Shmoo_Hole_Alarm_Allfail"
            });

            mainChar.AddRow(new FlowRow
            {
                Enable = "Error_Code_Bin_Out",
                Opcode = "BinTable",
                Parameter = "Bin_Char_shmoo_hole_alarm"
            });

            mainChar.AddRow(new FlowRow
            {
                Enable = "Error_Code_Bin_Out",
                Opcode = "BinTable",
                Parameter = "Bin_Char_shmoo_hole_allfail"
            });

            mainChar.AddRow(new FlowRow
            {
                Enable = "Error_Code_Bin_Out",
                Opcode = "BinTable",
                Parameter = "Bin_Char_shmoo_alarm_allfail"
            });

            mainChar.AddRow(new FlowRow
            {
                Enable = "Error_Code_Bin_Out",
                Opcode = "BinTable",
                Parameter = "Bin_Char_shmoo_allfail"
            });

            mainChar.AddRow(new FlowRow
            {
                Enable = "Error_Code_Bin_Out",
                Opcode = "BinTable",
                Parameter = "Bin_Char_shmoo_alarm"
            });

            mainChar.AddRow(new FlowRow
            {
                Enable = "Error_Code_Bin_Out",
                Opcode = "BinTable",
                Parameter = "Bin_Char_shmoo_hole"
            });
        }



        private List<FlowRow> GetHipFlowRows(CharPlanItem charItem, int testNum, List<CharPlanItem> initItems)
        {
            var resultFlowRows = new List<FlowRow>();
            var envList = new List<string>();
            if (!charItem.InProgInstance)
            {
                envList.Add("NotFoundInProd");
            }

            if (!string.IsNullOrEmpty(charItem.IsNeedMask))
            {
                envList.Add(charItem.IsNeedMask);
            }
            //if (SearchInfo.CheckItemMask(charItem)) envWord += ",AutoMask";
            if (!charItem.Use)
            {
                envList.Add("CharNoUse");
            }

            string envWord = string.Join(",", envList);

            foreach (CharPlanItem initItem in initItems)
            {
                _UpdateInitFlowRow(initItem, testNum, envWord, resultFlowRows);
            }

            switch (charItem.MeasType)
            {
                case "HIO":
                    _UpdateHioFlowRows(charItem, testNum, envWord, resultFlowRows);
                    break;

                case "HF":
                    _UpdateHfFlowRows(charItem, testNum, envWord, resultFlowRows);
                    break;

                default:
                    _UpdateHacFlowRows(charItem, testNum, envWord, resultFlowRows);
                    break;
            }
            return resultFlowRows;
        }



        private List<FlowRow> GetHipFlowRows(CharPlanItem charItem, int testNum)
        {
            var result = new List<FlowRow>();
            var envList = new List<string>();
            if (!charItem.InProgInstance)
            {
                envList.Add("NotFoundInProd");
            }

            if (!string.IsNullOrEmpty(charItem.IsNeedMask))
            {
                envList.Add(charItem.IsNeedMask);
            }
            //if (SearchInfo.CheckItemMask(charItem)) envWord += ",AutoMask";
            if (!charItem.Use)
            {
                envList.Add("CharNoUse");
            }

            string envWord = string.Join(",", envList);

            switch (charItem.MeasType)
            {
                case "HIO":
                    _UpdateHioFlowRows(charItem, testNum, envWord, result);
                    break;

                case "HF":
                    _UpdateHfFlowRows(charItem, testNum, envWord, result);
                    break;

                default:
                    _UpdateHacFlowRows(charItem, testNum, envWord, result);
                    break;
            }
            return result;
        }

        private void _UpdateHioFlowRows(CharPlanItem charItem, int testNum, string envWord, ICollection<FlowRow> result)
        {
            // reset flag
            if (LocalSpecs.GenAssignSiteVar)
            {
                result.Add(new FlowRow
                {
                    Opcode = ConstData.FlagInitOpCode,
                    Parameter = ConstData.FlagInit,
                    Enable = "Enable_" + charItem.MeasType
                });
            }

            // test
            result.Add(new FlowRow
            {
                Opcode = "Test",
                Parameter = charItem.InstanceName,
                TNum = _TNumHandler(testNum.ToString(CultureInfo.InvariantCulture)),
                Enable = "Enable_" + charItem.MeasType + " && !ShmooOnly",
                FailAction = LocalSpecs.GenAssignSiteVar ? ConstData.FailFlagSet : ConstData.FailFlagDummy
            });

            if (charItem.CharShmooSetup == null)
            {
                return;
            }

            if (charItem.IsUseRtosCmd)
            {
                result.Add(new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "RTOS_Boot",
                    Enable = "Enable_" + charItem.MeasType,
                    DeviceCondition = LocalSpecs.GenAssignSiteVar ? ConstData.FlagCondOpCode : "",
                    DeviceName = LocalSpecs.GenAssignSiteVar ? ConstData.FlagCond : ""
                });
            }

            result.Add(new FlowRow
            {
                Opcode = "Test",
                Parameter = charItem.InstanceName + "_Char",
                Enable = "Enable_" + charItem.MeasType,
                TNum = _TNumHandler((testNum + 51).ToString(CultureInfo.InvariantCulture)),
                Env = envWord,
                DeviceCondition = LocalSpecs.GenAssignSiteVar ? ConstData.FlagCondOpCode : "",
                DeviceName = LocalSpecs.GenAssignSiteVar ? ConstData.FlagCond : ""
            });
        }

        private void _UpdateHfFlowRows(CharPlanItem charItem, int testNum, string envWord, ICollection<FlowRow> result)
        {
            // use-limits
            var useLimitRow = new List<FlowRow>();
            string useLimitKey = !string.IsNullOrEmpty(charItem.Voltage)
                ? charItem.MappingPatternSet.ToLower() + "_" + charItem.Voltage.ToLower()
                : charItem.MappingPatternSet.ToLower() + "_" + "nv";
            if (LocalSpecs.UseLimitDict.ContainsKey(useLimitKey))
            {
                var instanceUseLimit = new List<FlowRow>();
                string selectKey = LocalSpecs.UseLimitDict[useLimitKey].Keys.FirstOrDefault(x => x.ToLower().Split(':')[1].Contains("_" + charItem.Type.ToLower() + "_"));
                if (!string.IsNullOrEmpty(selectKey))
                {
                    instanceUseLimit = LocalSpecs.UseLimitDict[useLimitKey][selectKey];
                }
                else
                {
                    instanceUseLimit = LocalSpecs.UseLimitDict[useLimitKey].Values.FirstOrDefault();
                }

                useLimitRow.AddRange(instanceUseLimit.Select(useLimit => new FlowRow
                {
                    Job = useLimit.Job,
                    Opcode = "use-limit",
                    Parameter = charItem.InstanceName + " " + charItem.CharShmooSetup.ShmooSetupName,
                    Enable = "Enable_" + charItem.MeasType,
                    Env = envWord,
                    LoLim = useLimit.LoLim,
                    HiLim = useLimit.HiLim,
                    TName = useLimit.TName,
                    FailAction = ConstData.FailFlagDummy
                }));
            }

            // run test
            result.Add(new FlowRow
            {
                Opcode = "Test",
                Parameter = charItem.Burst.ToLower() != "yes" && charItem.Payload1.ToUpper().Contains("INIT") ? Regex.Match(charItem.TestInstanceName, @"(?<sheet>\w+)" + charItem.BlockName).Groups["sheet"] + charItem.Payload1 + "_" + charItem.Voltage : charItem.InstanceName,
                Enable = "Enable_" + charItem.MeasType + "&&" + "!ShmooOnly",
                TNum = _TNumHandler((testNum + 35).ToString(CultureInfo.InvariantCulture)),
                Env = envWord,
                FailAction = ConstData.FailFlagDummy
            });
            useLimitRow.ForEach(result.Add);

            // run shmoo
            result.Add(new FlowRow
            {
                Opcode = "characterize",
                Parameter = charItem.Burst.ToLower() != "yes" && charItem.Payload1.ToUpper().Contains("INIT") ? Regex.Match(charItem.TestInstanceName, @"(?<sheet>\w+)" + charItem.BlockName).Groups["sheet"] + charItem.Payload1 + "_" + charItem.Voltage : charItem.InstanceName + " " + charItem.CharShmooSetup.ShmooSetupName,
                Enable = "Enable_" + charItem.MeasType + "&&" + "!TestOnly",
                TNum = _TNumHandler((testNum + 51).ToString(CultureInfo.InvariantCulture)),
                Env = envWord,
                FailAction = ConstData.FailFlagDummy
            });
            if (!_ignoreHfLimits)
            {
                useLimitRow.ForEach(result.Add);
            }
            else
            {
                var modifiedUseLimit = new List<FlowRow>();
                useLimitRow.ForEach(x => modifiedUseLimit.Add(x.Copy()));
                modifiedUseLimit.ForEach(x => { x.HiLim = ""; x.LoLim = ""; });
                modifiedUseLimit.ForEach(result.Add);
            }
            result.Add(_binShmooAlarmFlowRow);
        }

        private void _UpdateHacFlowRows(CharPlanItem charItem, int testNum, string envWord, ICollection<FlowRow> result)
        {
            if (charItem.IsUseRtosCmd)
            {
                result.Add(new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "RTOS_Boot",
                    Enable = "Enable_" + charItem.MeasType,
                });
            }

            result.Add(new FlowRow
            {
                Opcode = "Test",
                Parameter = charItem.InstanceName,
                TNum = _TNumHandler(testNum.ToString(CultureInfo.InvariantCulture)),
                Enable = "Enable_" + charItem.MeasType,
                Env = envWord,
                FailAction = ConstData.FailFlagDummy
            });

            // generate use-limits for hardip
            string[] uslArray = charItem.IpUse4.Split(',');
            string[] lslArray = charItem.IpUse5.Split(',');
            string[] unitsArray = charItem.IpUse6.Split(',');
            string[] testNameArray = charItem.TestInstanceName.Split(',');

            // check all the arraies has the same length
            if (uslArray.Length != testNameArray.Length || lslArray.Length != testNameArray.Length ||
                unitsArray.Length != testNameArray.Length)
            {
                GeneralFunc.WriteMessage("Format error of usl,lsl,units. " + charItem.BlockName);
            }

            List<List<string>> limits = Collections.Zip(new List<string[]> { uslArray, lslArray, unitsArray, testNameArray });
            IEnumerable<List<string>> idsLimits = limits.Where(x => x[3].Split('_').Length >= 6);
            List<string>[] idsDcvsPinLimits = idsLimits.Where(x =>
            {
                string tNamePin = x[3].Split('_')[5].ToUpper();
                var pins = new List<string>();
                if (LocalSpecs.ProgInfo.PinGroupDic.TryGetValue(tNamePin, out PinGroup value))
                {
                    pins = value.PinList.Select(y => y.PinName.Replace("_", "").ToUpper()).ToList();
                }
                else if (LocalSpecs.ProgInfo.PinDic.TryGetValue(tNamePin, out string value1))
                {
                    pins = new List<string>() { value1 };
                }

                foreach (string pin in pins)
                {
                    string search = LocalSpecs.ProgInfo.PinTypeInChannelDic.Keys.FirstOrDefault(y => y.Replace("_", "") == tNamePin);
                    if (LocalSpecs.ProgInfo.PinTypeInChannelDic.TryGetValue(pin, out string value1))
                    {
                        if (value1.StartsWith("DCVS"))
                        {
                            return true;
                        }
                    }
                }
                return false;
            }).OrderBy(x => x[3].Split('_')[5]).ToArray();

            List<string>[] idsDcviPinLimits = idsLimits.Where(x =>
            {
                string tNamePin = x[3].Split('_')[5].ToUpper();
                var pins = new List<string>();
                if (LocalSpecs.ProgInfo.PinGroupDic.TryGetValue(tNamePin, out PinGroup value))
                {
                    pins = value.PinList.Select(y => y.PinName).ToList();
                }
                else if (LocalSpecs.ProgInfo.PinDic.TryGetValue(tNamePin, out string value1))
                {
                    pins = new List<string>() { value1 };
                }

                foreach (string pin in pins)
                {
                    string search = LocalSpecs.ProgInfo.PinTypeInChannelDic.Keys.FirstOrDefault(y => y.Equals(pin, StringComparison.OrdinalIgnoreCase));
                    if (!string.IsNullOrEmpty(search))
                    {
                        if (LocalSpecs.ProgInfo.PinTypeInChannelDic[search] == "DCVI")
                        {
                            return true;
                        }
                    }
                }
                return false;
            }).OrderBy(x => x[3].Split('_')[5]).ToArray();

            IEnumerable<List<string>> otherLimits = limits.Where(x => idsDcvsPinLimits.Any(y => y[3] != x[3]) && idsDcviPinLimits.Any(y => y[3] != x[3]));

            var allLimits = new List<List<string>>();
            allLimits.AddRange(idsDcvsPinLimits);
            allLimits.AddRange(idsDcviPinLimits);
            allLimits.AddRange(otherLimits);

            foreach (List<string> limit in allLimits)
            {
                List<string> pinLimit = SearchInfo.GetLimitsOfPinFromPinGroup(limit[3], _genCSharp);
                pinLimit.ForEach(limitName =>
                {
                    var useLimitRow = new FlowRow
                    {
                        Opcode = "use-limit",
                        Parameter = charItem.InstanceName,
                        TName = limitName,
                        HiLim = limit[0],
                        LoLim = limit[1],
                        FailAction = ConstData.FailFlagDummy
                    };
                    _UpdateUseLimitUnit(limit[2], useLimitRow);
                    result.Add(useLimitRow);
                });
            }
        }

        private void _UpdateInitFlowRow(CharPlanItem charItem, int testNum, string envWord, ICollection<FlowRow> result)
        {
            // use-limits
            //var useLimitRow = new List<FlowRow>();
            //var useLimitKey = !string.IsNullOrEmpty(charItem.Voltage)
            //    ? charItem.MappingPatternSet.ToLower() + "_" + charItem.Voltage.ToLower()
            //    : charItem.MappingPatternSet.ToLower() + "_" + "nv";

            // run test
            result.Add(new FlowRow
            {
                Opcode = "Test",
                Parameter = charItem.Burst.ToLower() != "yes" && charItem.Payload1.ToUpper().Contains("INIT") ? Regex.Match(charItem.TestInstanceName, @"(?<sheet>\w+)" + charItem.BlockName).Groups["sheet"] + charItem.Payload1 + "_" + charItem.Voltage : charItem.InstanceName,
                Enable = "Enable_" + charItem.MeasType,
                TNum = _TNumHandler((testNum + 35).ToString(CultureInfo.InvariantCulture)),
                Env = envWord,
                FailAction = ConstData.FailFlagDummy
            });

        }

        private void _UpdateUseLimitUnit(string unit, FlowRow useLimitRow)
        {
            if (unit == "")
            {
                return;
            }

            switch (unit.ToUpper())
            {
                case "OHMS":
                    useLimitRow.Units = "Ohm";
                    break;

                default:
                    useLimitRow.Units = unit;
                    break;
            }
        }

        private FlowRow GenSweepCodeForRow(CharPlanItem charitem)
        {
            /* generate sweep code row */
            int steps = (Convert.ToInt16(charitem.CharShmooSetup.ShmooPins[0].StopPoint) - Convert.ToInt16(charitem.CharShmooSetup.ShmooPins[0].StartPoint) + 1) / Convert.ToInt16(charitem.CharShmooSetup.ShmooPins[0].StepSize);
            return new FlowRow
            {
                Opcode = "for",
                Parameter = "SrcCodeIndx = 0; SrcCodeIndx < " + steps + "; SrcCodeIndx++"
            };
        }

        private bool _HasSweepCode(CharPlanItem charitem)
        {
            if (charitem.CharShmooSetup == null)
            {
                return false;
            }

            if (charitem.CharShmooSetup.ShmooPins == null)
            {
                return false;
            }

            return charitem.CharShmooSetup.ShmooPins.FirstOrDefault(s => s.ShmooType != null
                && Regex.IsMatch(s.SweepType, "SweepCode", RegexOptions.IgnoreCase)) != null;
        }

        private SubFlowSheet GenerateShmooSetupInit(SubFlowSheet mainChar)
        {
            /* generate the init flag for shmoo setup */
            var devCharSetup = new FlowRow { Opcode = "create-site-var", Parameter = "Flow_Shmoo_DevCharSetup" };
            var yStep = new FlowRow { Opcode = "create-site-var", Parameter = "Flow_Shmoo_Y_Step" };
            var xStep = new FlowRow { Opcode = "create-site-var", Parameter = "Flow_Shmoo_X_Step" };
            var y = new FlowRow { Opcode = "create-site-var", Parameter = "Flow_Shmoo_Y" };
            var x = new FlowRow { Opcode = "create-site-var", Parameter = "Flow_Shmoo_X" };
            mainChar.AddRow(devCharSetup);
            mainChar.AddRow(yStep);
            mainChar.AddRow(xStep);
            mainChar.AddRow(y);
            mainChar.AddRow(x);
            return mainChar;
        }

        private List<FlowRow> GenerateShmooForLoop(CharPlanItem charitem, List<FlowRow> charRow)
        {
            /* generate the init flag for shmoo setup */
            var outputCharRow = new List<FlowRow>();
            if (charitem.CharShmooSetup == null || charitem.CharShmooSetup.ShmooPins == null)
            {
                return charRow;
            }

            bool isZLoop = charitem.CharShmooSetup.ShmooPins.FindAll(s => s.ShmooType == "Z").Count > 0;
            bool isYLoop = charitem.CharShmooSetup.ShmooPins.FindAll(s => s.ShmooType == "Y").Count > 0;
            bool isXLoop = charitem.CharShmooSetup.ShmooPins.FindAll(s => s.ShmooType == "X").Count > 0;

            if (!isYLoop && !isXLoop)
            {
                return outputCharRow;
            }

            outputCharRow.Add(new FlowRow
            {
                Opcode = ConstData.FlagInitOpCode,
                Parameter = "Flow_Shmoo_DevCharSetup " + charitem.CharShmooSetup.ShmooSetupName
            });

            outputCharRow.Add(new FlowRow
            {
                Opcode = "Test",
                Parameter = "Setup_Flow_Shmoo"
            });
            if (isZLoop)
            {
                outputCharRow.Add(new FlowRow
                {
                    Opcode = "For",
                    Parameter = "Flow_Shmoo_Z=0;Flow_Shmoo_Z<=Flow_Shmoo_Z_Step;Flow_Shmoo_Z+=1"
                });
            }
            if (isYLoop)
            {
                outputCharRow.Add(new FlowRow
                {
                    Opcode = "For",
                    Parameter = "Flow_Shmoo_Y=0;Flow_Shmoo_Y<=Flow_Shmoo_Y_Step;Flow_Shmoo_Y+=1"
                });
            }

            if (isXLoop)
            {
                outputCharRow.Add(new FlowRow
                {
                    Opcode = "For",
                    Parameter = "Flow_Shmoo_X=0;Flow_Shmoo_X<=Flow_Shmoo_X_Step;Flow_Shmoo_X+=1"
                });
            }

            outputCharRow.AddRange(charRow);

            if (isXLoop)
            {
                outputCharRow.Add(new FlowRow { Opcode = "Next" });
            }

            if (isYLoop)
            {
                outputCharRow.Add(new FlowRow { Opcode = "Next" });
            }

            return outputCharRow;
        }

        private void _ProcessTMPS_Flow(List<CharPlanSheet> charPlanSheets)
        {
            if (ProdProg.TmpsFlow == null)
            {
                return;
            }

            string job = LocalSpecs.CurrentJob;
            // Search for creating extra TMPS subflow
            // Search AdaptiveCooling for creating extra TMPS subflow

            //var tmpsGroups = charPlanSheets.SelectMany(p => p.CharList)
            //    .Where(p => !DataConvertor.GetCharTMPSFlow(p.Environment, _flowTmpsName).Equals("Flow_" + _flowTmpsName, StringComparison.OrdinalIgnoreCase))
            //    .GroupBy(p => DataConvertor.GetCharTMPSFlow(p.Environment, _flowTmpsName)).ToDictionary(p => p.Key, p => p.ToList());

            var tmpsGroups = charPlanSheets.SelectMany(p => p.CharList)
                .Where(p => !DataConvertor.GetCharAdaptiveCoolingTmpsFlow(p.AdaptiveCooling, _flowTmpsName, _job).Equals("Flow_" + _flowTmpsName, StringComparison.OrdinalIgnoreCase))
                .GroupBy(p => DataConvertor.GetCharAdaptiveCoolingTmpsFlow(p.AdaptiveCooling, _flowTmpsName, _job)).ToDictionary(p => p.Key, p => p.ToList());


            foreach (KeyValuePair<string, List<CharPlanItem>> tmpsGroup in tmpsGroups)
            {
                var newTmpsFlow = new SubFlowSheet(tmpsGroup.Key);
                DataConvertor.GetTMPS_temperature(tmpsGroup.Value.First().Environment, out string lowLimit, out string highLimit);
                DataConvertor.GetAdaptiveCooling_temperature(_job, out lowLimit, out highLimit);
                foreach (FlowRow row in ProdProg.TmpsFlow.Rows)
                {
                    if (row.Units.Equals("C", StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(row.Job) ||
                            row.Job.Split(',').ToList().Exists(p => p.Equals(job, StringComparison.OrdinalIgnoreCase)))
                        {
                            FlowRow newrow = row.Copy();
                            newrow.LoLim = lowLimit;
                            newrow.HiLim = highLimit;
                            newTmpsFlow.AddRow(newrow);
                            continue;
                        }
                    }
                    if (row.Opcode.Contains("assign-integer"))
                    {
                        DataConvertor.GetAdaptiveCoolingCount(_job, out string cnt);
                        FlowRow newrow = row.Copy();
                        newrow.Parameter = $"SrcCodeIndx1 {cnt}";
                        newTmpsFlow.AddRow(newrow);
                        continue;

                    }

                    newTmpsFlow.AddRow(row);
                }
                string tmpsFileName = Path.Combine(_outputFolder, newTmpsFlow.Name + ".txt");
                newTmpsFlow.Write(tmpsFileName, LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0");
                LocalSpecs.GenSheets.Add(newTmpsFlow);
            }
        }

        private void _ProcessNwireFlow(List<CharPlanSheet> charPlanSheets)
        {
            if (ProdProg.NwireFlow == null)
            {
                return;
            }
            // Search for creating extra TMPS subflow
            var frcSet = charPlanSheets.SelectMany(p => p.CharList)
                .Where(p => !string.IsNullOrEmpty(p.FreeRunningClock)).Select(p => p.FreeRunningClock).Distinct().ToList();
            foreach (string frc in frcSet)
            {
                if (frc.Split(':').Length >= 2)
                {
                    string frcName = frc.Split(':')[0].ToLower();
                    if (decimal.TryParse(frc.Split(':')[1], out decimal frcFreq))
                    {
                        string mhz = (frcFreq / 1000000).ToString().Replace(".", "p");
                        var newNwireFlow = new SubFlowSheet($"{ProdProg.NwireFlow.Name}_{mhz}MHz".Replace("_Default", ""));
                        foreach (FlowRow row in ProdProg.NwireFlow.Rows)
                        {
                            if (row.Parameter.ToLower().Contains("freerunclk") && row.Parameter.ToLower().Replace("_", "").Contains(frcName))
                            {
                                FlowRow newRow = row.Copy();
                                newRow.Parameter = $"{newRow.Parameter}_{mhz}MHz";
                                newNwireFlow.AddRow(newRow);
                            }
                            else
                            {
                                newNwireFlow.AddRow(row);
                            }
                        }
                        string nWireFileName = Path.Combine(_outputFolder, newNwireFlow.Name + ".txt");
                        newNwireFlow.Write(nWireFileName, LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0");
                        LocalSpecs.GenSheets.Add(newNwireFlow);
                    }
                }

            }
        }

        private void _ProcessMainInitFlow()
        {
            SubFlowSheet mainInitflowSheet = LocalSpecs.TestProgram.FlowSheetsAll.FirstOrDefault(p => p.Name
                .Equals("Flow_Table_Main_Init_Flows", StringComparison.OrdinalIgnoreCase));
            if (mainInitflowSheet != null)
            {
                FlowRow onProgramStartedBinOutFunction = mainInitflowSheet.Rows.FirstOrDefault(x => !x.IsBackup &&
                                                                                      x.Opcode.Equals("Test", StringComparison.OrdinalIgnoreCase) &&
                                                                                      x.Parameter.Equals("onProgramStartedBinOutFunction", StringComparison.OrdinalIgnoreCase));
                int binEfuseEcidOtherIdx = mainInitflowSheet.Rows.FindIndex(x => !x.IsBackup &&
                                                                                   x.Opcode.Equals("BinTable", StringComparison.OrdinalIgnoreCase) &&
                                                                                   x.Parameter.Equals("Bin_EFUSE_ecid_other", StringComparison.OrdinalIgnoreCase));
                FlowRow binInitFlowBinOut = mainInitflowSheet.Rows.FirstOrDefault(x => !x.IsBackup &&
                                                                                      x.Opcode.Equals("BinTable", StringComparison.OrdinalIgnoreCase) &&
                                                                                      x.Parameter.Equals("Bin_initFlow_BinOut", StringComparison.OrdinalIgnoreCase));

                if (onProgramStartedBinOutFunction == null)
                {
                    onProgramStartedBinOutFunction = new FlowRow() { Opcode = "Test", Parameter = "onProgramStartedBinOutFunction", FailAction = "F_initFlow_BinOut" };
                    if (binEfuseEcidOtherIdx != -1)
                    {
                        mainInitflowSheet.Rows.Insert(binEfuseEcidOtherIdx + 1, onProgramStartedBinOutFunction);
                    }
                    else
                    {
                        mainInitflowSheet.Rows.Add(onProgramStartedBinOutFunction);
                    }
                }

                int onProgramStartedBinOutFunctionIdx = mainInitflowSheet.Rows.FindIndex(x => !x.IsBackup &&
                                                                                      x.Opcode.Equals("Test", StringComparison.OrdinalIgnoreCase) &&
                                                                                      x.Parameter.Equals("onProgramStartedBinOutFunction", StringComparison.OrdinalIgnoreCase));

                if (binInitFlowBinOut == null)
                {
                    binInitFlowBinOut = new FlowRow() { Opcode = "BinTable", Parameter = "Bin_initFlow_BinOut" };

                    if (onProgramStartedBinOutFunctionIdx != -1)
                    {
                        mainInitflowSheet.Rows.Insert(onProgramStartedBinOutFunctionIdx + 1, binInitFlowBinOut);
                    }
                    else
                    {
                        mainInitflowSheet.Rows.Add(binInitFlowBinOut);
                    }
                }
                string outputFolder = LocalSpecs.InputParam.GenTxtOnly
                ? LocalSpecs.OutputFolder
                : Path.Combine(LocalSpecs.OutputFolder, ConstData.ModuleFolder, "Main");
                string outputPath = Path.Combine(outputFolder, mainInitflowSheet.Name + ".txt");
                mainInitflowSheet.Write(outputPath);
                LocalSpecs.GenSheets.Add(mainInitflowSheet);
            }
            else
            {
                LocalSpecs.MessageWriter.WriteLine("Can't get Flow_Table_Main_Init_Flows in test program!");
            }
        }

        private void _ProcessMainInitEnableWd()
        {
            SubFlowSheet mainInitEnableWdSheet = LocalSpecs.TestProgram.FlowSheetsAll.FirstOrDefault(p => p.Name
                .Equals("Flow_Table_Main_Init_EnableWd", StringComparison.OrdinalIgnoreCase));
            if (mainInitEnableWdSheet != null)
            {
                int ftIdx = mainInitEnableWdSheet.Rows.FindIndex(x => !x.IsBackup &&
                                                                                   x.Opcode.Equals("nop", StringComparison.OrdinalIgnoreCase) &&
                                                                                   x.Enable.Equals("FT", StringComparison.OrdinalIgnoreCase));
                FlowRow czSetupCheckEnableWord = mainInitEnableWdSheet.Rows.FirstOrDefault(x => !x.IsBackup &&
                                                                                             x.Enable.Equals("CZsetup_chk", StringComparison.OrdinalIgnoreCase));
                int printEndIdx = mainInitEnableWdSheet.Rows.FindLastIndex(x => !x.IsBackup &&
                                                                               x.Opcode.Equals("print", StringComparison.OrdinalIgnoreCase) &&
                                                                               x.Parameter.Equals("Flow_Table_Main_Init_EnableWd Stop", StringComparison.OrdinalIgnoreCase));

                if (czSetupCheckEnableWord == null)
                {
                    czSetupCheckEnableWord = new FlowRow() { Enable = "CZsetup_chk", Opcode = "nop" };
                    if (ftIdx != -1)
                    {
                        mainInitEnableWdSheet.Rows.Insert(ftIdx + 1, czSetupCheckEnableWord);
                    }
                    else
                    {
                        if (printEndIdx != -1)
                        {
                            mainInitEnableWdSheet.Rows.Insert(printEndIdx + 1, czSetupCheckEnableWord);
                        }
                    }
                    mainInitEnableWdSheet.Rows.Add(czSetupCheckEnableWord);
                }
                string outputFolder = LocalSpecs.InputParam.GenTxtOnly
                ? LocalSpecs.OutputFolder
                : Path.Combine(LocalSpecs.OutputFolder, ConstData.ModuleFolder, "Main");
                string outputPath = Path.Combine(outputFolder, mainInitEnableWdSheet.Name + ".txt");
                mainInitEnableWdSheet.Write(outputPath);
                LocalSpecs.GenSheets.Add(mainInitEnableWdSheet);
            }
            else
            {
                LocalSpecs.MessageWriter.WriteLine("Can't get Flow_Table_Main_Init_EnableWd in test program!");
            }
        }

        private string _TNumHandler(string tNum)
        {
            if (_genTNum)
            {
                return tNum;
            }

            return "";
        }
    }
}
