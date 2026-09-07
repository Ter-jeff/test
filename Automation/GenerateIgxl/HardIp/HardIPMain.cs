using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenCharBiz;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.HardIPUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;
using Automation.Utility.TpUpdate.HardIPEnableWordsUpdate;

using CommonLib.Enums;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LogLib.Static;

namespace Automation.GenerateIgxl.HardIp
{
    public class HardIpMain : WorkFlowBase<HardIpParaData>
    {
        protected HardIpInputData HardIpInputData;

        private static readonly Regex _regex1 = new Regex("^HARDIP_|DCTEST_IDS|DCTEST_GPIO|DCTEST_FailSafe|PLLDEBUG_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex2 = new Regex(@"HardIP_(?<name>\w+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex3 = new Regex(@"Inst_H_(?<name>\w+)", RegexOptions.IgnoreCase);
        private Dictionary<EnumBlock, HardIpInputData> _mapSheet
             = new Dictionary<EnumBlock, HardIpInputData>()
                {
                    { EnumBlock.Rtos, TestPlanStatic.RtosSheets },
                    { EnumBlock.Ids,  TestPlanStatic.IdsSheets }
                };


        public override bool PreCheckFlow(HardIpParaData paraData)
        {
            try
            {

                if (_mapSheet.TryGetValue(paraData.Block, out HardIpInputData data))
                {
                    HardIpInputData = data;
                }
                else
                {
                    HardIpInputData = new HardIpInputManager(EpWorkbook.TestPlanWorkbook, paraData).Read();
                }

                if (!LocalSpecs.Options.BypassPreCheck)
                {
                    new HardIpPreCheckMain().Check(HardIpInputData.PlanDic, HardIpInputData.HardIpDcSheet);
                }

                return true;
            }
            catch (Exception e)
            {
                Response.Report("HardIp has errors : " + e.StackTrace, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
                return false;
            }
        }

        public override void WorkFlow(HardIpParaData paraData)
        {
            try
            {
                #region Add default enableWd to Flow_Table_Main_Init_EnableWd sheet
                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("Enable_HardIP_FieldProcesingTTR", OpCode.Nop, HardIpConstData.HardIp, FolderStructure.DirMain, "");
                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("Enable_HardIP_DigitalTestLimitTTR", OpCode.Nop, HardIpConstData.HardIp, FolderStructure.DirMain, "");
                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("Enable_HardIP_TnameConstructionTTR", OpCode.Nop, HardIpConstData.HardIp, FolderStructure.DirMain, "");
                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("Enable_HardIP_AnalogMuxOutTTR", OpCode.Nop, HardIpConstData.HardIp, FolderStructure.DirMain, "");

                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("", OpCode.EnableFlowWd, HardIpConstData.HardIp, FolderStructure.DirMain, "Enable_HardIP_FieldProcesingTTR");
                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("", OpCode.EnableFlowWd, HardIpConstData.HardIp, FolderStructure.DirMain, "Enable_HardIP_DigitalTestLimitTTR");
                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("", OpCode.EnableFlowWd, HardIpConstData.HardIp, FolderStructure.DirMain, "Enable_HardIP_TnameConstructionTTR");
                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("", OpCode.EnableFlowWd, HardIpConstData.HardIp, FolderStructure.DirMain, "Enable_HardIP_AnalogMuxOutTTR");
                #endregion

                #region divide pattern according to pattern.ForceCondition
                ParseTestPlanByCondition parseTestPlanMain = new ParseTestPlanByCondition();
                parseTestPlanMain.ParseTestPlanPatternByCondition(HardIpInputData.PlanDic);
                #endregion

                #region Pre-Process for Misc Info
                MiscInfoPreProcessor.PreProcessMiscInfo(HardIpInputData.PlanDic);
                #endregion

                Response.Report("Generating TimeSet for HardIP ...", percentage: 25);
                List<TimeSetBasicSheet> timeSetSheets = new TimeSetSheetGenerator().GenTimeSet(HardIpInputData.PlanDic);

                Response.Report("Generating Multiple Init PatSet ...", percentage: 30);
                PatSetSheet patSetSheet = new PatSetSheetGenerator(HardIpInputData).GenPatSet(HardIpInputData.PlanDic, ScghStatic.ScghData);

                Response.Report("Generating Instance for HardIP ...", percentage: 45);
                List<InstanceSheet> instSheets = new InstanceGenerator(HardIpInputData).GenInst(HardIpInputData.PlanDic);

                Response.Report("Generating Test flow for HardIP ...", percentage: 55);
                List<SubFlowSheet> flowSheets = new FlowGenerator(HardIpInputData).GenFlow(HardIpInputData.PlanDic);

                Response.Report("Generating Bin Table for HardIP ...", percentage: 60);
                BinTableSheet binTableSheet = new BinTableSheetGenerator(HardIpInputData).GenBinTable(HardIpInputData.PlanDic);

                Response.Report("Generating Characterization for HardIP ...", percentage: 70);
                CharSheet charSheet = new CharSheetGenerator().GenCharSheet(HardIpInputData.PlanDic);

                Response.Report("Generating Relay Instance for HardIP ...", percentage: 75);
                new RelayInstanceGenerator().GenRelayInstance(HardIpInputData.PlanDic);

                Response.Report("Generating Nwire Instance for HardIP ...", percentage: 80);
                new NwireInstanceGenerator().GenNwireInstance(HardIpInputData.PlanDic, instSheets);

                Response.Report("Generating AcCategory for HardIP nWire ...", percentage: 85);
                new AcCategoryGenerator().GenAcCategory(HardIpInputData.PlanDic, instSheets);

                Response.Report("Generating HardIP Init Flag Flow ...", percentage: 95);
                if (HardIpInputData.HardIpParaData.Block != EnumBlock.Rtos)
                {
                    SubFlowSheet initFlow = new InitFlagGenerator().GenInitFlag(HardIpInputData.PlanDic, binTableSheet);
                    #region Add result sheets to result workbook
                    TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirHardIp, initFlow);
                    #endregion
                }
                if (HardIpInputData.HardIpRegAssigns.Count > 0)
                {
                    new RegAssignGenerator(HardIpInputData).WorkFlow(HardIpInputData.HardIpRegAssigns);
                }
                if (HardIpInputData.InterposeAssigns.Count > 0)
                {
                    new InterposeAssignGenerator().WorkFlow(HardIpInputData.InterposeAssigns.ToList());
                }
                if (HardIpInputData.HardIpParaData.Block == EnumBlock.Ids || HardIpInputData.HardIpParaData.Block == EnumBlock.Rtos)
                {
                    new IdsMappingTableGenerator(TestPlanStatic.IdsMappingSheet).WorkFlow();
                }

                Response.Report("Adding HardIP Sheets into Project Object ...", percentage: 95);

                var nonUsedItems = new List<string>();
                var tableUpdateStatus = new Dictionary<string, string>();
                var failControls = new List<FailControlData>();
                AddFlowSheetsToWorkbook(flowSheets, tableUpdateStatus, failControls, ref nonUsedItems);

                string failControlTablePath = new TtrToolWriter().WriteFailControlTable(FolderStructure.DirHardIp, failControls);
                if (!string.IsNullOrEmpty(failControlTablePath))
                {
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirHardIp, failControlTablePath);
                }

                AddInstanceSheetsToWorkbook(instSheets);

                TestProgram.IgxlWorkBk.AddBinTblSheet(FolderStructure.DirBinTable, binTableSheet);

                if (charSheet.Rows.Count > 0)
                {
                    TestProgram.IgxlWorkBk.AddCharSheet(FolderStructure.DirDevChar, charSheet);
                }

                foreach (TimeSetBasicSheet timeSetSheet in timeSetSheets)
                {
                    TestProgram.IgxlWorkBk.AddTimeSetSheet(FolderStructure.DirTimings, timeSetSheet);
                }

                if (patSetSheet.Rows.Count != 0)
                {
                    TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirPatSetsAll, patSetSheet);
                }

                JitterSheet jitterSheet = new JitterGenerator().GenJitterSheet(HardIpInputData.PlanDic);
                if (jitterSheet != null)
                {
                    TestProgram.IgxlWorkBk.AddJitterSheet(FolderStructure.DirCommon, jitterSheet);
                }

            }
            catch (Exception e)
            {
                string message = "HardIP AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }

        private void AddFlowSheetsToWorkbook(List<SubFlowSheet> flowSheets, Dictionary<string, string> tableUpdateStatus, List<FailControlData> failControls, ref List<string> nonUsedItems)
        {
            foreach (SubFlowSheet flowSheet in flowSheets)
            {
                if (flowSheet.Rows.Count > 0)
                {
                    string blockName = _regex2.Match(flowSheet.Name).Groups["name"].Value;
                    var enableList = HardIpInputData.TtrTables.Values.SelectMany(p => p.EnableWords).Distinct().ToList();
                    if (!HardIpInputData.TtrTables.ContainsKey(blockName) && blockName.EndsWith("_CZ"))
                    {
                        blockName = blockName.Substring(0, blockName.Length - 3);
                    }

                    if (HardIpInputData.TtrTables.TryGetValue(blockName, out EnableWordTable value))
                    {
                        (SubFlowSheet tmpFlow, List<FailControlData> failControlList) = new TtrToolWriter().UpdateFlowSheetEnableWords
                        (
                            flowSheet,
                            value,
                            enableList,
                            tableUpdateStatus,
                            ref nonUsedItems,
                            !LocalSpecs.Options.ConvertJobToEnableWord ? LocalSpecs.AllJobs : null

                        );
                        failControls.AddRange(failControlList);
                        TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirHardIp, tmpFlow);
                    }
                    else if (blockName.ToUpper().Contains("OVERLAY"))
                    {
                        if (HardIpInputData.TtrTables.TryGetValue("AMP", out EnableWordTable ampTable))
                        {
                            SubFlowSheet tmpFlow = new TtrToolWriter().UpdateOverlayFlowSheetEnableWords
                            (
                                flowSheet,
                                ampTable,
                                enableList,
                                tableUpdateStatus,
                                ref nonUsedItems,
                                !LocalSpecs.Options.ConvertJobToEnableWord ? LocalSpecs.AllJobs : null

                            );
                            TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirHardIp, tmpFlow);
                        }
                        TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirHardIp, flowSheet);
                    }
                    else
                    {
                        TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirHardIp, flowSheet);
                    }
                }
            }
        }

        private void AddInstanceSheetsToWorkbook(List<InstanceSheet> instSheets)
        {
            foreach (InstanceSheet instSheet in instSheets)
            {
                if (instSheet.Rows.Count > 0)
                {
                    string ip = _regex3.Match(instSheet.Name).Groups["name"].Value;
                    if (HardIpInputData.TtrTables.TryGetValue(ip, out EnableWordTable value))
                    {
                        IEnumerable<HardIpPattern> patterns = HardIpInputData.PlanDic.Select(x => x).SelectMany(x => x.Value.Rows);
                        var binCutOverlay = patterns.Where(x => !string.IsNullOrEmpty(x.BinCutOverlay.ateTestCondition)).Select(x => x.BinCutOverlay).GroupBy(x => x.ateTestCondition).ToDictionary(x => x.Key, x => x.First().overlayName);
                        InstanceSheet updatedSheet = new TtrToolWriter().UpdateInsSheetOverlays(instSheet, value, binCutOverlay);
                        TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirHardIp, updatedSheet);
                    }
                    else
                    {
                        TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirHardIp, instSheet);
                    }
                }
            }
        }

        public virtual Dictionary<string, HardIpSheet> ReadHardipSheets(HardIpParaData para, List<string> usedPatterns = null)
        {
            var hardIpReader = new TestPlanConverter(ScghStatic.ScghData);
            Dictionary<string, HardIpSheet> dic = hardIpReader.ReadHardipPatterns(EpWorkbook.TestPlanWorkbook, IsValidSheet, usedPatterns);
            var regAssignParser = new RegisterAssignParser();
            regAssignParser.ParseRegisterAssign(dic);
            return dic;
        }

        public virtual bool IsValidSheet(string sheetName)
        {
            if (_regex1.IsMatch(sheetName))
            {
                return true;
            }

            return false;
        }
    }
}
