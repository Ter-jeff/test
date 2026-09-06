using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.Basic;
using Automation.GenerateIgxl.BinCut;
using Automation.GenerateIgxl.BistBira;
using Automation.GenerateIgxl.EFuse;
using Automation.GenerateIgxl.EVS;
using Automation.GenerateIgxl.HardIp;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HTOL;
using Automation.GenerateIgxl.PostAction;
using Automation.GenerateIgxl.PreAction;
using Automation.GenerateIgxl.Scan;
using Automation.GenerateIgxl.SpiRom;
using Automation.PreCheck.AllParaData;
using Automation.Reader.ConfigFile.RtosTable;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using OfficeOpenXml;

namespace Automation
{
    public class GenerateIgxlMain
    {
        private static readonly object _lock = new object();
        private static int _returnValue;
        public static int ReturnValue
        {
            get
            {
                lock (_lock)
                {
                    return _returnValue;
                }
            }
            set
            {
                lock (_lock)
                {
                    _returnValue = value;
                }
            }
        }

        public void Run(bool skipCsvDoc = false, bool ignoreHardIpInfo = false)
        {
            #region Pre-Action
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.PreAction).Down)
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.PreAction, () =>
                {
                    Response.Report("Running Pre-Action ~", EnumMessageLevel.CheckPoint);
                    using (var preActionMain = new PreActionMain())
                    {
                        preActionMain.Execute(null);
                    }
                    Response.Report("Pre-Action Completed !", EnumMessageLevel.EndPoint);
                });
            }
            #endregion

            #region Basic
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Basic).Down)
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.Basic, () =>
                {
                    Response.Report("Running Basic ~", EnumMessageLevel.CheckPoint);
                    using (var basicMain = new BasicMain())
                    {
                        basicMain.Execute(null);
                    }
                    Response.Report("Basic Completed !", EnumMessageLevel.EndPoint);
                });
            }
            #endregion

            #region eFuse
            RunEFuse();
            #endregion

            #region Ids
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.HardIp).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Ids))
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.Ids, () =>
                {
                    Response.Report("Running IDS ~", EnumMessageLevel.CheckPoint);
                    var hardIpParaData = new HardIpParaData(EnumBlock.Ids);
                    using (var hardIpMain = new HardIpMain())
                    {
                        hardIpMain.Execute(hardIpParaData);
                    }
                    Response.Report("IDS Completed !", EnumMessageLevel.EndPoint);
                });
            }
            #endregion

            #region HardIP
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.HardIp).Down && LocalSpecs.IsModuleIncluded(BlockStatus.HardIp))
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.HardIp, () =>
                {
                    Response.Report("Running HardIP ~", EnumMessageLevel.CheckPoint);
                    bool checkSplitCzFlow = LocalSpecs.Options.SplitCzFlow;
                    var hardIpParaData = new HardIpParaData(EnumBlock.HardIp)
                    {
                        SplitCzFlow = checkSplitCzFlow
                    };
                    using (var hardIpMain = new HardIpMain())
                    {
                        hardIpMain.Execute(hardIpParaData);
                    }
                    Response.Report("HardIP Completed !", EnumMessageLevel.EndPoint);
                });
            }
            #endregion

            #region Tmps
            RunTmps();
            #endregion

            //Need to be generated after HardIP because DcTest_IDS rule.
            #region Rtos
            RunRtos();
            #endregion

            #region Scan
            RunScan();
            #endregion

            #region HTOL
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Htol).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Htol))
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.Htol, () =>
                {
                    Response.Report("Running HTOL ~", EnumMessageLevel.CheckPoint);
                    using (var htolMain = new HtolMain())
                    {
                        htolMain.Execute(null);
                    }
                    Response.Report("HTOL Completed !", EnumMessageLevel.EndPoint);
                });
            }
            #endregion

            #region Mbist
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Mbist).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Mbist))
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.Mbist, () =>
                {
                    Response.Report("Running MBIST ~", EnumMessageLevel.CheckPoint);
                    using (var bistBiraMain = new BistBiraMain())
                    {
                        bistBiraMain.Execute(null);
                    }
                    Response.Report("MBIST Completed !", EnumMessageLevel.EndPoint);
                });
            }
            #endregion

            #region EVS
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Evs).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Evs))
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.Evs, () =>
                {
                    Response.Report("Running EVS ~", EnumMessageLevel.CheckPoint);
                    using (var evsMain = new EvsMain())
                    {
                        evsMain.Execute(null);
                    }
                    Response.Report("EVS Completed !", EnumMessageLevel.EndPoint);
                });
            }
            #endregion

            #region BinCut
            LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.BinCut, () =>
            {
                using (var binCutMain = new BinCutMain())
                {
                    Response.Report("Generating Bincut Table ~", EnumMessageLevel.CheckPoint);
                    binCutMain.GenerateBincutRelatedFiles();
                    Response.Report("BinCut Table generation Completed !", EnumMessageLevel.EndPoint);
                    if (BlockStatus.GetAutomationBlockStatus(BlockStatus.BinCut).Down && LocalSpecs.IsModuleIncluded(BlockStatus.BinCut))
                    {
                        Response.Report("Running BinCut ~", EnumMessageLevel.CheckPoint);
                        {
                            binCutMain.Execute(null);
                        }
                        Response.Report("BinCut Completed !", EnumMessageLevel.EndPoint);
                    }
                }
            });
            #endregion

            #region PostAction
            LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.PostAction, () =>
            {
                Response.Report("Running Post-Action ~", EnumMessageLevel.CheckPoint);
                using (var postActionMain = new PostActionMain())
                {
                    postActionMain.Execute(null);
                    postActionMain.PrintFlow();
                }
                Response.Report("Post-Action Completed !", EnumMessageLevel.EndPoint);
            });
            #endregion

        }

        private static void RunEFuse()
        {
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Efuse).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Efuse))
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.Efuse, () =>
                {
                    if (LocalSpecs.Options.Device == EnumDevice.AP || LocalSpecs.Options.Device == EnumDevice.RF)
                    {
                        Response.Report("Running eFuse ~", EnumMessageLevel.CheckPoint);
                        using (var eFuseMain = new EFuseMain())
                        {
                            eFuseMain.Execute(null);
                        }
                        Response.Report("eFuse Completed !", EnumMessageLevel.EndPoint);
                    }
                });
            }
        }

        private static void RunTmps()
        {
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.HardIp).Down && LocalSpecs.IsModuleIncluded(nameof(EnumBlock.HardIp)) && LocalSpecs.IsModuleIncluded(nameof(EnumBlock.Ids)))
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.Tmps, () =>
                {
                    Response.Report("Generating TMPS ...", percentage: 90);
                    new TmpsGenerator().GenTmps(TestProgram.IgxlWorkBk.SubFlowSheets);
                    Response.Report("TMPS Completed !", EnumMessageLevel.EndPoint);
                });
            }
        }

        private void RunRtos()
        {
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Rtos).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Rtos))
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.Rtos, () =>
                {
                    Response.Report("Check Rtos Settings~", EnumMessageLevel.CheckPoint);
                    ExcelWorksheet rtosUartSheet = EpWorkbook.TestPlanWorkbook.Worksheets.ToList().Find(sheet => SearchInfo.IsHardipRtosSheet(sheet.Name));

                    using (var hardIpMain = new HardIpMain())
                    {
                        var hardIpParaData = new HardIpParaData(EnumBlock.Rtos);
                        hardIpMain.Execute(hardIpParaData);
                    }



                    using (var spiRomMain = new SpiRomMain())
                    {
                        spiRomMain.Execute(null);
                    }
                    Response.Report("Rtos Generation Completed!", EnumMessageLevel.EndPoint);
                    UpdateRtosTable(new UpdateRtosTableOptions()
                    {
                        SubGroupsSelector = _groupRtosArgRowsByFuncName,
                        ShouldSkipSubGroupRow = _skipByNotEqualGroupingKey,
                        SubflowSheetRowsSelector = _filterByFlowRowParameter,
                        InstanceSheetRowsSelector = _filterByInstanceRowTestName,
                    });
                    AddRtosConfiguration();
                });

            }
        }

        private static void RunScan()
        {
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Scan).Down && (LocalSpecs.IsModuleIncluded(BlockStatus.Scan) || LocalSpecs.IsModuleIncluded("CPM")))
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection(BlockConst.Scan, () =>
                {
                    Response.Report("Running Scan ~", EnumMessageLevel.CheckPoint);
                    using (var scanMain = new ScanMain())
                    {
                        scanMain.Execute(null);
                    }
                    Response.Report("Scan Completed !", EnumMessageLevel.EndPoint);
                });
            }
        }

        /// <summary>
        /// Move Rtos_Configuration sheet to T/P
        /// </summary>
        public static void AddRtosConfiguration()
        {
            if (EpWorkbook.TestPlanWorkbook.Worksheets.Any(x => x.Name.Equals("Rtos_Configuration", StringComparison.CurrentCultureIgnoreCase)))
            {
                ExcelWorksheet rtosConfigSheet = EpWorkbook.TestPlanWorkbook.Worksheets.First(x => x.Name.Equals("Rtos_Configuration", StringComparison.CurrentCultureIgnoreCase));

                rtosConfigSheet.ExportWorkBook2Txt(FolderStructure.DirRtos);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirRtos, rtosConfigSheet.Name);
            }
        }

        private static void AddFlowRowsToFlowSheets(
            FlowRows flowRows,
            KeyValuePair<string, string> keyValuePair
        )
        {
            if (!flowRows.Any())
            {
                return;
            }

            foreach (KeyValuePair<string, SubFlowSheet> flowsheet in TestProgram.IgxlWorkBk.SubFlowSheets)
            {
                if (flowsheet.Key.IndexOf("Flow_Rtos_UART", StringComparison.OrdinalIgnoreCase) >= 0
                    || flowsheet.Key.IndexOf("Flow_Rtos_IDS", StringComparison.OrdinalIgnoreCase) >= 0
                )
                {
                    int indexRow = flowsheet.Value.Rows.FindIndex(x =>
                        x.Parameter.IndexOf(keyValuePair.Key, StringComparison.CurrentCultureIgnoreCase) != -1);
                    if (indexRow >= 0)
                    {
                        flowsheet.Value.Rows.InsertRange(indexRow + 1, flowRows);
                    }
                }
            }
        }

        private static void AddInstanceRowsToFlowSheets(
            InstanceRows instanceRows,
            KeyValuePair<string, string> keyValuePair
        )
        {
            if (!instanceRows.Any())
            {
                return;
            }
            foreach (KeyValuePair<string, InstanceSheet> valuePair in TestProgram.IgxlWorkBk.InsSheets)
            {
                if (valuePair.Key.IndexOf("TestInst_Rtos", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    int indexRow = valuePair.Value.Rows
                        .FindIndex(x => x.TestName.IndexOf(keyValuePair.Key, StringComparison.CurrentCultureIgnoreCase) != -1);
                    if (indexRow >= 0)
                    {
                        valuePair.Value.Rows.InsertRange(indexRow + 1, instanceRows);
                    }
                }
            }
        }

        /// <summary>
        /// Update RTOS table
        /// </summary>
        /// <param name="options"></param>
        public static void UpdateRtosTable(UpdateRtosTableOptions options)
        {
            #region Update Rtos_Table info in instance sheets

            ExcelWorksheet rtosTableSheet = EpWorkbook.TestPlanWorkbook.Worksheets
                .ToList()
                .Find(x => x.Name.Equals("Rtos_Table", StringComparison.OrdinalIgnoreCase));
            if (rtosTableSheet == null)
            {
                return;
            }

            RtosTableSheet rtosTableInfos = RtosTableSheet.LoadConfig(rtosTableSheet);

            HandleArgRows(rtosTableInfos, options.SubGroupsSelector, options.ShouldSkipSubGroupRow);

            foreach (KeyValuePair<string, string> keyValuePair in rtosTableInfos.MeasRows)
            {
                FlowRows flowRows = GetAndUpdateFlowRows(options, keyValuePair);
                InstanceRows instanceRows = GetAndUpdateInstanceRows(options, keyValuePair);

                AddFlowRowsToFlowSheets(flowRows, keyValuePair);
                AddInstanceRowsToFlowSheets(instanceRows, keyValuePair);
            }

            #endregion
        }

        internal static void PopulateDc(
            IGrouping<string, RtosTableArgRow> grouping,
            InstanceRow instanceRow
        )
        {
            List<string> insArgs = instanceRow.ArgList.Split(',').ToList();
            foreach (RtosTableArgRow arg in grouping)
            {
                if (arg.ArgName == "DC Category" || arg.ArgName == "DCCategory")
                {
                    instanceRow.DcCategory = arg.ArgInfo;
                }
                else if (arg.ArgName == "DC Selector" || arg.ArgName == "DCSelector")
                {
                    instanceRow.DcSelector = arg.ArgInfo;
                }
                else if (arg.ArgName == "AC Category" || arg.ArgName == "ACCategory")
                {
                    instanceRow.AcCategory = arg.ArgInfo;
                }
                else if (arg.ArgName == "AC Selector" || arg.ArgName == "ACSelector")
                {
                    instanceRow.AcSelector = arg.ArgInfo;
                }
                else if (arg.ArgName == "Time Sets" || arg.ArgName == "TimeSets")
                {
                    instanceRow.TimeSets = arg.ArgInfo;
                }
                else if (arg.ArgName == "Pin Levels" || arg.ArgName == "PinLevels")
                {
                    instanceRow.PinLevels = arg.ArgInfo;
                }
                else
                {
                    int seqIndex = insArgs.FindIndex(s => s.Equals(arg.ArgName, StringComparison.OrdinalIgnoreCase));
                    if (seqIndex != -1)
                    {
                        instanceRow.Args[seqIndex] = arg.ArgInfo;
                    }
                }
            }
        }

        /// <summary>
        /// Update target flow rows parameter and return them
        /// </summary>
        /// <param name="options"></param>
        /// <param name="keyValuePair"></param>
        /// <returns></returns>
        private static FlowRows GetAndUpdateFlowRows(
            UpdateRtosTableOptions options,
            KeyValuePair<string, string> keyValuePair
        )
        {
            var flowRows = new FlowRows();
            foreach (KeyValuePair<string, SubFlowSheet> flowsheet in TestProgram.IgxlWorkBk.SubFlowSheets)
            {
                if (flowsheet.Key.IndexOf("DcTest_IDS", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    List<FlowRow> measIdsFlowRows = options.SubflowSheetRowsSelector(flowsheet, keyValuePair);
                    if (measIdsFlowRows.Any())
                    {
                        foreach (FlowRow row in measIdsFlowRows)
                        {
                            row.Parameter = Regex.Replace(row.Parameter, "^IDS_", "Rtos_");
                        }

                        flowRows.AddRange(measIdsFlowRows);
                    }
                }
            }

            return flowRows;
        }

        private static InstanceRows GetAndUpdateInstanceRows(
            UpdateRtosTableOptions options,
            KeyValuePair<string, string> keyValuePair
        )
        {
            var instanceRows = new InstanceRows();
            foreach (KeyValuePair<string, InstanceSheet> valuePair in TestProgram.IgxlWorkBk.InsSheets)
            {
                if (valuePair.Key.IndexOf("DcTest_IDS", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    List<InstanceRow> measIdsInstRows = options.InstanceSheetRowsSelector(valuePair.Value, keyValuePair);
                    if (measIdsInstRows.Any())
                    {
                        foreach (InstanceRow row in measIdsInstRows)
                        {
                            row.TestName = Regex.Replace(row.TestName, "^IDS_", "Rtos_");
                            int index = row.ArgList.Split(',').ToList().FindIndex(x => x.Equals("patt") || x.Equals("pattern"));
                            if (index != -1)
                            {
                                row.Args[index] = "";
                            }
                        }
                        instanceRows.AddRange(measIdsInstRows);
                    }
                }
            }
            return instanceRows;
        }

        private static void HandleArgRows(
            RtosTableSheet rtosTableInfos,
            Func<
                IGrouping<string, RtosTableArgRow>,
                IEnumerable<IGrouping<string, RtosTableArgRow>>
            > subGroupsSelector,
            Func<IGrouping<string, RtosTableArgRow>, InstanceRow, bool> shouldSkipSubGroupRow
        )
        {
            IEnumerable<IGrouping<string, RtosTableArgRow>> groups = rtosTableInfos.ArgRows
                .GroupBy(x => x.SheetName);

            foreach (IGrouping<string, RtosTableArgRow> group in groups)
            {
                //Get instance sheet
                InstanceSheet instSheet = TestProgram.IgxlWorkBk.InsSheets.Values.ToList()
                    .Find(p => p.Name.Equals(group.Key, StringComparison.OrdinalIgnoreCase));

                if (instSheet == null)
                {
                    continue;
                }

                IEnumerable<IGrouping<string, RtosTableArgRow>> subGroups = subGroupsSelector(group);

                foreach (IGrouping<string, RtosTableArgRow> grouping in subGroups)
                {
                    foreach (InstanceRow instanceRow in instSheet.Rows)
                    {
                        //Get function name
                        if (shouldSkipSubGroupRow(grouping, instanceRow))
                        {
                            continue;
                        }

                        PopulateDc(grouping, instanceRow);
                    }
                }
            }
        }

        private readonly Func<
            IGrouping<string, RtosTableArgRow>,
            IEnumerable<IGrouping<string, RtosTableArgRow>>
        > _groupRtosArgRowsByFuncName = (group) => group.GroupBy(
            x => x.FuncName.Split('.').LastOrDefault()
        );

        private readonly Func<
            IGrouping<string, RtosTableArgRow>,
            InstanceRow,
            bool
        > _skipByNotEqualGroupingKey = (grouping, instanceRow) => !instanceRow.VbtName
            .Split('.')
            .Last()
            .Equals(grouping.Key, StringComparison.OrdinalIgnoreCase);

        private readonly Func<
            KeyValuePair<string, SubFlowSheet>,
            KeyValuePair<string, string>,
            List<FlowRow>
        > _filterByFlowRowParameter = (flowsheet, keyValuePair) => flowsheet.Value.Rows
            .FindAll(flowRow => flowRow.Parameter
                .Equals(keyValuePair.Value, StringComparison.OrdinalIgnoreCase))
            .Select(x => x.Copy())
            .ToList();

        private readonly Func<
            InstanceSheet,
            KeyValuePair<string, string>,
            List<InstanceRow>
        > _filterByInstanceRowTestName = (instanceSheet, keyValuePair) => instanceSheet.Rows
            .FindAll(instanceRow => instanceRow.TestName
                .Equals(keyValuePair.Value, StringComparison.OrdinalIgnoreCase))
            .Select(x => x.Copy())
            .ToList();

        /// <summary>
        /// All necessary functions that is needed to update RTOS table
        /// </summary>
        public sealed class UpdateRtosTableOptions
        {
            /// <summary>
            /// Handle how to group or filter RTOS table args rows that is already grouped by 
            /// sheet name.
            /// </summary>
            public Func<
                IGrouping<string, RtosTableArgRow>,
                IEnumerable<IGrouping<string, RtosTableArgRow>>
            > SubGroupsSelector
            { get; set; }

            /// <summary>
            /// Decide if sub-group row should be skipped
            /// </summary>
            public Func<
                IGrouping<string, RtosTableArgRow>,
                InstanceRow,
                bool
            > ShouldSkipSubGroupRow
            { get; set; }

            /// <summary>
            /// Filter which rows in subflow sheet should be fetched.
            /// </summary>
            public Func<
                KeyValuePair<string, SubFlowSheet>,
                KeyValuePair<string, string>,
                List<FlowRow>
            > SubflowSheetRowsSelector
            { get; set; }

            /// <summary>
            /// Filter which instance rows in instance sheet should be fetched.
            /// </summary>
            public Func<
                InstanceSheet,
                KeyValuePair<string, string>,
                List<InstanceRow>
            > InstanceSheetRowsSelector
            { get; set; }
        }
    }
}
