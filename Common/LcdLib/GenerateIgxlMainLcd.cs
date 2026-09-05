using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.BinCut;
using Automation.GenerateIgxl.BistBira;
using Automation.GenerateIgxl.EFuse;
using Automation.GenerateIgxl.EVS;
using Automation.GenerateIgxl.HardIp;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
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
using IgxlLib.IgxlSheets;

using LcdLib.Basic;
using LcdLib.OTP;

using LogLib.Static;

using RfLib.Dvdc;

namespace LcdLib
{
    public class GenerateIgxlMainLcd
    {
        private readonly Func<IGrouping<string, RtosTableArgRow>, IEnumerable<IGrouping<string, RtosTableArgRow>>> _groupRtosArgRowsByFuncName = (group) => group.GroupBy(x => x.FuncName);

        private readonly Func<IGrouping<string, RtosTableArgRow>, InstanceRow, bool> _skipByNotEqualGroupingKey = (grouping, instanceRow) => !instanceRow.VbtName.EqualsIgnoreCase(grouping.Key);

        private readonly Func<KeyValuePair<string, SubFlowSheet>, KeyValuePair<string, string>,
            List<FlowRow>> _filterByFlowRowParameter = (flowsheet, keyValuePair) => [.. flowsheet.Value.Rows
            .FindAll(flowRow => flowRow.Parameter.EqualsIgnoreCase(keyValuePair.Value))
            .Select(flowRow => flowRow.Copy())];

        private readonly Func<InstanceSheet, KeyValuePair<string, string>, List<InstanceRow>
        > _filterByInstanceRowTestName = (instanceSheet, keyValuePair) => [.. instanceSheet.Rows.FindAll(instanceRow => instanceRow.TestName.EqualsIgnoreCase(keyValuePair.Value)
            ).Select(instanceRow => instanceRow.Copy())];

        public static void MeasureSection(string name, Action action)
        {
            var sw = Stopwatch.StartNew();
            action();
            sw.Stop();
            TimeSpan elapsed = sw.Elapsed;
            string text = elapsed.ToString("mm:ss:fff");
            Response.Report($"[Elapsed Time][{name}] {text}", EnumMessageLevel.CheckPoint);
        }

        public void Run(bool skipCsvDoc = false, bool ignoreHardIpInfo = false)
        {
            RunPreActionStep();
            RunBasicStep();
            RunEfuseStep();
            RunOtpStep();
            RunIdsStep();
            RunHardIpStep();
            RunTmpsStep();
            RunRtosStep();
            RunScanStep();
            RunHtolStep();
            RunMbistStep();
            RunEvsStep();
            RunDvdcStep();
            RunBinCutStep();
            RunPostActionStep();
        }

        private static void RunPreActionStep()
        {
            #region Pre-Action
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.PreAction).Down)
            {
                MeasureSection("Pre-Action", () =>
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
        }

        private static void RunBasicStep()
        {
            #region Basic
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Basic).Down)
            {
                MeasureSection("Basic", () =>
                {
                    Response.Report("Running Basic ~", EnumMessageLevel.CheckPoint);
                    using (var basicMain = new BasicMainLcd())
                    {
                        basicMain.Execute(null!);
                    }
                    Response.Report("Basic Completed !", EnumMessageLevel.EndPoint);
                });
            }
            #endregion
        }

        private static void RunEfuseStep()
        {
            #region eFuse
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Efuse).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Efuse))
            {
                MeasureSection("eFuse", () =>
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
            #endregion
        }

        private static void RunOtpStep()
        {
            #region Otp
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Otp).Down)
            {
                MeasureSection("Otp", () =>
                {
                    if (LocalSpecs.Options.Device == EnumDevice.LCD)
                    {
                        Response.Report("Running OTP ~", EnumMessageLevel.CheckPoint);
                        using (var otpMain = new OtpMain())
                        {
                            otpMain.Execute(null!);
                        }
                        Response.Report("OTP Completed !", EnumMessageLevel.EndPoint);
                    }
                });
            }
            #endregion
        }

        private static void RunIdsStep()
        {
            #region Ids
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.HardIp).Down && LocalSpecs.IsModuleIncluded(nameof(EnumBlock.Ids)))
            {
                MeasureSection("Ids", () =>
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
        }

        private static void RunHardIpStep()
        {
            #region HardIP
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.HardIp).Down && LocalSpecs.IsModuleIncluded(nameof(EnumBlock.HardIp)))
            {
                MeasureSection("HardIP", () =>
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
        }

        private static void RunTmpsStep()
        {
            #region Tmps
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.HardIp).Down && LocalSpecs.IsModuleIncluded(nameof(EnumBlock.HardIp)) && LocalSpecs.IsModuleIncluded(nameof(EnumBlock.Ids)))
            {
                MeasureSection("Tmps", () =>
                {
                    Response.Report("Generating TMPS ...", percentage: 90);
                    new TmpsGenerator().GenTmps(TestProgram.IgxlWorkBk.SubFlowSheets);
                    Response.Report("TMPS Completed !", EnumMessageLevel.EndPoint);
                });
            }
            #endregion
        }

        private void RunRtosStep()
        {
            //Need to be generated after HardIP because DcTest_IDS rule.
            #region Rtos
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Rtos).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Rtos))
            {
                MeasureSection("Rtos", () =>
                {
                    Response.Report("Check Rtos Settings~", EnumMessageLevel.CheckPoint);
                    using var hardIpMain = new HardIpMain();
                    var hardIpParaData = new HardIpParaData(EnumBlock.Rtos);
                    hardIpMain.Execute(hardIpParaData);


                    using (var spiRomMain = new SpiRomMain())
                    {
                        spiRomMain.Execute(null);
                    }
                    Response.Report("Rtos Generation Completed!", EnumMessageLevel.EndPoint);
                    Automation.GenerateIgxlMain.UpdateRtosTable(new Automation.GenerateIgxlMain.UpdateRtosTableOptions()
                    {
                        SubGroupsSelector = _groupRtosArgRowsByFuncName,
                        ShouldSkipSubGroupRow = _skipByNotEqualGroupingKey,
                        SubflowSheetRowsSelector = _filterByFlowRowParameter,
                        InstanceSheetRowsSelector = _filterByInstanceRowTestName,
                    });
                    Automation.GenerateIgxlMain.AddRtosConfiguration();
                });

            }
            #endregion
        }

        private static void RunScanStep()
        {
            #region Scan
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Scan).Down && (LocalSpecs.IsModuleIncluded(BlockStatus.Scan) || LocalSpecs.IsModuleIncluded("CPM")))
            {
                MeasureSection("Scan", () =>
                {
                    Response.Report("Running Scan ~", EnumMessageLevel.CheckPoint);
                    using (var scanMain = new ScanMain())
                    {
                        scanMain.Execute(null);
                    }
                    Response.Report("Scan Completed !", EnumMessageLevel.EndPoint);
                });
            }
            #endregion
        }

        private static void RunHtolStep()
        {
            #region HTOL
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Htol).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Htol))
            {
                MeasureSection("HTOL", () =>
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
        }

        private static void RunMbistStep()
        {
            #region Mbist
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Mbist).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Mbist))
            {
                MeasureSection("Mbist", () =>
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
        }

        private static void RunEvsStep()
        {
            #region EVS
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Evs).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Evs))
            {
                MeasureSection("EVS", () =>
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
        }

        private static void RunDvdcStep()
        {
            #region DVDC
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Dvdc).Down || BlockStatus.GetAutomationBlockStatus(BlockStatus.Lcd).Down)
            {
                MeasureSection("Dvdc", () =>
                {
                    Response.Report("Running Dvdc ~", EnumMessageLevel.CheckPoint);
                    var hardIpParaData = new HardIpParaData(EnumBlock.Dvdc);
                    using (var evsMain = new DvdcMain1())
                    {
                        evsMain.Execute(hardIpParaData);
                    }
                    Response.Report("Dvdc Completed !", EnumMessageLevel.EndPoint);
                });
            }
            #endregion
        }

        private static void RunBinCutStep()
        {
            #region BinCut
            MeasureSection(BlockConst.BinCut, () =>
            {
                using var binCutMain = new BinCutMain();
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
            });
            #endregion
        }

        private static void RunPostActionStep()
        {
            #region PostAction
            MeasureSection("Post-Action", () =>
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
    }
}
