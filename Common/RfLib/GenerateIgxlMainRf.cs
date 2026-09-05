using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.BinCut;
using Automation.GenerateIgxl.BistBira;
using Automation.GenerateIgxl.EFuse;
using Automation.GenerateIgxl.EVS;
using Automation.GenerateIgxl.HardIp;
using Automation.GenerateIgxl.HTOL;
using Automation.GenerateIgxl.PostAction;
using Automation.GenerateIgxl.PreAction;
using Automation.GenerateIgxl.Scan;
using Automation.GenerateIgxl.SpiRom;
using Automation.PreCheck.AllParaData;
using Automation.Static;
using Automation.Utility;

using CommonLib.Enums;
using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using LogLib.Static;

using OfficeOpenXml;

using ProjectConfigLib.ProjectConfig;

using RfLib.Basic;
using RfLib.Dvdc;

using TestPlanLib.BinCut.NonIgxlSheet;

namespace RfLib
{
    public class GenerateIgxlMainRf
    {
        public static bool Run(bool skipCsvDoc = false, bool ignoreHardIpInfo = false)
        {
            bool result = RunTestProgramGenerator();
            RunAutomationPostAction();
            return result;
        }

        public static bool GenerateTestProgram(List<string> tpSheetsList, List<string> scghSheetsList, bool skipCsvDoc = false, bool ignoreHardIpInfo = false)
        {
            return RunTestProgramGenerator();
        }

        internal static bool RunTestProgramGenerator()
        {
            bool checkSplitCzFlow = ProjectConfigSingleton.Instance().GetValue("HardIP", "SplitCzFlow").EqualsIgnoreCase("TRUE");
            //_action = writeMessage;

            bool settingsNotFound = false;

            RunPreActionStep();
            RunBasicStep();
            RunEfuseStep();
            RunHardIpStep(checkSplitCzFlow);
            //Need to be generated after HardIP because DCTEST_IDS rule.
            RunRtosStep();
            RunScanStep();
            RunHtolStep();
            RunMbistStep();
            RunEvsStep();
            RunDvdcStep();
            RunBinCutStep();

            return settingsNotFound;
        }

        private static void RunPreActionStep()
        {
            #region PreAction
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.PreAction).Down)
            {
                Response.Report("Running Pre-ModuleMain~", EnumMessageLevel.CheckPoint);
                if (LocalSpecs.VoltageTbFileName.Count != 0)
                {
                    MergeSheet.ParseTestSettingSheetToTestPlan(LocalSpecs.VoltageTbFileName, EpWorkbook.TestPlanWorkbook);
                }
                using (var preActionMain = new PreActionMain())
                {
                    preActionMain.Execute(null);

                }
                Response.Report("Complete Pre-ModuleMain!", EnumMessageLevel.EndPoint);
                if (LocalSpecs.EquationVoltagesFileName != "N/A")
                {

                    EpWorkbook.EquationVoltages = new ExcelPackage(new FileInfo(LocalSpecs.EquationVoltagesFileName)).Workbook;
                    ExcelWorksheet worksheet = EpWorkbook.EquationVoltages.Worksheets["EquationVoltages"];
                    if (worksheet != null)
                    {
                        var nonIgxlsheet = new BinCutNonIgxlBase(worksheet, FolderStructure.DirCommon, !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder));
                        bool isCs = !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder);
                        TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, Path.GetFileNameWithoutExtension(nonIgxlsheet.WorkFlow(isCs)));
                    }

                }
            }
            #endregion
        }

        private static void RunBasicStep()
        {
            #region Basic
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Basic).Down)
            {
                Response.Report("Check Basic Settings~", EnumMessageLevel.CheckPoint);
                using (var basicMain = new BasicMainRf())
                {
                    basicMain.Execute(null!);
                }
                Response.Report("Check Basic Settings Completed!", EnumMessageLevel.EndPoint);
            }
            #endregion
        }

        private static void RunEfuseStep()
        {
            #region eFuse
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Efuse).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Efuse))
            {
                Response.Report("Check eFuse Settings~", EnumMessageLevel.CheckPoint);
                if (LocalSpecs.Options.Device == EnumDevice.AP || LocalSpecs.Options.Device == EnumDevice.RF)
                {
                    var eFuseMain = new EFuseMain();
                    eFuseMain.Execute(null);
                }

                Response.Report("eFuse Generation Completed!", EnumMessageLevel.EndPoint);
            }
            #endregion
        }

        private static void RunHardIpStep(bool checkSplitCzFlow)
        {
            #region HardIP
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.HardIp).Down && (LocalSpecs.IsModuleIncluded("IDS") || LocalSpecs.IsModuleIncluded(BlockStatus.HardIp)))
            {
                Response.Report("Check HardIP Settings~", EnumMessageLevel.CheckPoint);
                try
                {
                    var hardIpParaData = new HardIpParaData(EnumBlock.HardIp)
                    {
                        SplitCzFlow = checkSplitCzFlow
                    };
                    using var hardIpMain = new HardIpMain();
                    hardIpMain.Execute(hardIpParaData);
                }
                catch (Exception ex)
                {
                    Response.Report("HardIP Unexpected Failed: " + ex.Message, EnumMessageLevel.Error);
                }
                Response.Report("HardIP Generation Completed!", EnumMessageLevel.EndPoint);
            }
            #endregion
        }

        private static void RunRtosStep()
        {
            #region Rtos
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Rtos).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Rtos))
            {
                RunRtosCoreGeneration();
                RtosTableInstanceUpdater.UpdateRtosTableInfoInInstanceSheets();
                MoveRtosConfigurationSheetToTp();
            }
            #endregion
        }

        private static void RunRtosCoreGeneration()
        {
            Response.Report("Check Rtos Settings~", EnumMessageLevel.CheckPoint);

            {
                if (!TestProgram.IgxlWorkBk.SubFlowSheets.Any(p => p.Key.Contains("DCTEST_IDS", StringComparison.CurrentCulture)))
                {
                    var hardIpParaData = new HardIpParaData(EnumBlock.Rtos);
                    using var hardIpMain = new HardIpMain();
                    hardIpMain.Execute(hardIpParaData);
                }

                using (var spiRomMain = new SpiRomMain())
                {
                    spiRomMain.Execute(null);
                }
                Response.Report("Rtos Generation Completed!", EnumMessageLevel.EndPoint);

            }
        }

        private static void MoveRtosConfigurationSheetToTp()
        {
            #region Move Rtos_Configuration sheet to T/P

            if (EpWorkbook.TestPlanWorkbook.Worksheets.Any(x => x.Name.EqualsIgnoreCase("Rtos_Configuration")))
            {
                ExcelWorksheet rtosConfigSheet = EpWorkbook.TestPlanWorkbook.Worksheets.First(x => x.Name.EqualsIgnoreCase("Rtos_Configuration"));

                rtosConfigSheet.ExportWorkBook2Txt(FolderStructure.DirRtos);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirRtos, rtosConfigSheet.Name);
            }

            #endregion
        }

        private static void RunScanStep()
        {
            #region Scan
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Scan).Down && (LocalSpecs.IsModuleIncluded(BlockStatus.Scan) || LocalSpecs.IsModuleIncluded("CPM")))
            {
                Response.Report("Check SCAN Settings~", EnumMessageLevel.CheckPoint);
                using (var scanMain = new ScanMain())
                {
                    scanMain.Execute(null);
                }
                Response.Report("SCAN Generation Completed!", EnumMessageLevel.EndPoint);
            }
            #endregion
        }

        private static void RunHtolStep()
        {
            #region HTOL
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Htol).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Htol))
            {
                Response.Report("Check HTOL Settings~", EnumMessageLevel.CheckPoint);
                using (var htolMain = new HtolMain())
                {
                    htolMain.Execute(null);
                }
                Response.Report("HTOL Generation Completed!", EnumMessageLevel.EndPoint);
            }
            #endregion
        }

        private static void RunMbistStep()
        {
            #region Mbist
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Mbist).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Mbist))
            {
                Response.Report("Check MBIST Settings~", EnumMessageLevel.CheckPoint);
                using (var bistBiraMain = new BistBiraMain())
                {
                    bistBiraMain.Execute(null);
                }
                Response.Report("MBIST Generation Completed!", EnumMessageLevel.EndPoint);
            }
            #endregion
        }

        private static void RunEvsStep()
        {
            #region EVS
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Evs).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Evs))
            {
                Response.Report("Check EVS Settings~", EnumMessageLevel.CheckPoint);
                new EvsMain().Execute(null);
                Response.Report("EVS Generation Completed!", EnumMessageLevel.EndPoint);
            }
            #endregion
        }

        private static void RunDvdcStep()
        {
            #region DVDC
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Dvdc).Down || BlockStatus.GetAutomationBlockStatus(BlockStatus.Lcd).Down)
            {
                try
                {
                    var hardIPparaData = new HardIpParaData(EnumBlock.Dvdc);
                    using var dvdcMain = new DvdcMain1();
                    dvdcMain.Execute(hardIPparaData);
                }
                catch (Exception ex)
                {
                    Response.Report("Wireless Unexpected Failed: " + ex.Message, EnumMessageLevel.Error);
                }
                Response.Report("Wireless Generation Completed!", EnumMessageLevel.EndPoint);
            }
            #endregion
        }

        private static void RunBinCutStep()
        {
            #region BinCut

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
            #endregion
        }

        private static void RunAutomationPostAction()
        {
            PostActionParaData postActionParaData = PostActionMain.GetPostActionParas();
            using (var postActionMain = new PostActionMain())
            {
                postActionMain.Execute(postActionParaData);
                postActionMain.PrintFlow();
            }
            Response.Report("Post-ModuleMain Completed!", EnumMessageLevel.EndPoint);
        }
    }
}
