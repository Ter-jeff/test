using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.GenerateIgxl.PreAction.AddPinGrp;
using Automation.GenerateIgxl.PreAction.CreateJobMap;
using Automation.GenerateIgxl.PreAction.GenChannelMap;
using Automation.GenerateIgxl.PreAction.GenEmptyBinTbl;
using Automation.GenerateIgxl.PreAction.GenPinMap;
using Automation.GenerateIgxl.PreAction.InitMainFlow;
using Automation.GenerateIgxl.PreAction.ReadBasLib;
using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.PreCheck.PreCheckManager;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Pattern;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlSheets;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.PreAction
{
    public class PreActionMain : WorkFlowBase<ParaData>
    {
        private PreActionInputData _preActionInputData;

        public override bool PreCheckFlow(ParaData paraData)
        {
            try
            {
                _preActionInputData = new PreActionInputManager(EpWorkbook.TestPlanWorkbook).Read();

                if (!LocalSpecs.Options.BypassPreCheck)
                {
                    new PreActionCheckManager(EpWorkbook.TestPlanWorkbook, paraData, _preActionInputData).PreCheckAll();
                }

                return true;
            }
            catch (Exception e)
            {
                Response.Report("Pre Action has errors : " + e.StackTrace, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
                return false;
            }
        }

        public override void WorkFlow(ParaData paraData)
        {
            try
            {
                Response.Report("Initializing Pre-ModuleMain ...", percentage: 10);

                ExceptionListSingleton.Instance().Initialize();

                UpdateHardIpInfo();

                Response.Report("Initializing vbt functions ...", percentage: 20);
                var basMain = new BasMain();
                (List<Function> functions, ReferenceSheet reference) = basMain.WorkFlow(FolderStructure.DirLib);
                TestProgram.VbtFunctionLib.AddVbtFunctionRange(functions);
                if (reference != null)
                {
                    TestProgram.IgxlWorkBk.AddReferenceSheet(FolderStructure.DirReference, reference);
                }

                Response.Report("Creating JobMapping ...", percentage: 25);
                var jobMapMain = new JobMapMain();
                jobMapMain.WorkFlow();

                Response.Report("Creating PinMap Object ...", percentage: 40);
                var pinMapMain = new PinMapMain();
                PinMapSheet sheet = pinMapMain.WorkFlow(_preActionInputData.PinMapSheet, _preActionInputData.PinGroupSheet, _preActionInputData.IoContinuity);

                if (sheet != null)
                {
                    TestProgram.IgxlWorkBk.PinMapPair = new KeyValuePair<string, PinMapSheet>(FolderStructure.DirPinMap, sheet);
                }
                else
                {
                    sheet = new PinMapSheet("PinMap");
                    TestProgram.IgxlWorkBk.PinMapPair = new KeyValuePair<string, PinMapSheet>(FolderStructure.DirPinMap, sheet);
                }

                #region channelMapSheet
                string extraFolder = Path.Combine(LocalSpecs.SettingFolder, "Settings", "ExtraSheets");
                if (Directory.Exists(extraFolder))
                {
                    var channelMapMain = new ChannelMapMain(LocalSpecs.Options.Device.ToString());
                    channelMapMain.WorkFlow(extraFolder);
                    channelMapMain.ModifyPinMapByChannelMap();
                }

                if (_preActionInputData.ChannelMapSheets != null)
                {
                    Response.Report("Creating ChannelMap ...", percentage: 50);
                    var channelMapMain = new ChannelMapMain(LocalSpecs.Options.Device.ToString());
                    channelMapMain.WorkFlow(_preActionInputData.ChannelMapSheets);
                    channelMapMain.ModifyPinMapByChannelMap();
                }
                #endregion

                Response.Report("Creating special Pin Group by TestSettingSheet ...", percentage: 60);
                var testSettings = MultiTestSettingSheetsSingleton.Instance();
                var specialPinGrpPlus = new SpecialPinGrpPlus();
                specialPinGrpPlus.WorkFlow(testSettings);

                Response.Report("Creating empty BinTable ...", percentage: 60);
                var emptyBinTblMain = new EmptyBinTblMain();
                emptyBinTblMain.WorkFlow();

                var initMainFlow = new GenInitMainFlow();
                initMainFlow.WorkFlow(true);

                GenSelsramMappingTable();

                PreCheckHardIpInfo();

                Response.Report("Pre-ModuleMain Completed!", percentage: 100);
            }
            catch (Exception e)
            {
                string message = "Pre-action AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }

        private void GenSelsramMappingTable()
        {
            ExcelWorksheet sheet = EpWorkbook.TestPlanWorkbook.Worksheets["SELSRM_Mapping_Table"] ?? EpWorkbook.TestPlanWorkbook.Worksheets["SELSRAM_Mapping_Table"];
            if (sheet != null)
            {
                Response.Report("Generating SELSRM Mapping Table ...");
                sheet.ExportWorkBook2Txt(FolderStructure.DirCommon);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, sheet.Name);
            }
        }

        private void PreCheckHardIpInfo()
        {
            if (LocalSpecs.HardIpInfos == null)
            {
                return;
            }

            var patInfoChecker = new PatInfoChecker();
            patInfoChecker.CheckPatInfoAll(LocalSpecs.HardIpInfos.SelectMany(x => x.Value).ToList());
        }

        private void UpdateHardIpInfo()
        {
            if (LocalSpecs.PatternListCsvFileName != "N/A")
            {
                // Add generate hardip_info.log here
                Response.Report("Start Generating HardIp Info ...");
                var myUpdate = new UpdateHardIpInfo(LocalSpecs.CurrentProject);
                var fileInfo = new FileInfo(LocalSpecs.PatternListCsvFileName);
                var patListCsv = new InputPatternListCsv(fileInfo);
                myUpdate.UpdateHardIP_Info(patListCsv.GetPatternTimeSet(LocalSpecs.TimeSetFolder), patListCsv.FullName);
                string hardIpInfoFile = Path.Combine(LocalSpecs.SettingFolder, "Settings", "HardIP", $"HardIP_PatInfo_{LocalSpecs.CurrentProject}.log");
                if (File.Exists(hardIpInfoFile))
                {
                    List<HardIpInfo> patInfo = new PatInfoReader().ExtractHardIpInfos(hardIpInfoFile);
                    LocalSpecs.HardIpInfos = new HardIpInfos(patInfo);
                }
            }
        }
    }
}
