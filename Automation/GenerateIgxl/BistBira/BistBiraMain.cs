using System;
using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.BinCut.Business;
using Automation.GenerateIgxl.BistBira.Base;
using Automation.GenerateIgxl.BistBira.NewLogicData;
using Automation.GenerateIgxl.BistBira.NonLogicData;
using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Pattern;

using CommonLib.Enums;
using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using OfficeOpenXml;

using ScghLib.Base;
using ScghLib.Enums;
using ScghLib.Reader;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.BistBira
{
    public class BistBiraMain : WorkFlowBase<ParaData>
    {
        protected BistBiraInputData BistBiraInputData;
        protected BistIgxlResult IgxlResult = new BistIgxlResult();

        public BistBiraMain()
        {
            CharSetupSingleton.Instance();
            MultiTestSettingSheetsSingleton.Instance();
        }

        public override bool PreCheckFlow(ParaData paraData)
        {
            try
            {
                BistBiraInputData = new BistBiraInputManager(EpWorkbook.TestPlanWorkbook, paraData).Read();

                return true;
            }
            catch (Exception e)
            {
                Response.Report("BistBira has errors : " + e.StackTrace, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
                return false;
            }
        }

        /// <summary>
        /// Bist Work Flow
        /// Note for meeting in Mar.15 2016
        /// 1. Get performance mode from pattern name. The tool should use whatever the last performance mode was used. We will assume simple flow has no complicated branching and merging for performance modes
        /// 2. For flow handling, we will follow the flow strictly as described and will not assume any rules or consider the content in notes column
        /// 3. Vnom_plus is not supported. Will change to vMax for now. Future support will be considered4.
        /// 4. F0100 performance mode is not defined. This is for Vmargin and retention. Chris Vu will add that to performance mode definition. We may consider if we should handle it as a special case in the future
        /// 5. Rules for inserting a bin table into the flow: If a fail branch jumps to a row that has a fail action, bin table will be created. All patterns that jump to the same fail action will be included in the same bin table.If a fail branch jumps to the row with no fail action, a flag will be generated but no binning will be performed. All patterns that jump to the same fail action will be included in the flag.The fail is for customer's future use.
        /// 6. The bin cut in the new MBIST format is not yet confirmed. Myst team agrees this will be discussed and confirmed later
        /// 7. Add the Label into instance name combination
        /// </summary>
        /// <param name="paraData">para</param>
        public override void WorkFlow(ParaData paraData)
        {
            try
            {
                Response.Report("Creating MBISTFailBlock ...", percentage: 40);
                if (string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
                {
                    string file = Path.Combine(LocalSpecs.PatternFolder, "Log", "CSV_Bist_Info", $"{LocalSpecs.GetProjectNameMapping}_Bist_Info_All.csv");
                    if (!string.IsNullOrEmpty(LocalSpecs.MbistInfoFileName))
                    {
                        file = LocalSpecs.MbistInfoFileName;
                    }

                    // must put before BistAutoGen 0906
                    GenMbistFailBlock(LocalSpecs.PatternListCsvFileName, LocalSpecs.HardIpInfoFileName, file);
                }

                Response.Report("Mbist Auto Gen", percentage: 50);
                BistAutoGen(BistBiraInputData.ProdFlowSheets);

                Response.Report("Adding Mbist Sheets into Project Object ...", percentage: 70);
                AddDataToLocalSpec(BistBiraInputData.Naming);

                Response.Report("Mbist Completed.", percentage: 100);
            }
            catch (Exception e)
            {
                string message = "BistBira AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }

        protected virtual void BistAutoGen(List<BistProdFlowSheet> prodFlowSheets)
        {
            var mbistDataStore = new MbistDataStore();
            mbistDataStore.DicPatSets.Clear();

            foreach (BistProdFlowSheet sheet in prodFlowSheets)
            {
                MbistSheet mbistSheet = sheet.MbistSheet;
                Response.Report($"Generating for sheet {mbistSheet.SheetName} ...", percentage: 50);
                ExcelWorksheet worksheet = EpWorkbook.ScghWorkbook.Worksheets[mbistSheet.SheetName];
                if (worksheet == null)
                {
                    Response.Report($"Warning: Cannot find WorkSheet:{mbistSheet.SheetName} in Scgh File", EnumMessageLevel.Warning, 70);
                    continue;
                }
                BistProdFlowSheet prodFlowSheet = sheet.Copy();

                var bistBurst = new BistBurst(BistBiraInputData.Naming);
                bool bEnableBurst = mbistSheet.MbistPatSetType == MbistPatSetType.BurstYes || mbistSheet.MbistPatSetType == MbistPatSetType.BurstNo;
                if (bEnableBurst)
                {
                    prodFlowSheet = bistBurst.BurstLabel(prodFlowSheet, BistBiraInputData.PatternDic, mbistSheet.MbistPatSetType, mbistDataStore);
                }
                AddEqnVoltage(ref prodFlowSheet);
                prodFlowSheet.CreateLabelDic();

                AcSpecSheet acSpecSheet = TestProgram.IgxlWorkBk.GetAcSpecsSheet();
                if (acSpecSheet != null)
                {
                    new BinCutAcSpecsWriter().GenAcSpecs(prodFlowSheet.Rows, acSpecSheet);
                }

                var nonLogicalLib = new BistNonLogicalLib(BistBiraInputData.Config, MultiTestSettingSheetsSingleton.Instance(), BistBiraInputData.EquationVoltage);

                // burst flow
                if (bistBurst.HasPatSet && bEnableBurst && LocalSpecs.Options.Device != EnumDevice.RF)
                {
                    bistBurst.GenPatSetSheet(mbistSheet.SheetName, mbistDataStore);
                }
                GenFlowByFlowsheet(mbistDataStore, mbistSheet, prodFlowSheet, nonLogicalLib, mbistSheet.MbistPatSetType, mbistSheet.MbistBinTableType, mbistSheet.MbistLoop);
            }
        }

        private void AddEqnVoltage(ref BistProdFlowSheet prodFlowSheet)
        {
            if (BistBiraInputData.EquationVoltage == null)
            {
                return;
            }

            foreach (BistProdFlowRow row in prodFlowSheet.Rows)
            {
                if (row.Voltage.ContainsIgnoreCase("_EQN"))
                {
                    if (BistBiraInputData.EquationVoltage.Rows.Exists(x => x.PerformanceMode.Equals(row.Voltage, StringComparison.OrdinalIgnoreCase)))
                    {
                        row.EqnVoltage = BistBiraInputData.EquationVoltage.Rows.Find(x => x.PerformanceMode.Equals(row.Voltage, StringComparison.OrdinalIgnoreCase)).AllOther;

                    }
                }
            }
        }

        protected void GenFlowByFlowsheet(MbistDataStore mbistDataStore, MbistSheet sheetName, BistProdFlowSheet bistFlowsheet, BistNonLogicalLib nonLogicalLib, MbistPatSetType mbistPatSetType, MbistBinTableType mbistBinTableType, bool mbistLoop)
        {
            BistBiraInputData.PerformanceModeFilter.WorkFlow(bistFlowsheet);

            BistBiraInputData.VoltageConverter.WorkFlow(ref bistFlowsheet);

            nonLogicalLib.WorkFlow(bistFlowsheet, IgxlResult, mbistDataStore, mbistPatSetType);

            var mbistFlowGeneratorMain = new MbistFlowGenerator(bistFlowsheet, BistBiraInputData.Config, mbistBinTableType, mbistLoop);
            (SubFlowSheet flows, List<BinTableRow> needBinoutList) = mbistFlowGeneratorMain.GenerateSubflow();
            IgxlResult.FlowSheets.Add(flows);
            bool isNeedEvsDeferredBinout = TestPlanStatic.MainFlowSheet != null && TestPlanStatic.MainFlowSheet.EvsDeferSubFlowSheets.Contains(bistFlowsheet.SheetName);
            IgxlResult.BinTableRows.AddRange(mbistFlowGeneratorMain.CreateBinTable(needBinoutList, isNeedEvsDeferredBinout));
        }

        public static string GetFolder(string pStr, BistNaming naming)
        {
            if (string.Equals(pStr, naming.ModuleCpu, StringComparison.OrdinalIgnoreCase))
            {
                return FolderStructure.DirCpuMbist;
            }
            if (string.Equals(pStr, naming.ModuleGpu, StringComparison.OrdinalIgnoreCase))
            {
                return FolderStructure.DirGpuMbist;
            }
            if (string.Equals(pStr, naming.ModuleSoc, StringComparison.OrdinalIgnoreCase))
            {
                return FolderStructure.DirSocMbist;
            }
            return FolderStructure.DirMbist;
        }

        protected virtual void AddDataToLocalSpec(BistNaming naming)
        {
            string folder;
            foreach (SubFlowSheet flow in IgxlResult.FlowSheets)
            {
                if (flow == null)
                {
                    continue;
                }

                string name = flow.Name;
                if (name.StartsWith("Flow_", StringComparison.OrdinalIgnoreCase))
                {
                    name = name.Substring("Flow_".Length);
                }
                string module = naming.GetModule(name);
                folder = GetFolder(module, BistBiraInputData.Naming);
                flow.SourceInfo.Block = nameof(EnumBlock.Mbist);
                TestProgram.IgxlWorkBk.AddSubFlowSheet(folder, flow);
            }

            foreach (InstanceSheet inst in IgxlResult.InstanceSheets)
            {
                inst.AddRow(new InstanceRow());
                string name = inst.Name;
                if (name.StartsWith("TestInst_", StringComparison.OrdinalIgnoreCase))
                {
                    name = name.Substring("TestInst_".Length);
                }
                string module = naming.GetModule(name);
                folder = GetFolder(module, BistBiraInputData.Naming);
                TestProgram.IgxlWorkBk.AddInsSheet(folder, inst);
            }

            List<BinTableRow> rows = RemoveDuplicateBinTableRows(IgxlResult.BinTableRows);
            BinTableSheet binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet(FolderStructure.DirBinTable);
            foreach (BinTableRow binTableRow in rows)
            {
                binTable.AddRow(binTableRow);
            }
        }

        private List<BinTableRow> RemoveDuplicateBinTableRows(List<BinTableRow> binTableRows)
        {
            var result = new List<BinTableRow>();
            foreach (BinTableRow row in binTableRows)
            {
                if (!result.Exists(x => x.Name.Equals(row.Name, StringComparison.OrdinalIgnoreCase) && x.Items.Count.Equals(row.Items.Count)))
                {
                    result.Add(row);
                }
            }
            return result;
        }

        internal virtual void GenMbistFailBlock(string patternListCsvFileName, string hardipInfo, string bistInfo)
        {
            if (LocalSpecs.PatternListCsvFileName != "N/A")
            {
                var sheetList = new List<string>
                {
                    NeededSheets.MbistScghCpu, NeededSheets.MbistScghGpu, NeededSheets.MbistScghSoc
                };
                var mbistInfo = new PatMbistInfoManager(patternListCsvFileName, sheetList, LocalSpecs.IsUfp);
                string mbistInfoConfig = Path.Combine(LocalSpecs.SettingFolder, "Settings", "Pattern", "MbistInfo.xml");
                if (File.Exists(mbistInfoConfig))
                {
                    mbistInfo.MbistInfoCfgFile = mbistInfoConfig;
                    mbistInfo.ReadMbistConfig();
                }
                mbistInfo.OutputPath = FolderStructure.DirCommonSheets;
                mbistInfo.FileName = "MBISTFailBlock.txt";
                if (!mbistInfo.GetMbistInfoFromServer(hardipInfo, bistInfo))
                {
                    Response.Report("Missing Mbist Info ", EnumMessageLevel.Warning, 5);
                }
                else
                {
                    mbistInfo.SaveMbistInfo();
                    mbistInfo.SaveFingerPrintMaxDepth("FingerPrintMaxDepth.txt");
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, "MBISTFailBlock");
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, "MBISTFailBlock_ModuleNameOnly");
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, "FingerPrintMaxDepth");
                }
            }
        }
    }
}
