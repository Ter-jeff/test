using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.InputChecker;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LogLib.Utility;

using OfficeOpenXml;

using TestPlanLib.Efuse.Input;

namespace Automation.GenerateIgxl.EFuse.Business
{
    public class EfuseGenerate
    {
        protected readonly EFuseInputData EFuseInputData;
        public List<BitDefTable> FinalBitDefTables = new List<BitDefTable>();
        public Dictionary<string, List<string>> PowerPinDic;
        public List<string> EfuseConfigDataEnableWd = new List<string>();
        public bool HasMultipleBlock = false;
        public EfuseGenerate(EFuseInputData eFuseInputData)
        {
            EFuseInputData = eFuseInputData;
            PowerPinDic = new EFusePreProcess().CollectEachBankPowerPin(EFuseInputData.EfuseArraySizeSheet);
        }

        public virtual void WorkFlow()
        {
            //Filter out SLT job and bira bank
            List<BitDefTable> bitDefTables = new EFusePreProcess().FilterBankForBitDefTableByArraySizeSheet(EFuseInputData.EfuseBitDefTables.Select(x => x.Copy()).ToList(), EFuseInputData.EfuseArraySizeSheet);

            //Algorithm cond
            FinalBitDefTables = AddConfigRowToBitDefTable(bitDefTables);

            //Add dummy fuse for bank_mon
            FinalBitDefTables = AddDummyFuseForBankMon(FinalBitDefTables);

            //Generate EFUSE_BitDef_Table
            WriteEFuseBitDef();
            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, EFuseConst.BitDefTableFileName);

            //Generate Config_table and Config_table_SVM

            bool hasSvmKey = CheckIfSvmKeyWord();
            CheckIfHasBkmProcess();

            if (EFuseInputData.EfuseConfigMainSheets.Count > 0)
            {
                List<BitDefTable> configTablesInBitDef = GetConfigRowInBitDefTables();
                WriteConfigTable(configTablesInBitDef, hasSvmKey);
            }
            //Generate BKM
            GenerateBkmInfoTable();

            #region Generate flow sheet, Instance sheet, BinTable and pattern set
            bool hasEarly = CheckIfEarlyConfig();
            EfuseFinalInstanceRows efuseFinalInstRows = TestPlanStatic.EfuseInstanceSheets != null && TestPlanStatic.EfuseInstanceSheets.Any() ? new EFusePreProcess().GetEfuseInstanceRows(TestPlanStatic.EfuseInstanceSheets, EFuseInputData.EfusePatternRows) : new EFusePreProcess().GetEfuseInstanceRows(EFuseInputData.EfusePatternRows, EFuseInputData.EfuseReadSheet == null ? new List<EfuseReadRow>() : EFuseInputData.EfuseReadSheet.Rows, hasEarly);

            #region Check unused pattern
            if (EFuseInputData.EfuseScghSheet != null)
            {
                new EfusePatternMapChecker().CheckUnusedPatInScgh(EFuseInputData.EfuseScghSheet, efuseFinalInstRows, EFuseInputData.EfusePatternRows);
            }

            #endregion

            if (TestPlanStatic.EfuseInstanceSheets != null && TestPlanStatic.EfuseInstanceSheets.Any())
            {
                GenerateEfuseSheetsByInstanceSheet(efuseFinalInstRows);
            }

            #endregion

            //Generate ECID_Check_List
            GenerateEcidCheckList();

            if (hasSvmKey)
            {
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, EFuseConst.ConfigTableSvmFileName);
            }

            //Generate FuseCheckTable
            ExcelWorksheet fuseCheckTable = EpWorkbook.TestPlanWorkbook.Worksheets["FuseCheckTable"] ?? EpWorkbook.TestPlanWorkbook.Worksheets["FuseCheckTable"];
            if (fuseCheckTable != null)
            {
                fuseCheckTable.ExportWorkBook2Txt(FolderStructure.DirCommon);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, fuseCheckTable.Name);
            }

            ExcelWorksheet otherLookUpTable = EpWorkbook.TestPlanWorkbook.Worksheets["Other_Lookup_Tables"] ?? EpWorkbook.TestPlanWorkbook.Worksheets["Other_Lookup_Tables"];
            if (otherLookUpTable != null)
            {
                otherLookUpTable.ExportWorkBook2Txt(FolderStructure.DirCommon);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, otherLookUpTable.Name);
            }
        }

        private void GenerateEfuseSheetsByInstanceSheet(EfuseFinalInstanceRows efuseFinalInstRows)
        {
            var flagList = new List<string>();
            var flowGroups = efuseFinalInstRows.GroupBy(x => x.EfuseInstanceRow.FlowName).ToList();
            var efuseInstanceGenerator = new EfuseGenerateInstanceCs();
            var instanceSheet = new InstanceSheet(EFuseConst.TestInstSheet);
            var flowSheets = new List<SubFlowSheet>();
            var flowList = flowGroups.Select(x => x.Key).Distinct().ToList();
            foreach (IGrouping<string, EfuseFinalInstanceRow> flowGroup in flowGroups)
            {
                foreach (EfuseFinalInstanceRow instance in flowGroup)
                {
                    if (instance.EfuseInstanceRow.IsCallFlow || instance.EfuseInstanceRow.IsOpcode)
                    {
                        continue;
                    }

                    instanceSheet.Rows.Add(efuseInstanceGenerator.GenerateInstanceRowByInstanceSheet(instance));
                }
                instanceSheet.AddHeaderFooter(flowGroup.Key);
                flowSheets.Add(GenerateSubFlowByInstanceSheet(flowGroup, ref flagList, flowList));
            }
            IEnumerable<string> enables = efuseFinalInstRows.Select(x => x.EfuseInstanceRow.EnableFlow.Replace("!", "")).ToList().Distinct();
            foreach (string enable in enables)
            {
                if (!string.IsNullOrEmpty(enable))
                {
                    TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd(enable, OpCode.Nop, EFuseConst.Efuse, FolderStructure.DirMain, "");
                }
            }
            if (TestPlanStatic.HarvestingTruthTableSheets != null && TestPlanStatic.HarvestingTruthTableSheets.Any())
            {
                List<string> binCheckInstanceList = efuseInstanceGenerator.GenerateBinCheckInstance(instanceSheet);
                flowSheets.Add(GenerateBinCheckFlow(binCheckInstanceList));
                flagList.Add("F_EFUSE_BinCheck");
            }
            GenEfuseInstanceSheet(instanceSheet);
            GetPatSetSheet(efuseFinalInstRows);
            flagList = flagList.Distinct().ToList();
            GenBinTable(flagList);
            foreach (SubFlowSheet flow in flowSheets)
            {
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirEfuse, flow);
            }
        }

        private void GenerateEcidCheckList()
        {
            var ecidCheckListTemplate = new FileInfo(Path.Combine(FolderStructure.DirIgLink, "ECID_Check_List_Template.xlsx"));
            List<string> header = new List<string> { "Fab_Lot_ID", "WaferID", "X", "Y", "HW_Bin" };
            List<string> example = new List<string> { "U9S536", "4", "1", "7", "10" };
            using (var ep = new ExcelPackage())
            {
                string sheetName = "ECID_Check_List";
                ep.Workbook.Worksheets.Add(sheetName);
                ExcelWorksheet ws = ep.Workbook.Worksheets[sheetName];

                int col = 1;
                foreach (string item in header)
                {
                    ws.Cells[1, col].Value = item;
                    ws.Cells[3, col].Value = "END";
                    col++;
                }

                col = 1;
                foreach (string item in example)
                {
                    ws.Cells[2, col].Value = item;
                    col++;
                }

                ep.SaveAs(ecidCheckListTemplate);
            }
        }

        protected void GenerateBkmInfoTable()
        {
            if (EFuseInputData.EfuseBkmInfoTable != null)
            {
                EFuseInputData.EfuseBkmInfoTable.Generate(FolderStructure.DirCommonSheets, EFuseConst.BkmName);

                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, EFuseConst.BkmName);
            }
        }

        private void GenBinTable(List<string> efuseFlags)
        {
            var binCutBinTableWriter = new EfuseBinTableWriter();
            binCutBinTableWriter.GenBinTable(efuseFlags);
        }

        private void GetPatSetSheet(List<EfuseFinalInstanceRow> efuseInstRows)
        {
            if (!efuseInstRows.Any())
            {
                return;
            }

            PatSetSheet patSetSheet = new EfusePatSetWriter().GetPatSetSheet(efuseInstRows);
            if (patSetSheet.Rows.Any())
            {
                TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirEfuse, patSetSheet);
            }
        }

        protected void GenEfuseInstanceSheet(InstanceSheet instanceSheet)
        {
            if (instanceSheet.Rows.Any())
            {
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirEfuse, instanceSheet);
            }
        }

        internal bool CheckIfEarlyConfig()
        {
            return FinalBitDefTables.Exists(x => EFuseConst.GetBankName(x.BlockName).Equals(BankType.Cfg) &&
                                                         x.ProgrammingStageIdx != -1 && x.Rows.Any(y =>
                                                         y.RowData[x.ProgrammingStageIdx].EndsWith(EFuseConst.Early, StringComparison.CurrentCultureIgnoreCase)));
        }

        private SubFlowSheet GenerateSubFlowByInstanceSheet(IGrouping<string, EfuseFinalInstanceRow> instanceSheetRow, ref List<string> flagList, List<string> flowList)
        {
            string sheetName = instanceSheetRow.Key;
            string subFlow = Regex.Replace(sheetName, "^Flow_efuse_", "", RegexOptions.IgnoreCase);
            string testPlanSheet = instanceSheetRow.FirstOrDefault()?.EfuseInstanceRow.SheetName;
            if (!sheetName.StartsWith("Flow_", StringComparison.CurrentCultureIgnoreCase))
            {
                sheetName = "Flow_" + sheetName;
            }

            var flowSheet = new SubFlowSheet(sheetName, testPlanSheet + ":" + sheetName);
            flowSheet.AddPrintStartRow(sheetName);
            if (instanceSheetRow.Select(x => x.EfuseInstanceRow).ToList().Where(x => !string.IsNullOrEmpty(x.JobTestStage))
                    .Select(x => x.JobTestStage.Replace("!", "")).Distinct().Count() == 1)
            {
                string jobName = instanceSheetRow.Select(x => x.EfuseInstanceRow).Select(x => x.JobTestStage).First();
                if (!jobName.Contains(","))
                {
                    flowSheet.JobNames = new List<string> { jobName };
                }
            }
            string bankFlag = "F_" + Combination.CombineByUnderLine(subFlow, EFuseConst.BankEnable);
            flowSheet.AddClearFlag(bankFlag);
            flowSheet.AddFlagToTrue(bankFlag, Combination.CombineByUnderLine(EFuseConst.BankEnable, sheetName));
            flowSheet.AddRow(new FlowRow { Opcode = OpCode.If, Parameter = bankFlag });
            string defaultFailFlag = "F_" + Combination.CombineByUnderLine(subFlow, EFuseConst.BankFail);
            foreach (EfuseFinalInstanceRow flowRow in instanceSheetRow)
            {
                if (flowRow.EfuseInstanceRow.IsCallFlow && (flowList.Contains(flowRow.EfuseInstanceRow.Instance) || flowRow.EfuseInstanceRow.Instance.Equals("Flow_efuse_BinCheck", StringComparison.CurrentCultureIgnoreCase)))
                {
                    flowSheet.AddRow(new FlowRow { Enable = flowRow.EfuseInstanceRow.EnableFlow, Job = flowRow.EfuseInstanceRow.JobTestStage, Opcode = OpCode.Call, Parameter = flowRow.EfuseInstanceRow.Instance });
                }
                else
                {
                    if (flowRow.EfuseInstanceRow.IsOpcode)
                    {
                        flowSheet.AddRow(new FlowRow { Enable = flowRow.EfuseInstanceRow.EnableFlow, Job = flowRow.EfuseInstanceRow.JobTestStage, Opcode = flowRow.EfuseInstanceRow.Opcode, Parameter = flowRow.EfuseInstanceRow.Instance });
                    }
                    else
                    {
                        flowSheet.AddTestAndBinTableRow(flowRow.EfuseInstanceRow.Instance, flowRow.EfuseInstanceRow.JobTestStage, flowRow.SiteVar.SiteVarCondition, flowRow.SiteVar.SiteVarValue, string.IsNullOrEmpty(flowRow.EfuseInstanceRow.FailFlag) ? defaultFailFlag : flowRow.EfuseInstanceRow.FailFlag, flowRow.EfuseInstanceRow.IsBypassBinOut, flowRow.EfuseInstanceRow.FailFlag.ContainsIgnoreCase("BANKREAD"), flowRow.EfuseInstanceRow.EnableFlow, flowRow.NopByEnableWord, flowRow.EfuseInstanceRow.BintableEnableWd, flowRow.EfuseInstanceRow.BinOutStage);
                    }
                    flagList.Add(string.IsNullOrEmpty(flowRow.EfuseInstanceRow.FailFlag) ? defaultFailFlag : flowRow.EfuseInstanceRow.FailFlag);
                }
            }
            flowSheet.AddRow(new FlowRow { Opcode = OpCode.EndIf });
            flowSheet.AddPrintEndRow(sheetName);
            flowSheet.AddReturnRow();
            return flowSheet;
        }

        internal SubFlowSheet GenerateBinCheckFlow(List<string> instanceList)
        {
            var flowSheet = new SubFlowSheet("Flow_efuse_BinCheck", "Instance_EFUSE:Flow_efuse_BinCheck");
            flowSheet.AddPrintStartRow("Flow_efuse_BinCheck");
            foreach (string instance in instanceList)
            {
                string binNum = instance.Split('_').Last();
                flowSheet.AddRow(new FlowRow { Opcode = OpCode.Test, Parameter = instance, FailAction = "F_EFUSE_BinCheck", DeviceCondition = "flag-true", DeviceName = binNum });
            }
            flowSheet.AddRow(new FlowRow { Opcode = OpCode.BinTable, Parameter = "Bin_EFUSE_BinCheck" });
            flowSheet.AddPrintEndRow("Flow_efuse_BinCheck");
            flowSheet.AddReturnRow();

            return flowSheet;
        }

        protected internal List<BitDefTable> GetConfigRowInBitDefTables()
        {
            var configTables = new List<BitDefTable>();
            foreach (BitDefTable table in FinalBitDefTables)
            {
                var configTable = new BitDefTable();
                int lsbIdx = table.LsbBitIdx;
                int msbIdx = table.MsbBitIdx;
                int algorithmIdx = table.AlgorithmIdx;
                if (lsbIdx != -1 && msbIdx != -1 && algorithmIdx != -1)
                {
                    configTable.Rows.AddRange(table.Rows.Where(x => x.RowData[algorithmIdx].Equals("cond", StringComparison.CurrentCultureIgnoreCase)));
                }
                int defaultOrRealIdx = table.DefaultOrRealIdx;
                int defaultValueIdx = table.DefaultValueIdx;
                if (defaultOrRealIdx != -1 && defaultValueIdx != -1)
                {
                    configTable.Rows.AddRange(
                        table.Rows.Where(
                            x =>
                                x.RowData[defaultOrRealIdx].Equals("default", StringComparison.CurrentCultureIgnoreCase) &&
                                x.RowData[defaultValueIdx].Equals("NA", StringComparison.CurrentCultureIgnoreCase)));
                }
                if (configTable.Rows.Any())
                {
                    if (table.BlockName.ContainsIgnoreCase("config_"))
                    {
                        configTable.BlockName = table.BlockName.Split('_')[1];
                    }
                    configTable.LsbBitIdx = lsbIdx;
                    configTable.MsbBitIdx = msbIdx;
                    configTables.Add(configTable);
                }
            }
            return configTables;
        }

        private void WriteConfigTable(List<BitDefTable> configTablesInBitDef, bool hasSvmKey)
        {
            var blockNames = configTablesInBitDef
                .Where(t => !string.IsNullOrWhiteSpace(t.BlockName))
                .Select(t => t.BlockName)
                .Distinct()
                .ToList();

            List<string> blocksToProcess = blockNames.Any() ? blockNames : new List<string> { null };
            HasMultipleBlock = blocksToProcess.Count > 1;
            foreach (string blockName in blocksToProcess)
            {
                string xlsPath = "";
                string txtPath = "";
                string configFileName = "";
                if (!string.IsNullOrEmpty(blockName))
                {

                    configFileName = EFuseConst.ConfigTableFileName.Replace("_", $"_{blockName}_");
                    xlsPath = Path.Combine(FolderStructure.DirCommonSheets, configFileName + ".xlsx");
                    txtPath = Path.Combine(FolderStructure.DirCommonSheets, configFileName + ".txt");
                }
                else
                {
                    configFileName = EFuseConst.ConfigTableFileName;
                    xlsPath = Path.Combine(FolderStructure.DirCommonSheets, configFileName + ".xlsx");
                    txtPath = Path.Combine(FolderStructure.DirCommonSheets, configFileName + ".txt");
                }

                if (File.Exists(txtPath))
                {
                    File.Delete(txtPath);
                }

                //print Config_table
                var svmDisableKey = new Regex(@"Follow Tracker sheet : if SVM disable fuse ""(?<svmValue>\S+)""", RegexOptions.IgnoreCase);
                WriteConfigRowData(xlsPath, txtPath, svmDisableKey, hasSvmKey, configTablesInBitDef, blockName);

                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, configFileName);
                //print Config_table_SVM
                if (hasSvmKey)
                {
                    xlsPath = Path.Combine(FolderStructure.DirCommonSheets, EFuseConst.ConfigTableSvmFileName + ".xlsx");
                    txtPath = Path.Combine(FolderStructure.DirCommonSheets, EFuseConst.ConfigTableSvmFileName + ".txt");

                    if (File.Exists(txtPath))
                    {
                        File.Delete(txtPath);
                    }

                    var svmEnableKey = new Regex(@"Follow Tracker sheet : if SVM enable fuse ""(?<svmValue>\S+)""", RegexOptions.IgnoreCase);
                    WriteConfigRowData(xlsPath, txtPath, svmEnableKey, true, configTablesInBitDef);
                }
            }
        }

        public static List<EfuseConfigMainSheet> FilterEfuseConfigSheets(IEnumerable<EfuseConfigMainSheet> source, string configType, bool hasMultipleBlock)
        {
            if (source == null)
            {
                return new List<EfuseConfigMainSheet>();
            }

            if (!string.IsNullOrEmpty(configType) && hasMultipleBlock)
            {
                return source
                    .Where(sheet =>
                        sheet?.SheetName != null &&
                        sheet.SheetName.ContainsIgnoreCase(configType))
                    .ToList();
            }

            return source.ToList();
        }

        protected void WriteConfigRowData(string xlsPath, string txtPath, Regex svmRegexKey, bool hasSvmKey, List<BitDefTable> configTablesInBitDef, string configType = null)
        {
            var excel = new ExcelPackage(new FileInfo(xlsPath));
            ExcelWorksheet configSheet = excel.Workbook.Worksheets.Add(EFuseConst.ConfigTableFileName);

            const int currCol = 1;
            int offsetCol = 0;

            //Recognize config is same or not
            CheckAllConfigIsSame();

            List<EfuseConfigMainSheet> efuseConfigSheets = FilterEfuseConfigSheets(
                EFuseInputData.EfuseConfigMainSheets,
                configType,
                HasMultipleBlock);

            foreach (EfuseConfigMainSheet sheet in efuseConfigSheets)
            {
                bool isFirstSheet = sheet.SheetName.Equals(efuseConfigSheets.First().SheetName);

                int currRow = 1;
                int dataCount = sheet.DataColCount;
                //print sheet name
                currRow = configSheet.Cells[currRow, currCol + offsetCol].PrintExcelRow(new[] { sheet.SheetName });
                //print title
                string title = "Condition,MSB BIT,LSB BIT,Bit Width,programming stage,";
                List<string> dataTitle = sheet.Header.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries).ToList().GetRange(sheet.DataStartCol - 1, dataCount);
                //The record data header is to initialize the enable word to the Flow_Table_Main_Init_EnableWd sheet
                EfuseConfigDataEnableWd.AddRange(dataTitle);
                if (isFirstSheet)
                {
                    title = title + string.Join(",", dataTitle) + ",Comment" + Regex.Replace(dataTitle.First(), @"\d+", "");
                }
                else
                {
                    title = string.Join(",", dataTitle) + ",Comment" + Regex.Replace(dataTitle.First(), @"\d+", "");
                }

                currRow = configSheet.Cells[currRow, currCol + offsetCol].PrintExcelRow(title.Split(','));

                //print data
                var configRow = new List<List<object>>();
                foreach (EfuseConfigMainRow row in sheet.Rows)
                {
                    if ((!IsInBitDefConfigBitRange(row, configTablesInBitDef) && row.IsAllSameData) || row.ValueToFuse.ContainsIgnoreCase("real"))
                    {
                        continue;
                    }

                    string programAt = row.FuseBlowLocation;
                    if (!string.IsNullOrEmpty(row.FuseBlowLocation) && row.FuseBlowLocation.ContainsIgnoreCase("FTF"))
                    {
                        programAt = Regex.Replace(row.FuseBlowLocation, "FTF", "FT3", RegexOptions.IgnoreCase);
                    }

                    if (hasSvmKey && svmRegexKey.IsMatch(row.ValueToFuse))
                    {
                        string svmValue = svmRegexKey.Match(row.ValueToFuse).Groups["svmValue"].ToString();
                        if (string.IsNullOrEmpty(svmValue))
                        {
                            svmValue = "SVM";
                        }

                        if (isFirstSheet)
                        {
                            configRow.Add(new List<object>
                            {
                                row.Description, row.Msb, row.Lsb, row.Width, programAt,
                                svmValue,svmValue,svmValue,svmValue,svmValue,svmValue,svmValue,svmValue,svmValue,svmValue,
                                row.ValueToFuse
                            });
                        }
                        else
                        {
                            configRow.Add(new List<object>
                            {
                                svmValue,svmValue,svmValue,svmValue,svmValue,svmValue,svmValue,svmValue,svmValue,svmValue,
                                row.ValueToFuse
                            });
                        }
                    }
                    else
                    {
                        var cellData = new List<object>();
                        if (isFirstSheet)
                        {
                            cellData = new List<object>
                            {
                                row.Description,
                                row.Msb,
                                row.Lsb,
                                row.Width,
                                programAt
                            };
                        }

                        cellData.AddRange(row.DataList);
                        cellData.Add(row.ValueToFuse);
                        configRow.Add(cellData);
                    }
                }

                configSheet.Cells[currRow, currCol + offsetCol].PrintExcelRowByList(configRow);
                offsetCol += title.Split(',').Length;
                if (sheet.SheetName.Equals(efuseConfigSheets.Last().SheetName))
                {
                    configSheet.Cells[2, currCol + offsetCol].PrintExcelRow(new[] { EFuseConst.End });
                }
            }
            configSheet.ExportSheet(txtPath);
        }

        private void CheckAllConfigIsSame()
        {
            if (EFuseInputData.EfuseConfigMainSheets.Any())
            {
                foreach (EfuseConfigMainRow row in EFuseInputData.EfuseConfigMainSheets.First().Rows)
                {
                    if (row.Description.Equals("kis_to_pmu_reset_secured", StringComparison.OrdinalIgnoreCase))
                    {
                    }
                    bool isSame = row.IsAllSameData;
                    foreach (EfuseConfigMainSheet sheet in EFuseInputData.EfuseConfigMainSheets)
                    {
                        if (sheet.Rows.Find(r => r.Description.Equals(row.Description)) != null)
                        {
                            isSame = isSame && sheet.Rows.Find(r => r.Description.Equals(row.Description)).IsAllSameData;
                        }
                    }
                    //update flag back
                    foreach (EfuseConfigMainSheet sheet in EFuseInputData.EfuseConfigMainSheets)
                    {
                        if (sheet.Rows.Find(r => r.Description.Equals(row.Description)) != null)
                        {
                            sheet.Rows.Find(r => r.Description.Equals(row.Description)).IsAllSameData = isSame;
                        }
                    }
                }
            }
        }

        internal bool IsInBitDefConfigBitRange(EfuseConfigMainRow currRow, List<BitDefTable> configTablesInBitDef)
        {
            try
            {
                int msbInt = currRow.Msb;
                int lsbInt = currRow.Lsb;
                if (
                    configTablesInBitDef.Any(table => table.Rows.Exists(row =>
                                                        lsbInt >= int.Parse(row.RowData[table.LsbBitIdx]) &&
                                                        lsbInt <= int.Parse(row.RowData[table.MsbBitIdx]) &&
                                                        msbInt >= int.Parse(row.RowData[table.LsbBitIdx]) &&
                                                        msbInt <= int.Parse(row.RowData[table.MsbBitIdx]))))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
            return false;
        }

        protected internal bool CheckIfSvmKeyWord()
        {
            return EFuseInputData.EfuseConfigMainSheets.Any(x => x.HasSvm);
        }

        protected internal void CheckIfHasBkmProcess()
        {
            bool hasBkmProcess = FinalBitDefTables.Any(x => x.Rows.Any(y => y.RowData[0].Equals("bkm_process", StringComparison.CurrentCultureIgnoreCase)));
            LocalSpecs.HasBkmProcess = hasBkmProcess;
        }

        private void WriteEFuseBitDef()
        {
            string fileNameXls = Path.Combine(FolderStructure.DirCommonSheets, EFuseConst.BitDefTableFileName + ".xlsx");
            string fileNameTxt = Path.Combine(FolderStructure.DirCommonSheets, EFuseConst.BitDefTableFileName + ".txt");

            WriteEFuseBitDef(fileNameXls, fileNameTxt);
        }

        protected void WriteEFuseBitDef(string lStrFileNameXls, string lStrFileNameTxt)
        {
            if (File.Exists(lStrFileNameTxt))
            {
                File.Delete(lStrFileNameTxt);
            }

            if (!Directory.Exists(Path.GetDirectoryName(lStrFileNameTxt)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(lStrFileNameTxt) ?? string.Empty);
            }

            var excel = new ExcelPackage(new FileInfo(lStrFileNameXls));
            ExcelWorksheet bitDefSheet = excel.Workbook.Worksheets.Add(EFuseConst.BitDefTableFileName);
            int currRow = 1;
            foreach (BitDefTable bitDefTable in FinalBitDefTables)
            {
                //print access mode
                bitDefSheet.Cells[currRow, 1].Value = bitDefTable.AccessMode;
                //print title
                currRow = bitDefSheet.Cells[currRow + 1, 1].PrintExcelRow(bitDefTable.Header.Split('\t'));
                //print data
                foreach (BitDefRow bitDefRow in bitDefTable.Rows)
                {
                    for (int idx = 0; idx < bitDefRow.RowData.Count; idx++)
                    {
                        bitDefSheet.Cells[currRow, idx + 1].PrintExcelRow(new[] { bitDefRow.RowData[idx] });
                    }

                    currRow++;
                }
            }
            bitDefSheet.Cells[currRow, 1].Value = EFuseConst.End;
            try
            {
                bitDefSheet.ExportSheet(lStrFileNameTxt);
            }
            catch (Exception e)
            {
                throw new Exception("Write error report failed. " + e.StackTrace);
            }
        }

        protected List<BitDefTable> AddConfigRowToBitDefTable(List<BitDefTable> bitDefTables)
        {
            var finalBitDefTables = new List<BitDefTable>();

            foreach (BitDefTable bitDefTable in bitDefTables)
            {
                var currRows = new List<BitDefRow>();
                currRows.AddRange(bitDefTable.Rows);
                bitDefTable.Rows.Clear();
                string[] parts = (bitDefTable.BlockName ?? string.Empty).Split('_');
                string blockKey = parts.Length > 1 ? parts[1] : string.Empty;
                foreach (BitDefRow bitDefRow in currRows)
                {
                    if (bitDefRow.RowData[bitDefTable.AlgorithmIdx].Equals("cond", StringComparison.CurrentCultureIgnoreCase))
                    {
                        AddCondConfigRows(bitDefTable, bitDefRow, blockKey);
                    }
                    else
                    {
                        CheckDatabaseRevision(bitDefTable, bitDefRow);
                        bitDefTable.Rows.Add(bitDefRow);
                    }
                }
                finalBitDefTables.Add(bitDefTable);
            }
            return finalBitDefTables;
        }

        private void AddCondConfigRows(BitDefTable bitDefTable, BitDefRow bitDefRow, string blockKey)
        {
            double msb = Convert.ToDouble(bitDefRow.RowData[bitDefTable.MsbBitIdx]);
            double lsb = Convert.ToDouble(bitDefRow.RowData[bitDefTable.LsbBitIdx]);
            EfuseConfigMainSheet firstConfigSheet = EFuseInputData.EfuseConfigMainSheets
                        .FirstOrDefault(x => x.SheetName.ContainsIgnoreCase(blockKey)) ?? EFuseInputData.EfuseConfigMainSheets.First();

            var configRows = firstConfigSheet.Rows.Where(x => x.Lsb >= lsb && x.Msb <= msb).ToList();
            foreach (EfuseConfigMainRow configRow in configRows)
            {
                var row = new BitDefRow
                {
                    Line = bitDefRow.Line,
                    RowNum = bitDefRow.RowNum,
                    SheetName = bitDefRow.SheetName
                };
                var rowData = new List<string>();
                rowData.AddRange(Enumerable.Repeat("", bitDefTable.MaxColIdx));
                row.RowData = rowData;
                bool isReal = configRow.ValueToFuse.ContainsIgnoreCase("real");
                PopulateConfigRowData(row, bitDefTable, bitDefRow, configRow, isReal);

                bitDefTable.Rows.Add(row);
            }
        }
        private void PopulateConfigRowData(BitDefRow row, BitDefTable bitDefTable, BitDefRow bitDefRow, EfuseConfigMainRow configRow, bool isReal)
        {
            if (bitDefTable.BankEfuseBitDefIdx != -1)
            {
                row.RowData[bitDefTable.BankEfuseBitDefIdx] = configRow.Description;
            }

            if (bitDefTable.MsbBitIdx != -1)
            {
                row.RowData[bitDefTable.MsbBitIdx] = configRow.Msb.ToString();
            }

            if (bitDefTable.LsbBitIdx != -1)
            {
                row.RowData[bitDefTable.LsbBitIdx] = configRow.Lsb.ToString();
            }

            if (bitDefTable.BitWidthIdx != -1)
            {
                row.RowData[bitDefTable.BitWidthIdx] = configRow.Width.ToString();
            }

            if (bitDefTable.Na1Idx != -1)
            {
                row.RowData[bitDefTable.Na1Idx] = "N/A";
            }

            if (bitDefTable.Na2Idx != -1)
            {
                row.RowData[bitDefTable.Na2Idx] = "N/A";
            }

            if (bitDefTable.Na3Idx != -1)
            {
                row.RowData[bitDefTable.Na3Idx] = "N/A";
            }

            if (bitDefTable.Na4Idx != -1)
            {
                row.RowData[bitDefTable.Na4Idx] = "N/A";
            }

            if (bitDefTable.ProgrammingStageIdx != -1)
            {
                if (isReal)
                {
                    row.RowData[bitDefTable.ProgrammingStageIdx] = configRow.FuseBlowLocation;
                }
                else
                {
                    row.RowData[bitDefTable.ProgrammingStageIdx] =
                        bitDefRow.RowData[bitDefTable.ProgrammingStageIdx];
                }
            }
            if (bitDefTable.LowLimitIdx != -1)
            {
                if (isReal)
                {
                    row.RowData[bitDefTable.LowLimitIdx] = "b0";
                }
                else
                {
                    row.RowData[bitDefTable.LowLimitIdx] =
                        bitDefRow.RowData[bitDefTable.LowLimitIdx];
                }
            }
            if (bitDefTable.HighLimitIdx != -1)
            {
                if (isReal)
                {
                    row.RowData[bitDefTable.HighLimitIdx] = "b".PadRight(configRow.Width + 1, '1');
                }
                else
                {
                    row.RowData[bitDefTable.HighLimitIdx] =
                        bitDefRow.RowData[bitDefTable.HighLimitIdx];
                }
            }
            if (bitDefTable.ResolutionIdx != -1)
            {
                row.RowData[bitDefTable.ResolutionIdx] = bitDefRow.RowData[bitDefTable.ResolutionIdx];
            }

            if (bitDefTable.AlgorithmIdx != -1)
            {
                if (isReal)
                {
                    row.RowData[bitDefTable.AlgorithmIdx] = "app";
                }
                else
                {
                    row.RowData[bitDefTable.AlgorithmIdx] = bitDefRow.RowData[bitDefTable.AlgorithmIdx];
                }
            }
            if (bitDefTable.DescriptionIdx != -1)
            {
                row.RowData[bitDefTable.DescriptionIdx] = bitDefRow.RowData[bitDefTable.DescriptionIdx];
            }

            if (bitDefTable.DefaultOrRealIdx != -1)
            {
                row.RowData[bitDefTable.DefaultOrRealIdx] = bitDefRow.RowData[bitDefTable.DefaultOrRealIdx];
            }

            if (bitDefTable.DefaultValueIdx != -1)
            {
                row.RowData[bitDefTable.DefaultValueIdx] = configRow.ValueToFuse;
            }
        }

        internal void CheckDatabaseRevision(BitDefTable bitDefTable, BitDefRow bitDefRow)
        {
            if (EFuseInputData.EfuseDatabaseRevision > 0)
            {
                bool isModifiedByTool = false;
                if (bitDefRow.RowData[bitDefTable.BankEfuseBitDefIdx].Equals("efuse_database_revision",
                    StringComparison.CurrentCultureIgnoreCase))
                {
                    if (bitDefRow.RowData[bitDefTable.DefaultOrRealIdx].Equals("real", StringComparison.CurrentCultureIgnoreCase))
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_MismatchDefaultType_01, bitDefRow.SheetName, bitDefRow.RowNum, 16, []);
                    }
                    string binRevision = ConvertDecimalToBinary(EFuseInputData.EfuseDatabaseRevision);
                    if (bitDefRow.RowData.Count > 15)
                    {
                        if (int.TryParse(bitDefRow.RowData[3], out int binLength))
                        {
                            binRevision = binRevision.PadLeft(binLength, '0').PadLeft(binLength + 1, 'b');
                            if (!bitDefRow.RowData[9].Equals(binRevision,
                                StringComparison.CurrentCultureIgnoreCase) || !bitDefRow.RowData[10].Equals(binRevision,
                                StringComparison.CurrentCultureIgnoreCase) || !bitDefRow.RowData[13].Equals(binRevision,
                                StringComparison.CurrentCultureIgnoreCase) || !bitDefRow.RowData[15].Equals(binRevision,
                                StringComparison.CurrentCultureIgnoreCase)) //low limit, high limit, description, default value
                            {
                                if (LocalSpecs.Options.Device == EnumDevice.RF)
                                {
                                    bitDefRow.RowData[9] = binRevision;
                                    bitDefRow.RowData[10] = binRevision;
                                    bitDefRow.RowData[13] = binRevision;
                                    bitDefRow.RowData[15] = binRevision;
                                    isModifiedByTool = true;
                                }
                                else
                                {
                                    //string message = "The database revision is not same with the efuse database revision sheet content, please check it!";
                                    ErrorReportManager.AddError(EFuseErrorType.E_MismatchRevision_01, bitDefRow.SheetName, bitDefRow.RowNum, 16, []);
                                }
                            }

                            if (isModifiedByTool)
                            {
                                //string message = "The database revision is judge by tool, please check it!";
                                ErrorReportManager.AddError(EFuseErrorType.W_RuleViolationRevision_01, bitDefRow.SheetName, bitDefRow.RowNum, 16, []);
                            }
                        }
                    }
                }
            }
        }

        protected internal List<BitDefTable> AddDummyFuseForBankMon(List<BitDefTable> bitDefTables)
        {
            var finalBitDefTables = new List<BitDefTable>();
            EfuseArraySizeRow monRowInEfuseArraySize = EFuseInputData.EfuseArraySizeSheet.Rows.Find(x => x.FuseArray.EndsWith(BankType.Mon, StringComparison.CurrentCultureIgnoreCase));
            if (monRowInEfuseArraySize == null)
            {
                return bitDefTables;
            }

            string ofBits = monRowInEfuseArraySize.OfBits;
            foreach (BitDefTable bitDefTable in bitDefTables)
            {
                if (bitDefTable.BlockName.Equals(BankType.Mon, StringComparison.CurrentCultureIgnoreCase))
                {
                    BitDefRow lastRow = bitDefTable.Rows.Last();
                    double maxBitFromArraySize = Convert.ToDouble(ofBits) / (bitDefTable.BitMode.Contains("2") ? 2 : 1);
                    double maxBitFromBitDef = Convert.ToDouble(lastRow.RowData[bitDefTable.LsbBitIdx]) > Convert.ToDouble(lastRow.RowData[bitDefTable.MsbBitIdx]) ? Convert.ToDouble(lastRow.RowData[bitDefTable.LsbBitIdx]) : Convert.ToDouble(lastRow.RowData[bitDefTable.MsbBitIdx]);
                    if (maxBitFromArraySize - 1 > maxBitFromBitDef)
                    {
                        var row = new BitDefRow { RowNum = lastRow.RowNum + 1, SheetName = lastRow.SheetName };
                        var rowData = new List<string>();
                        rowData.AddRange(Enumerable.Repeat("", bitDefTable.MaxColIdx));
                        row.RowData = rowData;
                        if (bitDefTable.BankEfuseBitDefIdx != -1)
                        {
                            row.RowData[bitDefTable.BankEfuseBitDefIdx] = "MON_condition_TE";
                        }

                        if (bitDefTable.MsbBitIdx != -1 && bitDefTable.LsbBitIdx != -1)
                        {
                            if (Convert.ToDouble(lastRow.RowData[bitDefTable.LsbBitIdx]) > Convert.ToDouble(lastRow.RowData[bitDefTable.MsbBitIdx]))
                            {
                                row.RowData[bitDefTable.MsbBitIdx] = (Convert.ToDouble(lastRow.RowData[bitDefTable.LsbBitIdx]) + 1).ToString(CultureInfo.InvariantCulture);
                                row.RowData[bitDefTable.LsbBitIdx] = (maxBitFromArraySize - 1).ToString(CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                row.RowData[bitDefTable.LsbBitIdx] = (Convert.ToDouble(lastRow.RowData[bitDefTable.MsbBitIdx]) + 1).ToString(CultureInfo.InvariantCulture);
                                row.RowData[bitDefTable.MsbBitIdx] = (maxBitFromArraySize - 1).ToString(CultureInfo.InvariantCulture);
                            }
                        }
                        if (bitDefTable.BitWidthIdx != -1)
                        {
                            row.RowData[bitDefTable.BitWidthIdx] = (Math.Abs(Convert.ToDouble(row.RowData[bitDefTable.LsbBitIdx]) - Convert.ToDouble(row.RowData[bitDefTable.MsbBitIdx])) + 1).ToString(CultureInfo.InvariantCulture);
                        }

                        if (bitDefTable.Na1Idx != -1)
                        {
                            row.RowData[bitDefTable.Na1Idx] = "N/A";
                        }

                        if (bitDefTable.Na2Idx != -1)
                        {
                            row.RowData[bitDefTable.Na2Idx] = "N/A";
                        }

                        if (bitDefTable.Na3Idx != -1)
                        {
                            row.RowData[bitDefTable.Na3Idx] = "N/A";
                        }

                        if (bitDefTable.Na4Idx != -1)
                        {
                            row.RowData[bitDefTable.Na4Idx] = "N/A";
                        }

                        if (bitDefTable.ProgrammingStageIdx != -1)
                        {
                            row.RowData[bitDefTable.ProgrammingStageIdx] = "SLT";
                        }

                        if (bitDefTable.LowLimitIdx != -1)
                        {
                            row.RowData[bitDefTable.LowLimitIdx] = "0";
                        }

                        if (bitDefTable.HighLimitIdx != -1)
                        {
                            row.RowData[bitDefTable.HighLimitIdx] = "0";
                        }

                        if (bitDefTable.ResolutionIdx != -1)
                        {
                            row.RowData[bitDefTable.ResolutionIdx] = "N/A";
                        }

                        if (bitDefTable.AlgorithmIdx != -1)
                        {
                            row.RowData[bitDefTable.AlgorithmIdx] = "app";
                        }

                        if (bitDefTable.DescriptionIdx != -1)
                        {
                            row.RowData[bitDefTable.DescriptionIdx] = "0";
                        }

                        if (bitDefTable.DefaultOrRealIdx != -1)
                        {
                            row.RowData[bitDefTable.DefaultOrRealIdx] = "Default";
                        }

                        if (bitDefTable.DefaultValueIdx != -1)
                        {
                            row.RowData[bitDefTable.DefaultValueIdx] = "0";
                        }

                        row.Line = string.Join(@"\t", row.RowData);
                        bitDefTable.Rows.Add(row);
                    }
                }
                finalBitDefTables.Add(bitDefTable);
            }
            return finalBitDefTables;
        }

        internal static string ConvertDecimalToBinary(int data)
        {
            string result = "";
            while (data != 0)
            {
                data = Math.DivRem(data, 2, out int rem);
                result = rem + result;
            }
            return result;
        }
    }
}
