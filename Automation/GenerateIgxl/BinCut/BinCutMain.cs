using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Business;
using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.BinCut;

namespace Automation.GenerateIgxl.BinCut
{
    public class BinCutMain : WorkFlowBase<ParaData>
    {
        private BinCutInputData _binCutInputData;

        public override bool PreCheckFlow(ParaData paraData)
        {
            try
            {
                Response.Report("Reading BinCut Files ...", percentage: 30);
                var binCutInputManager = new BinCutInputManager(EpWorkbook.TestPlanWorkbook);
                _binCutInputData = binCutInputManager.Read();
                return true;
            }
            catch (Exception e)
            {
                Response.Report("BinCut has errors : " + e.StackTrace, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
                return false;
            }
        }

        public override void WorkFlow(ParaData paraData)
        {
            try
            {
                Response.Report("Generating BinCut Files ...", percentage: 50);
                var workFlowManger = new BinCutWorkFlowManager(_binCutInputData);
                workFlowManger.Main();

                Response.Report("BinCut Completed!", percentage: 100);
            }
            catch (Exception e)
            {
                string message = "BinCut AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }

        public void GenerateBincutRelatedFiles()
        {
            bool isCs = !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder);

            Dictionary<string, string> modes = PerformanceModeSingleton.Instance().GetAllPerformanceModeDic();
            if (EpWorkbook.BinCutWorkbook != null)
            {
                Dictionary<string, string> result = BinCutService.NonIgxlSheetProcess(modes, isCs, EpWorkbook.BinCutWorkbook, FolderStructure.DirBinCut);
                foreach (KeyValuePair<string, string> item in result)
                {
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirBinCut, Path.GetFileNameWithoutExtension(item.Key));
                }
            }
            foreach (string filePath in LocalSpecs.BinCutShadowFileNames)
            {
                string shadowStage = Path.GetFileNameWithoutExtension(filePath).Split('_').Last();
                var package = new ExcelPackage(new FileInfo(filePath));
                Dictionary<string, string> result = BinCutService.NonIgxlSheetProcess(modes, isCs, package.Workbook, FolderStructure.DirBinCut, false, shadowStage);
                foreach (KeyValuePair<string, string> item in result)
                {
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirBinCut, Path.GetFileNameWithoutExtension(item.Key));
                }
            }

            if (EpWorkbook.BinCutPostWorkbook != null)
            {
                Dictionary<string, string> resultOutside = BinCutService.NonIgxlSheetProcess(modes, isCs, EpWorkbook.BinCutPostWorkbook, FolderStructure.DirBinCut);
                foreach (KeyValuePair<string, string> item in resultOutside)
                {
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirBinCut, Path.GetFileNameWithoutExtension(item.Key));
                }
            }

            if (EpWorkbook.BinCutModeSeqWorkbook != null)
            {
                Dictionary<string, string> resultOutside = BinCutService.NonIgxlSheetProcess(modes, isCs, EpWorkbook.BinCutModeSeqWorkbook, FolderStructure.DirBinCut);
                foreach (KeyValuePair<string, string> item in resultOutside)
                {
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirBinCut, Path.GetFileNameWithoutExtension(item.Key));
                }
            }

            if (EpWorkbook.PowerBinningWorkbooks != null && EpWorkbook.PowerBinningWorkbooks.Any())
            {
                Dictionary<string, string> filterSourceFile = new Dictionary<string, string>();
                foreach (ExcelWorkbook workbook in EpWorkbook.PowerBinningWorkbooks)
                {
                    if (workbook == null)
                    {
                        continue;
                    }

                    Dictionary<string, string> resultOutside = BinCutService.NonIgxlSheetProcess(modes, isCs, workbook, FolderStructure.DirBinCut);
                    string suffix = string.Empty;
                    foreach (KeyValuePair<string, string> item in resultOutside)
                    {
                        string fileName = Path.GetFileName(item.Key);
                        if (!filterSourceFile.ContainsKey(fileName))
                        {
                            filterSourceFile.Add(fileName, item.Key);
                        }
                        else
                        {
                            ErrorReportManager.AddError(BinCutErrorType.W_DuplicateProgramSheet_01, "", 1, 0, $"Duplicate sheet from {item.Value} in PowerBinning Workbook!", new string[] { item.Value });
                            continue;
                        }
                        TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirBinCut, Path.GetFileNameWithoutExtension(fileName));
                    }
                }
            }

        }
    }
}
