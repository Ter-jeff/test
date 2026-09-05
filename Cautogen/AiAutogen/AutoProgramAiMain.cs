using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.AiAutogen.AutoProgramAi.Write;

using DebugPlanReaderLib.DebugPlan;

using IgxlLib;
using IgxlLib.IgxlSheets;

using LogLib.Utility;

namespace Cautogen.AiAutogen
{
    public static class AutoProgramAiMain
    {
        public static string enableWord = "";

        public static string Main(string jobName, string testProgram, string patternFolder, string outPutFolder, string outPutProgName,
            string enableWords, string doPayloadAfterSRMDSSC, string allSuspendDatalogFalse, DebugPlanMain debugTestPlan, bool isCSharp = false, string csharpLibraryFolder = null)
        {
            try
            {
                string tempFolder = Path.Combine(Path.GetDirectoryName(testProgram) != null ? Path.GetDirectoryName(testProgram) : string.Empty, "Temp");
                if (!Directory.Exists(tempFolder))
                {
                    Directory.CreateDirectory(tempFolder);
                }

                string outputIgxl = Path.Combine(outPutFolder, outPutProgName + ".igxl");
                if (File.Exists(outputIgxl))
                {
                    File.Delete(outputIgxl);
                }

                var igxlData = new IgxlDataReader(testProgram, jobName, isCSharp, csharpLibraryFolder);
                var useNewCharLib = false;

                if (igxlData.VbtFunctionLib
                    .GetFunctionByName("Functional_T_char")
                    .Parameters.Split(',')
                    .Any(x => string.Equals(x, "INIT_PATSET", StringComparison.OrdinalIgnoreCase)))
                {
                    useNewCharLib = true;
                    LogHelper.Info(@"New Functional_T_char ...");

                }

                useNewCharLib = isCSharp;

                var flowTMPS = igxlData.FlowSheets.FirstOrDefault(x => x.Name.StartsWith("Flow_TMPS", StringComparison.OrdinalIgnoreCase));
                var genTMPS = flowTMPS != null;
                var igxlSheets = new List<IIgxlSheet>();

                LogHelper.Info(@"Updating TimeSet ...");
                var updateTimeSet = new UpdateTimeSet().Work(debugTestPlan, igxlData.TimeSetBasicSheets,
                    igxlData.CurrentPortMapSheet, patternFolder);
                igxlSheets.AddRange(updateTimeSet);

                LogHelper.Info(@"Updating AC spec ...");
                var updateAcSpecs = new UpdateAcSpecs().Work(igxlData.CurrentAcSpecSheet, updateTimeSet);
                igxlSheets.Add(updateAcSpecs);

                LogHelper.Info(@"Generating PatSets_All_CZ ...");
                var patSetsAllCz = debugTestPlan.GenPatSetAllSheet(patternFolder, _GetTesterType(testProgram));
                igxlSheets.Add(patSetsAllCz);

                LogHelper.Info(@"Updating PatSets_All ...");
                var patSetsAll = igxlData.PatSetsAll.Remove(patSetsAllCz.Rows.Select(x => x.PatSetName).ToList());
                igxlSheets.Add(patSetsAll);

                LogHelper.Info(@"Updating Pattern_Subroutine ...");
                var patSetSubRows = debugTestPlan.GenPatSetSubRows(patternFolder);
                if (patSetSubRows.Any())
                {
                    igxlData.PatSetSubSheet.AddRows(patSetSubRows);
                    igxlSheets.Add(igxlData.PatSetSubSheet);
                }

                LogHelper.Info(@"Generating PatSets_CZ ...");
                var patSetCz = debugTestPlan.GenPatSetSheet();
                igxlSheets.Add(patSetCz);

                if (!isCSharp)
                {
                    LogHelper.Info(@"Generating VBT_LIB_PV.bas ...");
                    var basFiles = debugTestPlan.GenBas(enableWords, testProgram);
                    igxlSheets.AddRange(basFiles);
                }


                LogHelper.Info(@"Generating Inst_CZ ...");
                var instanceSheet = debugTestPlan.GenInstSheet(igxlData.VbtFunctionLib, updateAcSpecs, doPayloadAfterSRMDSSC, igxlData.LevelSheets, patSetCz, useNewCharLib);
                igxlSheets.Add(instanceSheet);

                LogHelper.Info(@"Generating Flow_CZ ...");
                var updateTMPS = new UpdateTmpsFlow().Work(flowTMPS, debugTestPlan, jobName);
                igxlSheets.AddRange((IEnumerable<IIgxlSheet>)updateTMPS);
                var flowSheet = debugTestPlan.GenFlowSheet(genTMPS);
                igxlSheets.Add(flowSheet);

                LogHelper.Info(@"Generating DFC_List ...");
                var dfcList = debugTestPlan.GenDfcList();
                igxlSheets.Add(dfcList);

                LogHelper.Info(@"Generating Char_CZ ...");
                var charSheet = debugTestPlan.GenCharSheet(igxlData.CurrentPinMapSheet, igxlData.CurrentAcSpecSheet, igxlData.CurrentPortMapSheet, allSuspendDatalogFalse);
                igxlSheets.Add(charSheet);

                LogHelper.Info(@"Updating BinTable ...");
                var updateBinTable = new UpdateBinTable().Work(igxlData.BinTableSheets);
                igxlSheets.Add(updateBinTable);

                LogHelper.Info(@"Updating TestInst_Common ...");
                var updateTestInstComm = new UpdateTestInstComm(isCSharp).Work(igxlData.InstanceSheets, igxlData.VbtFunctionLib);
                if (updateTestInstComm != null)
                    igxlSheets.Add(updateTestInstComm);

                LogHelper.Info(@"Updating JobList ...");
                var updateJobList = new UpdateJobList().Work(igxlData.JobListSheet, patSetsAllCz,
                    patSetCz,
                    instanceSheet, charSheet);
                igxlSheets.Add(updateJobList);

                LogHelper.Info(@"Updating GlobalSpecSheet ...");
                var pins = debugTestPlan.AiTestPlanSheets.SelectMany(x => x.Rows)
                    .SelectMany(x => x.Pins).Select(x => x.ShmooName).Distinct().ToList();
                var updateGlobalSpecSheet = new UpdateGlobalSpecSheet().Work(igxlData.GlobalSpecSheet, pins);
                igxlSheets.Add(updateGlobalSpecSheet);

                LogHelper.Info(@"Updating Reference Sheet ...");
                var updateRReferenceSheet = new UpdateReferenceSheet().Work(igxlData.ReferenceSheets.FirstOrDefault(p => p.Rows.Count() > 0));
                igxlSheets.Add(updateRReferenceSheet);

                LogHelper.Info(@"Inserting Char_CZ into Main_Flow ...");
                var mainFlow = igxlData.JobListSheet.Rows
                    .Find(x => x.JobName.Equals(jobName, StringComparison.CurrentCultureIgnoreCase)).FlowTable;
                var mainFlowSheet = igxlData.FlowSheets.Find(x =>
                    x.Name.Equals(mainFlow, StringComparison.CurrentCultureIgnoreCase));
                var updateMainFlow = new UpdateMainFlow().Work(mainFlowSheet, flowSheet.Name);
                igxlSheets.Add(updateMainFlow);

                LogHelper.Info($"Generating test program {outputIgxl} ...");
                File.Copy(testProgram, outputIgxl, true);
                IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets, tempFolder);

                return outputIgxl;
            }
            catch (Exception e)
            {
                LogHelper.Error(e.StackTrace);
                throw e;
            }
        }
        private static string _GetTesterType(string testProgram)
        {
            var dir = Path.GetDirectoryName(testProgram) != null ? Path.GetDirectoryName(testProgram) : string.Empty;
            if (string.IsNullOrEmpty(dir))
                return ".PAT";

            var result = "";
            var configTxt = new List<string>();
            var testerType = "";
            if (File.Exists(Path.Combine(dir, "SimulatedConfig.txt")))
            {
                configTxt = File.ReadLines(Path.Combine(dir, "SimulatedConfig.txt")).ToList();
            }
            else
            {
                if (Directory.Exists(@"C:\Program Files (x86)\Teradyne\IG-XL\"))
                {
                    var currentSearch = Directory.GetFiles(@"C:\Program Files (x86)\Teradyne\IG-XL\", "CurrentConfig.txt", SearchOption.AllDirectories);
                    if (currentSearch.Any())
                    {
                        configTxt = File.ReadLines(currentSearch.First()).ToList();
                    }
                }
            }

            if (configTxt.Any())
            {
                var typeIdx = configTxt.FindIndex(x => x.IndexOf("<Configuration>", StringComparison.OrdinalIgnoreCase) != -1);
                if (configTxt.Count > typeIdx + 1)
                {
                    testerType = configTxt[typeIdx + 1].Trim();
                }
            }

            if (!string.IsNullOrEmpty(testerType) && string.Equals(testerType, "UltraFLEXplus", StringComparison.OrdinalIgnoreCase))
            {
                result = ".PATX";
            }
            else
            {
                result = System.Environment.MachineName.IndexOf("FXP") != -1 ? ".PATX" : ".PAT";
            }
            return result;
        }
    }
}
