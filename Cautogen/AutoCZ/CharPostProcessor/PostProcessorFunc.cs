using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.Bussiness;
using Cautogen.AutoCZ.CharPostProcessor.Controller;
using Cautogen.AutoCZ.CharPostProcessor.ExportProgram;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.InputReader;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility;
using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;
using Cautogen.AutoCZ.CharPostProcessor.Utility.VbtModuleManager;
using Cautogen.common.IgxlProgramLib.IgxlProgramParser;
using Cautogen.common.ReaderWriter.Reader;

using IgxlLib;
using IgxlLib.IgxlSheets;

using LogLib.Utility;

using TestPlanLib.PatternListCsvFile;

namespace Cautogen.AutoCZ.CharPostProcessor
{
    public class PostProcessorFunc
    {
        private readonly InputParam _param = new InputParam();
        private IProgress<string> _progress = null;

        public PostProcessorFunc(InputParam param, IProgress<string> progress = null)
        {
            _param = param;
            _progress = progress;
            LocalSpecs.Reset(param, progress);  //reset static containers
        }


        public string Execute(out string enableWord, bool isAutoCz = false)
        {
            enableWord = "";
            try
            {
                //read post settings
                var readerCtrl = new ReaderCtrl(new List<IReader>
                {
                    new PostSettings(_param.PostSettings)
                });

                readerCtrl.WorkFlow();

                string exportProg = Path.Combine(_param.OutputFolder, "exportProg");
                if (!LocalSpecs.IsDirectory && !string.IsNullOrEmpty(_param.ProgWorkBookPath))
                {
                    IgxlManager.ExportWorkBook(_param.ProgWorkBookPath, exportProg);
                }

                _param.TimeSetVersionDic = GetTimeSetVersionDic(_param.TimeSetFolder, exportProg);

                ExportMain.CopyExportFile(_param.OutputFolder);
                if (_param.GenCSharp)
                {
                    BasMain.Parse(_param.CsLibraryPath, true);
                }
                else
                {
                    BasMain.Parse(exportProg);
                }

                if (BasMain.VbtFunctionLib
                    .GetFunctionByName("Functional_T_char")
                    .Parameters.Split(',')
                    .Any(x => string.Equals(x, "INIT_PATSET", StringComparison.OrdinalIgnoreCase)))
                {
                    _param.UseNewTChar = true;
                }

                if (_param.GenCSharp)
                {
                    _param.UseNewTChar = true;
                }

                ReaderMain.Run(_param);

                //load prod program to LocalSpecs using ig-data 
                LoadProdProgram(_param, exportProg);

                AddSheetToIgxlWorkBk(_param.IgxlProgram);

                ProdProg.Parse(_param.JobName, _param.ChannelMapName);

                //generate output sheets
                GeneratorMain.Run(_param, out enableWord);

                //early termination
                if (_param.GenTxtOnly)
                {
                    _progress.Report("Txt Sheet generate done...Please check");
                    return "";
                }

                //add Flow_Char to Main_Flow
                LogHelper.Info(@"Generating Flow_Char ...");
                GenMainFlow.WorkFlow();

                //add VBT_LIB_PV
                if (!string.IsNullOrEmpty(_param.EnableWords))
                {
                    LogHelper.Info(@"Generating VBT_LIB_PV.Bas ...");
                    new GenBas().WorkFlow(_param.EnableWords, _param.ProgWorkBookPath, enableWord);
                }

                //update JobList sheet
                LogHelper.Info(@"Updating JobList ...");
                UpdateJobListSheet.WorkFlow(LocalSpecs.TestProgram);

                LogHelper.Info(@"Updating TimeSet Sheet");
                UpdateTimeSetSheet.WorkFlow();

                LogHelper.Info(@"Copying others in original test program ...");
                LocalSpecs.TestProgram.PrintOthers(Path.Combine(_param.OutputFolder, ConstData.TrunkFolder), exportProg,
                    Path.Combine(_param.OutputFolder, ConstData.Others));

                //generate Igxl porgram
                //Preserve Extra Sheet
                RemoveDuplicateInstance.WorkFlow(_param.OutputFolder);
                LogHelper.Info(@"Generating IgxlProgram ...");

                var genSheetsDict = new Dictionary<string, IIgxlSheet>();
                foreach (IIgxlSheet sheet in LocalSpecs.GenSheets)
                {
                    AddSheetToIgLink(sheet);
                }
                //Copy Solution folder into output result

                if (_param.GenCSharp)
                {
                    string libraryPath = Path.Combine(_param.OutputFolder, "IGLink", "central_library_cs", "bin");
                    if (!Directory.Exists(libraryPath))
                    {
                        Directory.CreateDirectory(libraryPath);
                    }
                    FolderOperation.CopyFilesRecursively(_param.CsLibraryPath, libraryPath);
                }

                return "";
            }
            catch (Exception e)
            {
                if (LocalSpecs.MessageWriter != null)
                {
                    LocalSpecs.MessageWriter.WriteLine("Generating IG-XL program failed " + e.Message);
                }

                throw;
            }
            finally
            {
                if (LocalSpecs.MessageWriter != null)
                {
                    LocalSpecs.MessageWriter.Close();
                }
            }
        }

        private static Dictionary<string, TimeSetItem> GetTimeSetVersionDic(string timeSetFolder, string exportProg)
        {
            var timeSetFiles = new List<string>();
            if (!string.IsNullOrEmpty(timeSetFolder) && Directory.Exists(timeSetFolder))
            {
                timeSetFiles = Directory.GetFiles(timeSetFolder, "TIMESET*.txt", SearchOption.TopDirectoryOnly).ToList();
            }

            if (!string.IsNullOrEmpty(exportProg) && Directory.Exists(exportProg))
            {
                timeSetFiles.AddRange(Directory.GetFiles(exportProg, "TIMESET*.txt", SearchOption.TopDirectoryOnly).ToList());
            }

            var dic = new Dictionary<string, TimeSetItem>(StringComparer.CurrentCultureIgnoreCase);
            foreach (string file in timeSetFiles)
            {
                string setName = Path.GetFileName(file);
                if (Regex.IsMatch(setName, @"_\d+.TXT$", RegexOptions.IgnoreCase))
                {
                    var timeItem = new TimeSetItem();
                    string timeset = Regex.Match(setName, @"(?<str>.*)_\d+.TXT$", RegexOptions.IgnoreCase).Groups["str"].ToString().ToUpper();
                    int paraVer = Convert.ToInt32(Regex.Match(setName, @".*_(?<ver>\d+).TXT$", RegexOptions.IgnoreCase).Groups["ver"].ToString());
                    timeItem.Version = paraVer;

                    //Timing Mode:	Single
                    using (var readerTime = new StreamReader(File.OpenRead(file)))
                    {
                        while (!readerTime.EndOfStream)
                        {
                            string line = readerTime.ReadLine();
                            if (Regex.IsMatch(line, @"Timing\s+Mode:.*Master\s+Timeset", RegexOptions.IgnoreCase))
                            {
                                string timeMode =
                                    Regex.Match(line, @"Timing\s+Mode:\t+(?<str>\w+)\t+").Groups["str"].ToString();
                                if (string.IsNullOrEmpty(timeItem.TimeMod))
                                {
                                    timeItem.TimeMod = timeMode;
                                }

                                if (timeMode != timeItem.TimeMod)
                                {
                                    timeItem.TimeMod += "/" + timeMode;
                                }

                                break;
                            }
                        }
                    }

                    if (!dic.ContainsKey(timeset))
                    {
                        dic.Add(timeset, timeItem);
                    }
                    else
                    {
                        if (dic[timeset].Version < timeItem.Version)
                        {
                            dic[timeset].Version = timeItem.Version;
                        }

                        dic[timeset].TimeMod = timeItem.TimeMod;
                    }
                }
            }
            return dic;
        }

        private static void LoadProdProgram(InputParam param, string exportProg)
        {
            var watch = new Stopwatch();
            watch.Start();
            LocalSpecs.MessageWriter.WriteLine("Loading Test Program " + param.ProgWorkBookPath);

            if (Directory.Exists(exportProg))
            {
                try
                {
                    if (LocalSpecs.CharPlanSheets == null)
                    {
                        LocalSpecs.MessageWriter.WriteLine("There are no char plan !!!");
                    }

                    var usedPayloads = LocalSpecs.CharPlanSheets.Where(p => p.IsHardIp)
                        .SelectMany(p => p.GetReferencePayloads()).Distinct().ToList();
                    if (param.IgxlProgram == null)
                    {
                        LocalSpecs.TestProgram.LoadIgxlProgramAsync(exportProg, usedPayloads, param.JobName);

                    }
                }
                catch (Exception e)
                {
                    LocalSpecs.MessageWriter.WriteLine(e.ToString());
                    LocalSpecs.MessageWriter.WriteLine(exportProg + " can not be load program !");
                    throw new IOException("Fail to load prod program");
                }
            }
            else if (LocalSpecs.TestProgram != null)
            {
                LocalSpecs.MessageWriter.WriteLine("Loading Test Program From T-Autogen");

            }
            LocalSpecs.MessageWriter.WriteLine("Loading Test Program: " + watch.ElapsedMilliseconds + "ms");
        }

        private static void AddSheetToIgLink(IIgxlSheet sheet)
        {
            Type sheetType = sheet.GetType();
            if (sheetType == typeof(BinTableSheet))
            {
                string dir = Path.Combine(LocalSpecs.OutputFolder, ConstData.Binfolder);
                string sheetPath = Path.Combine(dir, sheet.Name + ".txt");

                if (!File.Exists(sheetPath))
                {
                    return;
                }

                Automation.Static.TestProgram.IgxlWorkBk.AddBinTblSheet(dir, (BinTableSheet)sheet);
            }
            else
            {
                string dir = Path.Combine(LocalSpecs.OutputFolder, ConstData.CzFolder);
                string sheetPath = Path.Combine(dir, sheet.Name + ".txt");

                if (!File.Exists(sheetPath))
                {
                    return;
                }

                if (sheetType == typeof(SubFlowSheet) && !sheet.Name.StartsWith("Main_Flow"))
                {
                    Automation.Static.TestProgram.IgxlWorkBk.AddSubFlowSheet(dir, (SubFlowSheet)sheet);
                }
                else if (sheetType == typeof(InstanceSheet))
                {
                    Automation.Static.TestProgram.IgxlWorkBk.AddInsSheet(dir, (InstanceSheet)sheet);
                }
                else if (sheetType == typeof(CharSheet))
                {
                    Automation.Static.TestProgram.IgxlWorkBk.AddCharSheet(dir, (CharSheet)sheet);
                }
                else if (sheetType == typeof(PatSetSheet))
                {
                    Automation.Static.TestProgram.IgxlWorkBk.AddPatSetSheet(dir, (PatSetSheet)sheet);
                }
            }
        }

        private static void AddSheetToIgxlWorkBk(IgxlProgram igxlProgram)
        {
            if (LocalSpecs.FileStructure.Count == 0)
            {
                return;
            }

            AddBasicSheetsToIgxlWorkBk(igxlProgram);

            AddSpecSheetsToIgxlWorkBk(igxlProgram);

            AddPinAndChannelSheetsToIgxlWorkBk(igxlProgram);

            AddPatternSheetsToIgxlWorkBk(igxlProgram);

            AddFlowAndInstanceSheetsToIgxlWorkBk(igxlProgram);
        }

        private static void AddBasicSheetsToIgxlWorkBk(IgxlProgram igxlProgram)
        {
            if (igxlProgram.JoblistSheet != null)
            {
                Automation.Static.TestProgram.IgxlWorkBk.AddJobListSheet(Automation.Static.FolderStructure.DirJob, igxlProgram.JoblistSheet);
            }

            if (igxlProgram.ReferenceSheets != null)
            {
                foreach (ReferenceSheet referenceSheet in igxlProgram.ReferenceSheets)
                {
                    Automation.Static.TestProgram.IgxlWorkBk.AddReferenceSheet(Automation.Static.FolderStructure.DirReference, referenceSheet);
                }
            }

            if (igxlProgram.CharSheets != null)
            {
                foreach (CharSheet charSheet in igxlProgram.CharSheets)
                {
                    Automation.Static.TestProgram.IgxlWorkBk.AddCharSheet(Automation.Static.FolderStructure.DirDevChar, charSheet);
                }
            }

            if (igxlProgram.BintableSheets != null)
            {
                foreach (BinTableSheet binTableSheet in igxlProgram.BintableSheets)
                {
                    LocalSpecs.FileStructure.TryGetValue(binTableSheet.Name + ".txt", out string dst);
                    string dstFolder = Path.GetDirectoryName(dst);

                    Automation.Static.TestProgram.IgxlWorkBk.AddBinTblSheet(dstFolder, binTableSheet);

                    // extra
                    if (binTableSheet.Name == "Bin_Table_HardIP" || binTableSheet.Name == "Bin_Table_Rtos")
                    {
                        Automation.Static.TestProgram.IgxlWorkBk.AddBinTblSheet(Automation.Static.FolderStructure.DirHardIp, binTableSheet);
                    }
                }
            }

            if (igxlProgram.PortMapSheets != null)
            {
                foreach(PortMapSheet portMapSheet in igxlProgram.PortMapSheets)
                {
                    Automation.Static.TestProgram.IgxlWorkBk.AddPortMapSheet(Automation.Static.FolderStructure.DirPorts, portMapSheet);
                }
            }
        }

        private static void AddSpecSheetsToIgxlWorkBk(IgxlProgram igxlProgram)
        {
            if (igxlProgram.GlbSpecSheet != null)
            {
                Automation.Static.TestProgram.IgxlWorkBk.GlbSpecSheetPair = new KeyValuePair<string, GlobalSpecSheet>(Automation.Static.FolderStructure.DirGlbSpec, igxlProgram.GlbSpecSheet);
            }

            if (igxlProgram.DcSpecSheets != null)
            {
                foreach (DcSpecSheet dcSpecSheet in igxlProgram.DcSpecSheets)
                {
                    Automation.Static.TestProgram.IgxlWorkBk.AddDcSpecSheet(Automation.Static.FolderStructure.DirDcSpec, dcSpecSheet);
                }
            }

            if (igxlProgram.LevelSheets != null)
            {
                foreach (LevelSheet levelSheet in igxlProgram.LevelSheets)
                {
                    LocalSpecs.FileStructure.TryGetValue(levelSheet.Name + ".txt", out string dst);
                    string dstFolder = Path.GetDirectoryName(dst);

                    Automation.Static.TestProgram.IgxlWorkBk.AddLevelSheet(dstFolder, levelSheet);
                }
            }

            if (igxlProgram.AcSpecSheets != null)
            {
                foreach (AcSpecSheet acSpecSheet in igxlProgram.AcSpecSheets)
                {
                    Automation.Static.TestProgram.IgxlWorkBk.AddAcSpecSheet(Automation.Static.FolderStructure.DirAcSpec, acSpecSheet);
                }
            }

            if (igxlProgram.TimeSetSheets != null)
            {
                foreach (TimeSetBasicSheet timeSetBasicSheet in igxlProgram.TimeSetSheets)
                {
                    LocalSpecs.FileStructure.TryGetValue(timeSetBasicSheet.Name + ".txt", out string dst);
                    string dstFolder = Path.GetDirectoryName(dst);

                    Automation.Static.TestProgram.IgxlWorkBk.AddTimeSetSheet(dstFolder, timeSetBasicSheet);
                }
            }
        }

        private static void AddPinAndChannelSheetsToIgxlWorkBk(IgxlProgram igxlProgram)
        {
            if (igxlProgram.ChannelMapSheets != null)
            {
                foreach (ChannelMapSheet channelMapSheet in igxlProgram.ChannelMapSheets)
                {
                    Automation.Static.TestProgram.IgxlWorkBk.AddChannelMapSheet(Automation.Static.FolderStructure.DirChannelMap, channelMapSheet);
                }
            }

            if (igxlProgram.PinMaps != null)
            {
                foreach (PinMapSheet pinMapSheet in igxlProgram.PinMaps)
                {
                    Automation.Static.TestProgram.IgxlWorkBk.PinMapPair = new KeyValuePair<string, PinMapSheet>(Automation.Static.FolderStructure.DirPinMap, pinMapSheet);
                }
            }
        }

        private static void AddPatternSheetsToIgxlWorkBk(IgxlProgram igxlProgram)
        {
            if (igxlProgram.PatSetSheets != null)
            {
                foreach (PatSetSheet patSetAll in igxlProgram.PatSetSheets)
                {
                    LocalSpecs.FileStructure.TryGetValue(patSetAll.Name + ".txt", out string dst);
                    string dstFolder = Path.GetDirectoryName(dst);

                    Automation.Static.TestProgram.IgxlWorkBk.AddPatSetSheet(dstFolder, patSetAll);
                }
            }

            PatSetSubSheet patSetSub = igxlProgram.PatSetSubSheets?.FirstOrDefault(p => p.Name == "Pattern_Subroutine");
            if (patSetSub != null)
            {
                Automation.Static.TestProgram.IgxlWorkBk.AddPatSetSubSheet(Automation.Static.FolderStructure.DirPatSetsAll, patSetSub);
            }
        }

        private static void AddFlowAndInstanceSheetsToIgxlWorkBk(IgxlProgram igxlProgram)
        {
            if (igxlProgram.MainFlowSheets != null)
            {
                foreach (MainFlow mainFlow in igxlProgram.MainFlowSheets)
                {
                    Automation.Static.TestProgram.IgxlWorkBk.AddMainFlowSheet(Automation.Static.FolderStructure.DirMain, mainFlow);
                }
            }

            if (igxlProgram.FlowSheets != null)
            {
                foreach (SubFlowSheet flowSheet in igxlProgram.FlowSheets)
                {
                    if (flowSheet.Name.Contains("Main_Flow"))
                    {
                        continue;
                    }

                    LocalSpecs.FileStructure.TryGetValue(flowSheet.Name + ".txt", out string dst);
                    string dstFolder = Path.GetDirectoryName(dst);

                    Automation.Static.TestProgram.IgxlWorkBk.AddSubFlowSheet(dstFolder, flowSheet);
                }
            }

            if (igxlProgram.InstanceSheets != null)
            {
                foreach (InstanceSheet instanceSheet in igxlProgram.InstanceSheets)
                {
                    LocalSpecs.FileStructure.TryGetValue(instanceSheet.Name + ".txt", out string dst);
                    string dstFolder = Path.GetDirectoryName(dst);

                    Automation.Static.TestProgram.IgxlWorkBk.AddInsSheet(dstFolder, instanceSheet);
                }
            }
        }

        private static bool _ProcessFlowSheet(SubFlowSheet sheet, List<string> Payloads)
        {
            var testItems = sheet.Rows.Where(p => p.Opcode.Equals("test", StringComparison.OrdinalIgnoreCase) &&
                                                      !string.IsNullOrEmpty(p.Parameter)).Select(s => s.Parameter).ToList();
            foreach (string item in testItems)
            {
                if (Payloads.Exists(p => Regex.IsMatch(item, p, RegexOptions.IgnoreCase)))
                {
                    return true;
                }
            }

            return false;
        }

        private string GenerateIgxlProgram(string outputFolder, string jobName)
        {
            GeneralFunc.WriteMessage("Generating IG-XL program... ");

            //copy static files
            //CopyXmlAndSimulatedConfig(outputFolder);

            //get job whose name == jobName in the prod t/p
            string mainJob = LocalSpecs.TestProgram.GetJobList().FirstOrDefault(
                a => a.Equals(jobName, StringComparison.OrdinalIgnoreCase));
            if (mainJob == null)
            {
                return "";
            }

            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(LocalSpecs.InputParam.ProgWorkBookPath);
            if (fileNameWithoutExtension == null)
            {
                return "";
            }

            string projectName = fileNameWithoutExtension.Split('_')[0];
            projectName = projectName != "" ? projectName : "CharProgram";
            string inFolder = Path.Combine(outputFolder, ConstData.TrunkFolder);
            var genIgxl = new GenIgxlProg();
            string result = genIgxl.GenIgxlProgram(projectName, mainJob, inFolder, outputFolder);

            if (_param.GenCSharp)
            {
                string solutionFolderName = Path.GetFileName(FolderOperation.GetParentDirectory(_param.CsLibraryPath, 2));
                string solutionOutputPath = Path.Combine(_param.OutputFolder, "IGLink", solutionFolderName, "bin");
                string solutionStandardPath = Path.Combine(_param.OutputFolder, "IGLink", "central_library_cs", "bin");
                if (!Directory.Exists(solutionOutputPath))
                {
                    Directory.CreateDirectory(solutionOutputPath);
                }

                if (!Directory.Exists(solutionStandardPath))
                {
                    Directory.CreateDirectory(solutionStandardPath);
                }

                FolderOperation.CopyFilesRecursively(_param.CsLibraryPath, solutionOutputPath);
                FolderOperation.CopyFilesRecursively(_param.CsLibraryPath, solutionStandardPath);
            }

            //print log
            if (result != "")
            {
                LocalSpecs.MessageWriter.WriteLine("Generating IG-XL program finished, check result in " + result);
            }

            return result;
        }

        private static void CopyXmlAndSimulatedConfig(string outputFolder)
        {
            //copy xml_Files
            string xmlFolder = Path.Combine(outputFolder, ConstData.XmlFolder);
            if (!Directory.Exists(xmlFolder))
            {
                Directory.CreateDirectory(xmlFolder);
            }

            foreach (string xmlFile in Directory.GetFiles("xml_Files"))
            {
                string fileName = Path.GetFileName(xmlFile);
                if (fileName == null)
                {
                    continue;
                }

                string newFile = Path.Combine(xmlFolder, fileName);
                File.Copy(Path.Combine(Directory.GetCurrentDirectory(), xmlFile), newFile, true);
            }

            //copy simulation config file
            string simuFile = Path.Combine(outputFolder, "IGLink", "SimulatedConfig.txt");
            File.Copy(Path.Combine(Directory.GetCurrentDirectory(), "simulationConfig", "SimulatedConfig.txt"), simuFile, true);
        }
    }
}
