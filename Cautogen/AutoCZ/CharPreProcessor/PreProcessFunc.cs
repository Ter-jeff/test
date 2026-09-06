using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

using Cautogen.AutoCZ.CharPreProcessor.ErrorReport;
using Cautogen.AutoCZ.CharPreProcessor.ImFile;
using Cautogen.AutoCZ.CharPreProcessor.InputReader;
using Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader;
using Cautogen.AutoCZ.CharPreProcessor.PreCheck;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager.PatternError;
using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.IgxlProgramLib.IgxlProgramParser;
using Cautogen.common.IgxlProgramMappingLib;
using Cautogen.common.ReaderWriter.Reader;
using Cautogen.common.ReaderWriter.Writer;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib;

using LogLib.Static;
using LogLib.Utility;

namespace Cautogen.AutoCZ.CharPreProcessor
{
    public class PreProcessFunc
    {
        /* perproties */
        public static string VersionStamp = "V9.37";
        public static string FileTimeStamp = "";
        private readonly ParamData _param;
        private static bool _commandLineMode = false;
        private static string _project = "";
        private static string _patListFile = "";
        private string _imfilePath = "";
        private IgxlProgram _igxlProgram;
        private MappingResult _mappingResult;
        private string exportDir;
        public string ImFilePath
        {
            get
            {
                if (string.IsNullOrEmpty(_imfilePath))
                {
                    string charPlanFilename = Path.GetFileNameWithoutExtension(_param.CharPlanFile);
                    _imfilePath = Path.Combine(_param.TarDic,
                        charPlanFilename + "_Result_" + VersionStamp + FileTimeStamp + ".xlsx");
                }
                return _imfilePath;
            }
        }
        private string ErrorReportPath
        {
            get { return ImFilePath.Replace("_Result", "_ErrorReport"); }
        }

        public IgxlProgram IGXLProgram { get { return _igxlProgram; } }

        /* constructor */
        public PreProcessFunc(ParamData param, IProgress<string> progress = null)
        {
            FileTimeStamp = UtilityFunction.GetTimeStamp();
            _param = param;

            // reset previous result
            UtilityMain.Reset();
            UtilityMain.UtilityData.InputParam = param;
            CharPlan.Reset();
            MessageWriter.Progress = progress;
        }
        private bool _LoadProgram()
        {
            //create output folder
            if (!Directory.Exists(_param.TarDic))
            {
                Directory.CreateDirectory(_param.TarDic);
            }

            try
            {
                exportDir = Path.Combine(_param.TarDic, "exportProg");
                MessageWriter.WriteMessage("Export Test program to exportProg", EnumMessageLevel.Info);
                if (!string.IsNullOrEmpty(_param.BaseProgram) && File.Exists(_param.BaseProgram))
                {
                    IgxlManager.ExportWorkBook(_param.BaseProgram, exportDir);
                }

                MessageWriter.WriteMessage("Load Test program", EnumMessageLevel.Info);
                _igxlProgram = new IgxlProgram(_param.JobName);
                if (!string.IsNullOrEmpty(_param.BaseProgram))
                {
                    _igxlProgram.LoadIgxlForPreProcess(exportDir);
                    UtilityMain.UtilityData.InputParam.PinMapFile = Directory.GetFiles(exportDir, "*pinmap*.txt")[0];
                    UtilityMain.UtilityData.InputParam.GlobalSpecsFile = Directory.GetFiles(exportDir, "Global%20Specs.txt")[0];
                }
                else
                {
                    _igxlProgram.LoadIgxlForPreProcess();
                    UtilityMain.UtilityData.InputParam.PinMapFile = TestProgram.IgxlWorkBk.PinMapPair.Key + ".txt";
                    UtilityMain.UtilityData.InputParam.GlobalSpecsFile = TestProgram.IgxlWorkBk.GlbSpecSheetPair.Key + ".txt";
                }
                UtilityMain.UtilityData.FrcList = _igxlProgram.FrcList;
                UtilityMain.UtilityData.AcSpecsSymbols = _igxlProgram.AcSpecsSymbols;

                _mappingResult = new MappingResult(_igxlProgram.InstanceSheets, _igxlProgram.PatSetSheets);

                bool programWithNewTchar = true;
                //if (_igxlProgram.TestLibraryManager
                //    .GetFunctionByName("Functional_T_char")
                //    .Parameters.Split(',')
                //    .Any(x => string.Equals(x, "INIT_PATSET", StringComparison.OrdinalIgnoreCase)))
                //{
                //    programWithNewTchar = true;
                //}

                if (UtilityMain.UtilityData.InputParam.CharPreCheckForNewTChar != programWithNewTchar)
                {
                    if (programWithNewTchar)
                    {
                        Response.Report("Test program uses new Functional_T_char, tool will check plan with new format rule.");
                    }
                    else
                    {
                        Response.Report("Test program uses previous Functional_T_char, tool will check plan the previous format rule.");
                    }
                }
                UtilityMain.UtilityData.InputParam.CharPreCheckForNewTChar = programWithNewTchar;

                //Get main flow sheet for determinate which char plan need to process
                IgxlLib.IgxlSheets.SubFlowSheet mainFlowSheet = _igxlProgram.GetMainFlowSheet();


                //Get the A column CharPlan
                if (mainFlowSheet != null)
                {
                    // Temporarily disable Multiple Char Plan settings
                    //List<IgxlLib.IgxlBase.FlowRow> charPlanFromTPList = mainFlowSheet.Rows.Where(p => !string.IsNullOrEmpty(p.ColumnA)).ToList();
                    List<IgxlLib.IgxlBase.FlowRow> charPlanFromTPList = [];

                    if (Directory.Exists(_param.CharPlanFile))
                    {
                        foreach (IgxlLib.IgxlBase.FlowRow charPlanFromTP in charPlanFromTPList)
                        {
                            if (charPlanFromTP == null)
                            {
                                continue;
                            }
                            string charPlanFromTPPath = Path.Combine(_param.CharPlanFile, charPlanFromTP.ColumnA);
                            _param.CharPlanFileList.Add(charPlanFromTPPath);
                            LogHelper.Info($"Get CharPlan from {mainFlowSheet.Name}");
                            LogHelper.Info($"CharPlan: {charPlanFromTP}");
                            LogHelper.Info($"CharPlan Path: {charPlanFromTPPath}");
                        }
                    }

                    if (_param.CharPlanFileList.Count != 0)
                    {
                        EpplusExtensions.MergeWorkbooks(_param.CharPlanFileList, Path.Combine(_param.TarDic, @"merged_.xlsx"));
                        _param.CharPlanFile = Path.Combine(_param.TarDic, @"merged_.xlsx");
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                Response.Report(e.Message, EnumMessageLevel.Error);
                return false;
            }
        }

        private void GetSelSeamTableFromTP()
        {
            if (string.IsNullOrEmpty(_param.BaseProgram))
            {
                List<string> nonIgxlsheets = TestProgram.NonIgxlSheetsList.SheetList;
                string selSramSheet = nonIgxlsheets.Where(s => Regex.IsMatch(s, "SELSRM.*Mapping.*Table", RegexOptions.IgnoreCase)).FirstOrDefault();
                _param.SelSramMappingTable = selSramSheet;

            }
            else
            {
                if (!string.IsNullOrEmpty(exportDir))
                {
                    //SELSRM_Mapping_Table.txt
                    ZipArchive f = ZipFile.Open(_param.BaseProgram, ZipArchiveMode.Read);
                    string root = Path.GetDirectoryName(_param.BaseProgram);
                    var SelSramTable = f.Entries.ToList().Where(p => Regex.IsMatch(p.Name, "SELSRM.*Mapping.*Table", RegexOptions.IgnoreCase)).ToList();
                    if (SelSramTable.Count > 0)
                    {
                        _param.SelSramMappingTable = Path.Combine(root, SelSramTable[0].FullName);
                        SelSramTable[0].ExtractToFile(_param.SelSramMappingTable, true);
                    }

                }
            }
        }

        /* methods */
        public string Run(bool isCheckOnly = false) // generate intermediate file and Precheck inputs
        {
            try
            {
                if (!_LoadProgram() && !isCheckOnly)
                {
                    throw new Exception("Failed to load test program.");
                }

                _param.IgxlProgram = IGXLProgram;

                GetSelSeamTableFromTP();

                _patListFile = _param.PatListFile;
                if (!File.Exists(_param.CharPlanFile) && _param.CharPlanFileList.Count == 0)
                {
                    throw new Exception(_param.CharPlanFile + "does NOT exist!");
                }

                MessageWriter.SetMessageText();
                _project = _param.ProjectName;
                _commandLineMode = _param.CommandLineMode;
                _InputRead();
                _PreCheck();
                // print the error report
                if (ErrorManager.ErrorListDict.Count > 0)
                {
                    var writer = new ErrorReportCtrl(
                        UtilityMain.UtilityData.InputParam.CharPlanFile,
                        ErrorReportPath,
                        new List<IExcelSheetWriter>
                        {
                            new PatternErrorSheet(),
                            new ErrorReportSheet(),
                            new CorrectItemSheet(),
                        });
                    writer.Write();
                }

                // generate intermediate file
                if (!isCheckOnly)
                {
                    var excelWriter = new ImFileWriter();
                    excelWriter.WriteIntermediateFile(CharPlan.CharPlanSheetDict, ImFilePath);
                }

                if (ErrorManager.ErrorListDict.Count > 0)
                {
                    Response.Report($"Here is some error, please check error report in {ErrorReportPath}");
                }

                else
                {
                    Response.Report($"PreProcessor finished! No errors. ");
                }
            }
            catch (Exception e)
            {
                Response.Report($"\r\nPreProcessor failed!   {e.Message}");

                if (File.Exists(ImFilePath))
                {
                    File.Delete(ImFilePath);
                }

                throw e;
            }
            return ImFilePath;
        }

        public string RunCheckerCmd(string outputPath)
        {
            string pathErrorReport = "";
            if (string.IsNullOrEmpty(outputPath))
            {
                outputPath = ErrorReportPath;
            }

            try
            {
                if (!File.Exists(_param.CharPlanFile))
                {
                    throw new Exception(_param.CharPlanFile + "does NOT exist!");
                }

                if (!string.IsNullOrEmpty(_param.BaseProgram))
                {
                    if (!_LoadProgram())
                    {
                        throw new Exception("Failed to load test program.");
                    }
                }
                _patListFile = _param.PatListFile;

                MessageWriter.SetMessageText();
                _project = _param.ProjectName;
                _InputRead();
                _PreCheck();

                // print the error report
                if (ErrorManager.ErrorListDict.Count > 0)
                {
                    var writer = new ErrorReportCtrl(
                        UtilityMain.UtilityData.InputParam.CharPlanFile,
                        outputPath,
                        new List<IExcelSheetWriter>
                        {
                            new PatternErrorSheet(),
                            new ErrorReportSheet(),
                            new CorrectItemSheet(),
                        });
                    writer.Write();
                    pathErrorReport = writer.ExcelFilePath;
                }
                if (ErrorManager.ErrorListDict.Count > 0)
                {
                    Response.Report("Here is some error, please check error report in " + outputPath);
                }
                else
                {
                    Response.Report("PreProcessor finished! No errors. ");

                }
            }
            catch (Exception e)
            {
                Response.Report($"\r\nPreProcessor failed! {e.Message} ");

                if (File.Exists(ImFilePath))
                {
                    File.Delete(ImFilePath);
                }

                throw e;
            }
            return pathErrorReport;
        }

        private void _InputRead()
        {
            var inputReaderCtrl = new InputReaderCtrl(new List<IReader>
            {
                new PatInfoReader(_param.PatinfoFile),
                new PatternListInputReader(_param.PatListFile),
                new PinmapInputReader(UtilityMain.UtilityData.InputParam.PinMapFile),
                new GlobalSpecsReader(UtilityMain.UtilityData.InputParam.GlobalSpecsFile),
                new AcSpecsReader(_param.AcSpecsFile),
                //20230725 take out for new mapping rule
                //new InputDefReader(_param.DefFile),
                new SelSrmReader(_param.SelSramMappingTable),
                new CharPlan(_mappingResult, _param.CharPlanFile, _param.IsUseRtosCmd, _param.IsMergeHlv,_param.ShmooPowerPinHightoLow),
            });
            inputReaderCtrl.WorkFlow();
        }

        private static void _PreCheck()
        {
            var preCheckCtrl = new PreCheckCtrl(new List<IPreCheck>
            {
                new SelsramPatternChecker(),
                new DigSrcChecker(),
                new IllegalCharChecker(),
                //new MissingPatternChecker(PatternListInputReader.PatternList, _project, _patListFile, _commandLineMode),
                new DuplicatePatternChecker(),
                new MixedSiDmChecker(),
                new HacChecker(),
                new ShmooChecker(),
                new ForceConditionChecker(),
                new EmptySheetChecker(),
                new RetentionChecker(),
                new UslLslChecker(),
                new HarvTableChecker(),
                new Userder1Checker(),
                new VihVilChecker(CharPlan.HardIpSheets),
                new ShmooPinsChecker(),
                new TestNameChecker(),
                new PinSeqChecker(),
                new PowerNetNameChecker(),
                new PowerRunScenarioFormatChecker(),
                //new PatternVersionChecker(PatternListInputReader.PatternList,_project, _commandLineMode),
                new RtosChecker(),
                new AtLeastOnePayloadChecker(),
                new PerformanceModeMatchesDomainChecker(),
                new ManualAcSpecsChecker(),
                new ProgramMappingResultChecker(),
                new PatternCellChecker()
            });
            preCheckCtrl.WorkFlow();
        }
    }
}
