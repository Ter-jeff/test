using System.ComponentModel;
using System.Text.RegularExpressions;

using AutogenCommandLine;
using AutogenCommandLine.CommandLineOptions;

using Automation.Static;

using CommonLib.ErrorReport.Base;

using EfuseCheckCmdLib.EFuse;

using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using LogLib.Utility;

using MyCommandLineLib;

using NLog;

using PatternCompile;

using ProjectConfigLib.ProjectConfig;

using RFPatternTool;

using TestPlanLib.Efuse;

namespace RF_PatternTool
{
    internal static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        private static void Main(string[] args)
        {
            if (args is null)
            {
                throw new ArgumentNullException(nameof(args));
            }

            CommandLineManager.ParseArgList(args, out ICommandLineOptions options, out EntryOptions _);
            if (options != null)
            {
                Run((BenchLogOptions)options);
            }
        }

        static private PinMapSheet _pinmap = null;
        static private LoaderEfuseBitDef _efuseBitDef = null;
        static private HashSet<string> _addrFor64 = null;
        static private string _exeResult = "Exception";

        private static void Run(BenchLogOptions options)
        {
            if (options == null)
            {
                return;
            }

            options.IsAddComment = "N";

            CmdlineService cmdlineservice = new CmdlineService();

            if (options.Mode == "PreCheck")
            {
                CmdlineContext cmdlinecontext = CreateContext(options, new HashSet<string>());
                CmdlinePreCheckBenchLog(cmdlinecontext, options);
            }
            else if (options.Mode == "GenerateProgram")
            {
                CmdlineGeneratePattern(cmdlineservice, options);

                if (_exeResult == "Done")
                {
                    CmdlineContext cmdlinecontext = CreateContext(options, _addrFor64);
                    CmdlineGenerateProgram(cmdlinecontext, options);
                }
            }
        }

        private static void CmdlinePreCheckBenchLog(CmdlineContext cmdlinecontext, BenchLogOptions options)
        {
            string precheckFile = Path.Combine(cmdlinecontext.OutputDir, "PreCheck_Result.log");
            NLogMain.SetNLogInTrace(precheckFile);
            Logger logger = LogManager.GetCurrentClassLogger();
            LocalSpecs.CurrentProject = options.ProjectName;
            LocalSpecs.PinMap = options.PinMap;

            List<PatternItem> patterns = new List<PatternItem>();
            string currPat = "";
            List<string> failLogs = new List<string>();
            Dictionary<string, List<string>> functionmap = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            _exeResult = "Exception";
            try
            {
                logger.Info("Load EfuseBitDefition file ...");
                _efuseBitDef = new LoaderEfuseBitDef(cmdlinecontext.EfuseBitDefinitionPath, "");
                if (!string.IsNullOrEmpty(cmdlinecontext.EfuseBitDefinitionPath))
                {
                    string projectFilepath = Path.Combine(Directory.GetCurrentDirectory(), "Settings", $"EfuseScriptDescription_default.xml");
                    EfuseScriptConfig config = EfuseAlgorithmCheck.LoadEfuseConfigFile(projectFilepath);
                    _efuseBitDef.Parse(config);
                }

                logger.Info("Load PinMap file ...");
                if (!string.IsNullOrEmpty(cmdlinecontext.PinMapPath))
                {
                    _pinmap = new ReadPinMapSheet().GetSheet(cmdlinecontext.PinMapPath);
                }
                if (_efuseBitDef.BitDefTable.Rows.Count() == 0 || _pinmap == null)
                {
                    throw new Exception("Input file error occurred");
                }

                logger.Info("Precheck BenchLog ...");

                string[] allOverWriteLogs = options.OverWriteLogs.Split(',');


                HashSet<string> namingDictionary = new HashSet<string>();
                int index = 0;
                foreach (string filename in Directory.GetFiles(cmdlinecontext.BenchLogPath, "*.csv"))
                {
                    logger.Info("Getting " + filename + " ...");

                    MappingItem info = LogCheckBusiness.GetNamingInfo(filename);

                    var item = new PatternItem();
                    item.BenchLog = Path.GetFileNameWithoutExtension(filename);
                    item.Name = Path.GetFileNameWithoutExtension(filename);
                    item.PatName = info.Log;
                    item.Status = "Idle";
                    item.FilePath = Path.Combine(Path.GetDirectoryName(filename), item.PatName);
                    item.Address = filename;
                    item.Init.AddRange(info.Inits);
                    item.Index = index;
                    item.IsOverWrite = options.OverWriteLogs == "Y" || allOverWriteLogs.Contains(Path.GetFileName(filename));

                    if (namingDictionary.Contains(item.Address))
                    {
                        continue;
                    }

                    patterns.Add(item);

                    namingDictionary.Add(item.Address);
                    index++;
                }

                var logs = new List<LogCheckBusiness>();

                logger.Info("Check init pattern ...");
                CmdlineService.CheckInit(cmdlinecontext.BenchLogPath, patterns.GroupBy(p => p.Name).ToDictionary(p => p.Key, p => p.FirstOrDefault().Init));
                CmdlineService.GetSetup(cmdlinecontext.ProjectName);
                foreach (PatternItem pat in patterns)
                {
                    logger.Info("Checking " + pat.Name + " ...");

                    currPat = pat.Name;

                    var logcheck = new LogCheckBusiness();
                    logcheck.EfusebitDefRows = _efuseBitDef.BitDefTable.Rows;
                    logcheck.Pinmap = _pinmap;
                    logcheck.FunctionMaps = functionmap;
                    logs.Add(logcheck);

                    var totalLogs = new List<string>();
                    totalLogs.AddRange(CmdlineService.CmdlinePackageBenchLogChecker(cmdlinecontext, logcheck, pat));
                }

                logger.Info("Cross check logs ...");
                CmdlineService.CrossCheck(logs);

                foreach (RFLogError err in RFLogManager.Errors)
                {
                    string logerrmsg = string.Format("Log:" + err.LogFile + ", ErrorType:" + err.Type.ToString());
                    if (err.Level == EnumErrorLevel.Error)
                    {
                        logger.Error(logerrmsg);
                    }
                    else if (err.Level == EnumErrorLevel.Warning)
                    {
                        logger.Warn(logerrmsg);
                    }
                    else
                    {
                        logger.Info(logerrmsg);
                    }
                }

                logger.Info("Generate error report ...");
                CmdlineService.GenerateErrorReportNew(cmdlinecontext.OutputDir, logs.SelectMany(p => p.LogFile).ToList());
                _exeResult = "Done";
            }
            catch (Exception e)
            {
                if (string.IsNullOrEmpty(currPat))
                {
                    logger.Error("Exception message:" + e.Message);
                }
                else
                {
                    bool failPatFlag = false;
                    foreach (PatternItem pat in patterns)
                    {
                        if (failPatFlag)
                        {
                            failLogs.Add(pat.Name);
                        }

                        if (pat.Name == currPat)
                        {
                            failPatFlag = true;
                        }
                    }

                    logger.Error("Exception message:" + e.Message + " in " + currPat);
                }
            }
            finally
            {
                failLogs.AddRange(RFLogManager.Errors.Where(err => err.Level == EnumErrorLevel.Error).Select(err => err.LogFile).Distinct());
                foreach (PatternItem pat in patterns)
                {
                    if (!failLogs.Contains(pat.Name))
                    {
                        logger.Info("Benchlog_" + pat.Name + ":" + "PASS");
                    }
                    else
                    {
                        logger.Info("Benchlog_" + pat.Name + ":" + "FAIL");
                    }
                }
                logger.Info("The execution result is:" + _exeResult);
            }

        }
        private static void CmdlineGeneratePattern(CmdlineService cmdlineservice, BenchLogOptions options)
        {
            string patterngenFile = Path.Combine(options.Output, "PatternGen_Result.log");
            NLogMain.SetNLogInTrace(patterngenFile);
            Logger logger = LogManager.GetCurrentClassLogger();

            _exeResult = "Exception";
            try
            {
                logger.Info("Load EfuseBitDefition file ...");
                if (!string.IsNullOrEmpty(options.EfuseBitDefition))
                {
                    _efuseBitDef = new LoaderEfuseBitDef(options.EfuseBitDefition, "");
                    string projectFilepath = Path.Combine(Directory.GetCurrentDirectory(), "Settings", $"EfuseScriptDescription_default.xml");
                    EfuseScriptConfig config = EfuseAlgorithmCheck.LoadEfuseConfigFile(projectFilepath);
                    _efuseBitDef.Parse(config);
                }

                logger.Info("Load PinMap file ...");
                if (!string.IsNullOrEmpty(options.PinMap))
                {
                    _pinmap = new ReadPinMapSheet().GetSheet(options.PinMap);
                }

                logger.Info("Load 64bit_registers file ...");
                _addrFor64 = new HashSet<string>();
                PopulateAddrFor64(options.AddrFor64);

                CmdlineContext cmdlinecontext = CreateContext(options, _addrFor64);

                if (_efuseBitDef.BitDefTable.Rows.Count() == 0 || _pinmap == null)
                {
                    throw new Exception("Input file error occurred");
                }

                logger.Info("Generate Pattern ...");

                GenerateType selGenType = cmdlinecontext.PatternType == "ARF/FW" ? GenerateType.ARF : GenerateType.HTOL;

                BindingList<PatternItem> patternList = new BindingList<PatternItem>();
                Dictionary<string, MappingItem> namingDictionary = new Dictionary<string, MappingItem>();
                string[] silicons = cmdlinecontext.SiliconVersion.Split(',');
                string[] allOverWriteLogs = options.OverWriteLogs.Split(',');

                if (Directory.Exists(cmdlinecontext.BenchLogPath) && silicons.Count() == 2)
                {
                    BuildPatternList(cmdlinecontext, options, selGenType, silicons, allOverWriteLogs, patternList, namingDictionary, logger);

                    #region multi thread
                    Dictionary<string, List<PatternGenItem>> patGenItems = new Dictionary<string, List<PatternGenItem>>();
                    var patconv = new PatternConv();
                    var infoRef = new List<string>();

                    double jtagFreq = double.Parse(cmdlinecontext.JtagFreq);
                    int taskCounts = patternList.Count;
                    var toDoLists = new List<Task>();
                    var toTallogs = new List<string>();

                    var tasks = new List<Task<PatternResult>>();
                    var finalWriteSrc = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    logger.Info("Get " + cmdlinecontext.ProjectName + " setup ...");

                    CmdlineService.GetSetup(cmdlinecontext.ProjectName);

                    CreatePatternTasks(cmdlineservice, cmdlinecontext, selGenType, jtagFreq, taskCounts, patternList, namingDictionary, patGenItems, infoRef, tasks, logger);
                    #endregion

                    PatternResult[] results = Task.WhenAll(tasks).Result;

                    #region New Strcuture write, compress and compile pattern.
                    WritePatternFiles(results);
                    var allFiles = results
                        .SelectMany(r => r.Files)
                        .GroupBy(f => f.FileName)
                        .Select(g => g.First())
                        .ToList();
                    bool needGlobal = results.Any(r => r.NeedGlobalSub);

                    string outputDir = results.First().OutputDir;
                    CmdlineService.CompilePat(cmdlinecontext, allFiles, outputDir, patconv, logger);

                    #endregion

                    MergeWriteSrcRows(results, finalWriteSrc, toTallogs);
                    toTallogs = toTallogs.Distinct().ToList();
                    if (finalWriteSrc.Count != 0)
                    {
                        GenerateWriteSrcFile(finalWriteSrc, options.Output);
                    }

                    logger.Info("Writing HardIPInfo.log ...");
                    string hardipInfoFiles = Path.Combine(cmdlinecontext.OutputDir, "HardIPInfo.log");
                    if (File.Exists(hardipInfoFiles))
                    {
                        File.Delete(hardipInfoFiles);
                    }

                    List<string> txtFiles = new List<string>();
                    foreach (string info in infoRef)
                    {
                        txtFiles.AddRange(Directory.GetFiles(info, "*.txt", SearchOption.AllDirectories));
                    }

                    List<string> fileContents = new List<string>();
                    object lockfileContents = new object();
                    Parallel.ForEach(txtFiles, txtFile =>
                    {
                        string content = File.ReadAllText(txtFile);
                        lock (lockfileContents)
                        {
                            fileContents.Add(content);
                        }
                    });

                    fileContents.Sort();
                    string finalContent = string.Join("", fileContents);
                    File.WriteAllText(hardipInfoFiles, finalContent);

                    options.PatternInfo = hardipInfoFiles;

                    logger.Info("Update pat set, pat subr, init sets, gen list, LUT items ...");
                    var scgh = new Dictionary<string, string>();
                    var patset = new Dictionary<string, string>();
                    var patsubr = new Dictionary<string, List<string>>();
                    var initSets = new Dictionary<string, List<string>>();
                    var generatedList = new List<string>();
                    var lutItems = new List<LutItem>();
                    CollectPatternGenItems(patGenItems, scgh, patset, patsubr, initSets, generatedList, lutItems);

                    logger.Info("Generate PatSet Info ...");
                    PatternGenBusiness.GeneratePatSetInfo(cmdlinecontext.OutputDir, patset, patsubr, cmdlinecontext.PatternFolder);
                    logger.Info("Generate LUT ...");
                    PatternGenBusiness.GenerateLUT(cmdlinecontext.OutputDir, lutItems);
                    logger.Info("Generate Related SCGH ...");
                    options.Scgh = PatternGenBusiness.GenerateRelatedSCGH(cmdlinecontext.ProjectName, scgh, initSets, cmdlinecontext.OutputDir, "ARF");
                    logger.Info("Generate CPP Table ...");
                    PatternGenBusiness.GenerateCPPTable(cmdlinecontext.OutputDir, patGenItems.SelectMany(p => p.Value).SelectMany(p => p.Cpp).ToList());

                    string key = string.Format("_{0}_", cmdlinecontext.SiliconVersion.Replace(",", ""));
                    DeleteBenchLogFolders(cmdlinecontext, key);
                }
                _exeResult = "Done";
            }
            catch (Exception e)
            {
                PrintInnerExcept(e, logger);
            }
            finally
            {
                logger.Info("The execution result is:" + _exeResult);
            }
        }

        private static void PopulateAddrFor64(string addrFor64Path)
        {
            if (!string.IsNullOrEmpty(addrFor64Path))
            {
                using (StreamReader reader = new StreamReader(addrFor64Path))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (uint.TryParse(line.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out uint num))
                        {
                            _addrFor64.Add(Convert.ToString(num, 2));
                        }
                        else
                        {
                            throw new Exception($"Error address {line} in {addrFor64Path}");
                        }
                    }
                }
            }
        }

        private static void BuildPatternList(CmdlineContext cmdlinecontext, BenchLogOptions options, GenerateType selGenType, string[] silicons, string[] allOverWriteLogs, BindingList<PatternItem> patternList, Dictionary<string, MappingItem> namingDictionary, Logger logger)
        {
            int index = 0;
            foreach (string filepath in Directory.GetFiles(cmdlinecontext.BenchLogPath, "*.csv"))
            {
                logger.Info("Parse " + filepath + " ...");

                string filename = Path.GetFileName(filepath);

                MappingItem info = PatternGenBusiness.GetNamingInfo(
                    filepath, selGenType, silicons[0], silicons[1], cmdlinecontext.IsFullSweep);
                var item = new PatternItem();
                item.BenchLog = Path.GetFileNameWithoutExtension(cmdlinecontext.BenchLogPath);
                item.Name = info.Log;
                item.PatName = info.GetFullPattern();
                item.Status = "Idle";
                item.FilePath = Path.Combine(Path.GetDirectoryName(filepath), item.PatName);
                item.Address = filepath;
                item.Init.AddRange(info.Inits);
                item.Index = index;
                item.IsOverWrite = options.OverWriteLogs == "Y" || allOverWriteLogs.Contains(filename);

                if (namingDictionary.ContainsKey(item.Address))
                {
                    continue;
                }

                patternList.Add(item);

                namingDictionary.Add(item.Address, info);
                index++;
            }
        }

        private static void CreatePatternTasks(CmdlineService cmdlineservice, CmdlineContext cmdlinecontext, GenerateType selGenType, double jtagFreq, int taskCounts, BindingList<PatternItem> patternList, Dictionary<string, MappingItem> namingDictionary, Dictionary<string, List<PatternGenItem>> patGenItems, List<string> infoRef, List<Task<PatternResult>> tasks, Logger logger)
        {
            for (int i = 0; i < taskCounts; i++)
            {
                logger.Info("Generate " + patternList[0].Address + " ...");

                PatternItem patternFile = patternList[0];
                patternList.RemoveAt(0);

                string folder = Path.GetDirectoryName(patternFile.FilePath);
                infoRef.Add(folder);

                List<PatternGenItem> genItems = new List<PatternGenItem>();
                patGenItems[patternFile.FilePath] = genItems;

                var patgen = new PatternGenBusiness(
                    cmdlinecontext.RegisterMapPath,
                    ref genItems,
                    selGenType,
                    jtagFreq,
                    patternFile.IsOverWrite,
                    cmdlinecontext.AddComment == "Y");

                patgen.EfusebitDefRows = _efuseBitDef.BitDefTable.Rows;

                if (!namingDictionary.TryGetValue(patternFile.Address, out MappingItem mappingItem))
                {
                    throw new KeyNotFoundException($"MappingItem not found. Address = {patternFile.Address}");
                }

                PatternItem localPatternFile = patternFile;
                PatternGenBusiness localPatgen = patgen;
                PatternGenBusiness.CurrDateCode = cmdlinecontext.TimeStamp;
                MappingItem localMappingItem = mappingItem;

                tasks.Add(Task.Run(() => cmdlineservice.CreatPatternContent(cmdlinecontext, localPatgen, localPatternFile, localMappingItem, logger)));
            }
        }

        private static void WritePatternFiles(PatternResult[] results)
        {
            foreach (PatternResult r in results)
            {
                foreach (PatternFile file in r.Files)
                {
                    string path = Path.Combine(r.OutputDir, file.FileName);

                    if (!File.Exists(path))
                    {
                        File.WriteAllLines(path, file.Content);
                    }

                }
            }
        }

        private static void MergeWriteSrcRows(PatternResult[] results, Dictionary<string, string> finalWriteSrc, List<string> toTallogs)
        {
            foreach (PatternResult r in results)
            {
                foreach (KeyValuePair<string, string> kv in r.WriteSrcRows)
                {
                    finalWriteSrc[kv.Key] = kv.Value;
                }
                toTallogs.AddRange(r.Files.Select(f => f.FileName));
            }
        }

        private static void CollectPatternGenItems(Dictionary<string, List<PatternGenItem>> patGenItems, Dictionary<string, string> scgh, Dictionary<string, string> patset, Dictionary<string, List<string>> patsubr, Dictionary<string, List<string>> initSets, List<string> generatedList, List<LutItem> lutItems)
        {
            foreach (PatternGenItem patitem in patGenItems.SelectMany(p => p.Value))
            {
                if (generatedList.Contains(patitem.ScghName.Key))
                {
                    continue;
                }

                generatedList.Add(patitem.ScghName.Key);
                scgh.Add(patitem.ScghName.Key, patitem.ScghName.Value);
                patset.Add(patitem.Patset.Key, patitem.Patset.Value);
                foreach (KeyValuePair<string, string> item in patitem.PatSubr)
                {
                    if (!patsubr.ContainsKey(item.Key))
                    {
                        patsubr.Add(item.Key, new List<string>());
                    }

                    if (!patsubr[item.Key].Contains(item.Value))
                    {
                        patsubr[item.Key].Add(item.Value);
                    }
                }

                initSets.Add(patitem.Pattern, patitem.InitDictionary);
                if (patitem.LutItem != null)
                {
                    lutItems.Add(patitem.LutItem);
                }
            }
        }

        private static void DeleteBenchLogFolders(CmdlineContext cmdlinecontext, string key)
        {
            Parallel.ForEach(Directory.GetDirectories(cmdlinecontext.BenchLogPath), item =>
            {
                if (Regex.IsMatch(item, key, RegexOptions.IgnoreCase))
                {
                    Directory.Delete(item, true);
                }
                else if (Regex.IsMatch(item, "Global", RegexOptions.IgnoreCase))
                {
                    Directory.Delete(item, true);
                }
            }
            );
        }

        private static void PrintInnerExcept(Exception e, Logger logger)
        {
            logger.Error("Exception message:" + e.Message);

            if (e.InnerException != null)
            {
                PrintInnerExcept(e.InnerException, logger);
            }

        }

        private static void CmdlineGenerateProgram(CmdlineContext cmdlinecontext, BenchLogOptions options)
        {
            string programgenFile = Path.Combine(options.Output, "ProgramGen_Result.log");
            NLogMain.SetNLogInTrace(programgenFile);
            Logger logger = LogManager.GetCurrentClassLogger();
            logger.Info("Generate Program ...");

            _exeResult = "Exception";
            try
            {
                LocalSpecs.IsBenchLog = true;
                LocalSpecs.BasLibraryFolder = options.LibraryPath;
                LocalSpecs.CurrentProject = options.ProjectName;
                LocalSpecs.TestPlanFileName = options.TestPlanShell;
                LocalSpecs.TimeSetFolder = options.Output;
                LocalSpecs.PatternFolder = options.Output;
                LocalSpecs.ScghFileName = options.Scgh;
                LocalSpecs.HardIpInfoFileName = options.PatternInfo;
                LocalSpecs.TarFolder = options.Output;
                LocalSpecs.CppDspControl = options.CppDspControl;
                LocalSpecs.TarFolder = options.Output;
                LocalSpecs.BasLibraryFolder = options.LibraryPath;
                ProjectConfigSingleton.Instance().LoadProjectConfig(LocalSpecs.ProjectIniFileName);
                LocalSpecs.Options = new Options(ProjectConfigSingleton.Instance());

                CmdlineService.CmdlineGenTemplate(cmdlinecontext);
                _exeResult = "Done";
            }
            catch (Exception e)
            {
                logger.Error("Exception message:" + e.Message);
            }
            finally
            {
                logger.Info("The execution result is:" + _exeResult);
            }
        }

        private static CmdlineContext CreateContext(BenchLogOptions options, HashSet<string> addrFor64)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            return new CmdlineContext
            {
                BenchLogPath = options.BenchLog,
                OutputDir = options.Output,
                RegisterMapPath = options.RegisterMap,
                PinMapPath = options.PinMap,
                EfuseBitDefinitionPath = options.EfuseBitDefition,

                ProjectName = options.ProjectName,
                SiliconVersion = options.ProjectSilicon,
                PatternFolder = options.PatFolder,

                PatternType = options.PatternType,
                PreSetupPatterns = options.PreSetupPatterns,
                TimeStamp = options.TimeStamp,
                JtagFreq = options.JTAGFreq,

                IsR16 = options.IsR16 == "Y",
                IsFullSweep = options.IsFullSweep == "Y",
                AddComment = options.IsAddComment,
                DebugMode = "N",

                TestPlanShellPath = options.TestPlanShell,
                LibraryPath = options.LibraryPath,
                CppDspControl = options.CppDspControl,

                PatternInfoPath = options.PatternInfo,
                ScghPath = options.Scgh,

                AddrFor64InBin = addrFor64 ?? new HashSet<string>()
            };
        }

        private static void GenerateWriteSrcFile(Dictionary<string, string> writeSrc_Rows, string outputDir)
        {
            string tmpoutputDir = Path.Combine(outputDir, "StoreNameBW.txt");
            var writeSrcFile = new StreamWriter(tmpoutputDir);
            writeSrcFile.WriteLine($"EFUSE_NNAME, BIT_WIDTH");
            foreach (KeyValuePair<string, string> row in writeSrc_Rows)
            {
                writeSrcFile.WriteLine($"{row.Key}, {row.Value}");
            }
            writeSrcFile.Close();
        }
    }
}
