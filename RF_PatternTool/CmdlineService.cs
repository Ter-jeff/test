using System.Collections.Concurrent;
using System.Drawing;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.PreAction.ReadBasLib;
using Automation.InputManager.Data;
using Automation.Library;
using Automation.PreCheck.AllParaData;
using Automation.Static;

using CommonLib.ErrorReport.Base;
using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using EfuseCheckCmdLib.EFuse;

using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using LogLib.Static;
using LogLib.Utility;

using NLog;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using PatternCompile;

using RF_PatternTool.PatternGen;
using RF_PatternTool.Template;
using RF_PatternTool.VbtGen;

using RfLib.Dvdc.GenFlow;
using RfLib.Dvdc.GenTemplate;
using RfLib.Dvdc.Reader.CapturePostProcess;

using TestPlanLib.Efuse;
using TestPlanLib.Singleton;
using TestPlanLib.VbtLib;

namespace RF_PatternTool
{
    public class CmdlineService
    {
        public static string ExeDir = AppContext.BaseDirectory;
        private PinMapSheet _pinmap = null;
        public static void GetSetup(string project)
        {
            string templatePath = Path.Combine(ExeDir, "template", project.ToLower());
            if (!Directory.Exists(templatePath))
            {
                templatePath = Path.Combine(ExeDir, "template", "proxima");
            }

            List<Task> alltasks = new List<Task>();

            var loadBodyTemp = new Task(() =>
            {
                TemplateSet.BodyTemp = File.ReadAllText(Path.Combine(templatePath, "body.txt"));
            });
            var loadReadTemp = new Task(() =>
            {
                TemplateSet.ReadTemp = File.ReadAllText(Path.Combine(templatePath, "Read.txt"));
            });
            var loadSubrTemp = new Task(() =>
            {
                TemplateSet.SubrTemp = File.ReadAllText(Path.Combine(templatePath, "subroutine.txt"));
            });
            var loadWriteTemp = new Task(() =>
            {
                TemplateSet.WriteTemp = File.ReadAllText(Path.Combine(templatePath, "Write.txt"));
            });
            var loadMloopTemp = new Task(() =>
            {
                TemplateSet.MloopTemp = File.ReadAllText(Path.Combine(templatePath, "matchLoop.txt"));
            });

            loadBodyTemp.Start();
            loadReadTemp.Start();
            loadSubrTemp.Start();
            loadWriteTemp.Start();
            loadMloopTemp.Start();

            alltasks.Add(loadBodyTemp);
            alltasks.Add(loadReadTemp);
            alltasks.Add(loadSubrTemp);
            alltasks.Add(loadWriteTemp);
            alltasks.Add(loadMloopTemp);

            if (File.Exists(Path.Combine(templatePath, "Read64.txt")))
            {
                var loadRead64Temp = new Task(() =>
                {
                    TemplateSet.Read64Temp = File.ReadAllText(Path.Combine(templatePath, "Read64.txt"));
                });
                loadRead64Temp.Start();
                alltasks.Add(loadRead64Temp);
            }
            if (File.Exists(Path.Combine(templatePath, "Write64.txt")))
            {
                var loadWrite64Temp = new Task(() =>
                {
                    TemplateSet.Write64Temp = File.ReadAllText(Path.Combine(templatePath, "Write64.txt"));
                });
                loadWrite64Temp.Start();
                alltasks.Add(loadWrite64Temp);
            }

            Task.WaitAll(alltasks.ToArray());
        }
        public static void CrossCheck(List<LogCheckBusiness> logs)
        {

            #region check duplicate pattern name

            CheckDuplicatePatternName(logs);

            #endregion

            #region check duplicate test name

            CheckDuplicateTestName(logs);

            #endregion

            #region check duplicate store name

            CheckDuplicateStoreName(logs);

            #endregion

            #region check duplicate calc store name

            CheckDuplicateCalcStoreName(logs);

            #endregion

            #region check duplicate meas store name

            CheckDuplicateMeasStoreName(logs);

            #endregion

        }

        private static void CheckDuplicatePatternName(List<LogCheckBusiness> logs)
        {
            var patNameGrps = new Dictionary<string, List<string>>();
            foreach (LogCheckBusiness log in logs)
            {
                foreach (string patname in log.PatNames)
                {
                    if (!patNameGrps.ContainsKey(patname))
                    {
                        patNameGrps.Add(patname, new List<string>());
                    }

                    patNameGrps[patname].Add(log.LogName);
                }
            }
            foreach (KeyValuePair<string, List<string>> patNameIssue in patNameGrps.Where(p => p.Value.Count > 1))
            {
                if (patNameIssue.Value.Distinct().Count() == 1)
                {
                    continue;
                }

                foreach (string dic in patNameIssue.Value)
                {
                    string otherSheets = string.Join(",", patNameIssue.Value.Except(new List<string> { dic }));
                    var err = new RFLogError();
                    err.Type = ErrorType.DuplicatePatternNameOutside;
                    err.Level = EnumErrorLevel.Error;
                    err.Message = string.Format("PatName : {0} \r\nDuplicate on logger : {1}", patNameIssue.Key, otherSheets);
                    RFLogManager.Push(err, dic, 0, "INTERFACE");
                }
            }
        }

        private static void CheckDuplicateTestName(List<LogCheckBusiness> logs)
        {
            var testNameGrps = new Dictionary<string, List<string>>();
            foreach (LogCheckBusiness log in logs)
            {
                foreach (string tn in log.LogTestName)
                {
                    if (!testNameGrps.ContainsKey(tn))
                    {
                        testNameGrps.Add(tn, new List<string>());
                    }

                    testNameGrps[tn].Add(log.LogName);
                }
            }
            foreach (KeyValuePair<string, List<string>> tnIssue in testNameGrps.Where(p => p.Value.Count > 1))
            {
                foreach (string dic in tnIssue.Value)
                {
                    string otherSheets = string.Join(",", tnIssue.Value.Except(new List<string> { dic }));
                    var err = new RFLogError();
                    err.Type = ErrorType.DuplicateTestNameOutside;
                    err.Level = EnumErrorLevel.Error;
                    err.Message = string.Format("TestName : {0} \r\nDuplicate on logger : {1}", tnIssue.Key, otherSheets);
                    RFLogManager.Push(err, dic, 0, "TESTNAME");
                }
            }
        }

        private static void CheckDuplicateStoreName(List<LogCheckBusiness> logs)
        {
            var storeNameGrps = new Dictionary<string, List<string>>();
            foreach (LogCheckBusiness log in logs)
            {
                foreach (string sn in log.LogStoreName)
                {
                    if (!storeNameGrps.ContainsKey(sn))
                    {
                        storeNameGrps.Add(sn, new List<string>());
                    }

                    storeNameGrps[sn].Add(log.LogName);
                }
            }
            foreach (KeyValuePair<string, List<string>> snIssue in storeNameGrps.Where(p => p.Value.Count > 1))
            {
                foreach (string dic in snIssue.Value)
                {
                    string otherSheets = string.Join(",", snIssue.Value.Except(new List<string> { dic }));
                    var err = new RFLogError();
                    err.Type = ErrorType.DuplicateStoreNameOutside;
                    err.Level = EnumErrorLevel.Error;
                    err.Message = string.Format("StoreName : {0} \r\nDuplicate on logger : {1}", snIssue.Key, otherSheets);
                    RFLogManager.Push(err, dic, 0, "POSTPROCESS");
                }
            }
        }

        private static void CheckDuplicateCalcStoreName(List<LogCheckBusiness> logs)
        {
            var calcStoreNameGrps = new Dictionary<string, List<string>>();
            foreach (LogCheckBusiness log in logs)
            {
                foreach (string csn in log.LogCalcStoreName)
                {
                    if (!calcStoreNameGrps.ContainsKey(csn))
                    {
                        calcStoreNameGrps.Add(csn, new List<string>());
                    }

                    calcStoreNameGrps[csn].Add(log.LogName);
                }
            }
            foreach (KeyValuePair<string, List<string>> csnIssue in calcStoreNameGrps.Where(p => p.Value.Count > 1))
            {
                foreach (string dic in csnIssue.Value)
                {
                    string otherSheets = string.Join(",", csnIssue.Value.Except(new List<string> { dic }));
                    var err = new RFLogError();
                    err.Type = ErrorType.DuplicateCalcStoreNameOutside;
                    err.Level = EnumErrorLevel.Error;
                    err.Message = string.Format("CalcStoreName : {0} \r\nDuplicate on logger : {1}", csnIssue.Key, otherSheets);
                    RFLogManager.Push(err, dic, 0, "POSTPROCESS");
                }
            }
        }

        private static void CheckDuplicateMeasStoreName(List<LogCheckBusiness> logs)
        {
            var measStoreNameGrps = new Dictionary<string, List<string>>();
            foreach (LogCheckBusiness log in logs)
            {
                foreach (string msn in log.LogMeasStoreName)
                {
                    if (!measStoreNameGrps.ContainsKey(msn))
                    {
                        measStoreNameGrps.Add(msn, new List<string>());
                    }

                    measStoreNameGrps[msn].Add(log.LogName);
                }
            }
            foreach (KeyValuePair<string, List<string>> msnIssue in measStoreNameGrps.Where(p => p.Value.Count > 1))
            {
                foreach (string dic in msnIssue.Value)
                {
                    string otherSheets = string.Join(",", msnIssue.Value.Except(new List<string> { dic }));
                    var err = new RFLogError();
                    err.Type = ErrorType.DuplicateMeasStoreNameOutside;
                    err.Level = EnumErrorLevel.Error;
                    err.Message = string.Format("MeasStoreName : {0} \r\nDuplicate on logger : {1}", msnIssue.Key, otherSheets);
                    RFLogManager.Push(err, dic, 0, "POSTPROCESS");
                }
            }
        }

        public static void CheckInit(string path, Dictionary<string, List<string>> patterns)
        {
            //var resultErr = new List<Error>();
            var files = Directory.GetFiles(path).Select(Path.GetFileNameWithoutExtension).ToList();
            foreach (KeyValuePair<string, List<string>> pattern in patterns)
            {
                if (pattern.Value.Count == 0)
                {
                    continue;
                }

                var errorP = new List<string>();
                foreach (string initP in pattern.Value)
                {
                    if (!files.Exists(p => p.Equals(initP, StringComparison.OrdinalIgnoreCase)))
                    {
                        errorP.Add(initP);
                    }
                }
                if (errorP.Count > 0)
                {
                    string message = string.Format("Init Pattern: \"{0}\" . On BenchLogs: {1}. Not Exist In Directory: {2}",
                        string.Join(",", errorP), pattern.Key, path);
                    ErrorMessageBox.Show(message, "Init Patterns not Found");
                }
            }
        }

        public static List<string> CmdlinePackageBenchLogChecker(CmdlineContext cmdlinecontext, LogCheckBusiness patgen, PatternItem patternFile)
        {
            var logs = new List<string>();

            patgen.CheckBench(patternFile.Address, patternFile.IsOverWrite, new Dictionary<string, Register>(StringComparer.OrdinalIgnoreCase), cmdlinecontext.OutputDir);

            return logs;
        }
        public static void GenerateErrorReportNew(string path, List<BenchLogFile> logs)
        {
            List<RFLogError> s = RFLogManager.Errors.FindAll(p => string.IsNullOrEmpty(p.LogFile));
            var logCategory = RFLogManager.Errors.GroupBy(p => p.LogFile).ToDictionary(p => p.Key, p => p.ToList());
            Thread.Sleep(1000);
            foreach (KeyValuePair<string, List<RFLogError>> logIssue in logCategory)
            {
                if (logIssue.Key.Equals("CrossCheck", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                string logreport = Path.Combine(path, string.Format("PreChkErr_{0}_{1}.csv", logIssue.Key, DateTime.Now.ToString("yyyyMMdd_HHmmss")));

                using (var ep = new ExcelPackage(new FileInfo(logreport)))
                {
                    ExcelWorkbook wb = ep.Workbook;
                    #region Highlight cell
                    var format = new ExcelTextFormat();
                    format.Delimiter = ',';
                    format.TextQualifier = '"';
                    format.DataTypes = new[] { eDataTypes.String };
                    ExcelWorksheet ws = wb.Worksheets.Add("log");
                    int rowindex = 1;
                    List<BenchLogFile> log = logs.FindAll(p => p.LogFile.Equals(logIssue.Key, StringComparison.OrdinalIgnoreCase));
                    if (log.Count == 0)
                    {
                        continue;
                    }

                    var ssss = log.SelectMany(p => p.RawDatas).ToList();
                    foreach (string data in log.SelectMany(p => p.RawDatas))
                    {
                        ws.Cells[rowindex, 1].LoadFromText(data, format);
                        List<RFLogError> errors = logIssue.Value.FindAll(p => p.Row == rowindex);
                        foreach (RFLogError error in errors)
                        {
                            List<int> indexColumns = GetErrorColumnIndexes(log, ws, error);
                            foreach (int colIndex in indexColumns)
                            {
                                error.LogAddress = ws.Cells[rowindex, colIndex].Address;
                                ws.Cells[rowindex, colIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                if (error.Level == EnumErrorLevel.Error)
                                {
                                    ws.Cells[rowindex, colIndex].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                }

                                if (error.Level == EnumErrorLevel.Warning)
                                {
                                    ws.Cells[rowindex, colIndex].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                }
                            }
                        }

                        rowindex++;
                    }
                    #endregion

                    #region
                    WriteSummarySheet(wb, logIssue.Value);
                    #endregion
                    ep.Save();
                }
            }
        }

        private static List<int> GetErrorColumnIndexes(List<BenchLogFile> log, ExcelWorksheet ws, RFLogError error)
        {
            var indexColumns = new List<int>();
            foreach (string header in error.Header.Split(','))
            {
                //	REG_FIELD_NAME	DATA	Default	Efuse name	COMMENT	PostProcess	TESTNAME	LimitLo	LimitHi	Units
                switch (header.ToUpper())
                {
                    case "OPERATION":
                        indexColumns.Add(log.First().IndexOperation + 1);
                        break;
                    case "INTERFACE":
                        indexColumns.Add(log.First().IndexInterface + 1);
                        break;
                    case "ADDRESS":
                        indexColumns.Add(log.First().IndexAddress + 1);
                        break;
                    case "REGDATA":
                        indexColumns.Add(log.First().IndexRegData + 1);
                        break;
                    case "MSB":
                        indexColumns.Add(log.First().IndexMSB + 1);
                        break;
                    case "LSB":
                        indexColumns.Add(log.First().IndexLSB + 1);
                        break;
                    case "FIELD_VAL":
                        indexColumns.Add(log.First().IndexData + 1);
                        break;
                    case "EFUSENAME":
                        indexColumns.Add(log.First().IndexFuseName + 1);
                        break;
                    case "POSTPROCESS":
                        indexColumns.Add(log.First().IndexPostProcess + 1);
                        break;
                    case "TESTNAME":
                        indexColumns.Add(log.First().IndexTestName + 1);
                        break;
                    case "LOLIMIT":
                        indexColumns.Add(log.First().IndexLowLimit + 1);
                        break;
                    case "HILIMIT":
                        indexColumns.Add(log.First().IndexHighLimit + 1);
                        break;
                    default:
                        for (int i = 1; i <= ws.Dimension.Columns; i++)
                        {
                            indexColumns.Add(i);
                        }
                        break;
                }
            }

            return indexColumns;
        }

        private static void WriteSummarySheet(ExcelWorkbook wb, List<RFLogError> logErrors)
        {
            ExcelWorksheet ws = wb.Worksheets.Add("Summary");
            wb.Worksheets.MoveToStart("Summary");
            ws.Cells[1, 1].Value = "Category";
            ws.Cells[1, 2].Value = "Row";
            ws.Cells[1, 3].Value = "Message";
            ws.Cells[1, 4].Value = "Level";
            ws.Cells[1, 1, 1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[1, 1, 1, 4].Style.Fill.BackgroundColor.SetColor(Color.RoyalBlue);
            int rowindex = 2;
            Dictionary<ErrorType, List<RFLogError>> errorGroups = logErrors.GroupBy(p => p.Type).ToDictionary(p => p.Key, p => p.Distinct().ToList());
            foreach (KeyValuePair<ErrorType, List<RFLogError>> errorGroup in errorGroups)
            {
                ws.Cells[rowindex, 1].Value = errorGroup.Key;
                ws.Cells[rowindex, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[rowindex, 1].Style.Fill.BackgroundColor.SetColor(Color.Tomato);
                int collapIndex = 1;
                int groupindex = 1;
                foreach (RFLogError error in errorGroup.Value)
                {
                    ws.Cells[rowindex, 2].Value = error.Row;
                    ws.Cells[rowindex, 3].IsRichText = true;
                    ws.Cells[rowindex, 3].RichText.Add(error.Message);
                    if (!string.IsNullOrEmpty(error.LogAddress))
                    {
                        var link = new Uri(string.Format("#\'log\'!{0}", error.LogAddress), UriKind.Relative);
                        ws.Cells[rowindex, 3].Hyperlink = link;
                    }
                    ws.Cells[rowindex, 4].Value = error.Level;

                    ws.Cells[rowindex, 2, rowindex, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[rowindex, 2, rowindex, 4].Style.Fill.BackgroundColor.SetColor(Color.Transparent);
                    if (error.Level == EnumErrorLevel.Error)
                    {
                        ws.Cells[rowindex, 4].Style.Fill.BackgroundColor.SetColor(Color.Red);
                    }
                    else
                    {
                        ws.Cells[rowindex, 4].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    }

                    if (groupindex != 1)
                    {
                        ws.Row(rowindex).Collapsed = true;
                        ws.Row(rowindex).OutlineLevel = collapIndex;
                    }
                    ws.Cells[rowindex, 1, rowindex, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    ws.Cells[rowindex, 1, rowindex, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    groupindex++;
                    rowindex++;
                }
                collapIndex++;

            }
            for (int i = 1; i <= ws.Dimension.Columns; i++)
            {
                ws.Column(i).TryAutoFit();
            }
        }
        public static void CmdlineGenTemplate(CmdlineContext cmdlinecontext)
        {

            string scghpath = cmdlinecontext.ScghPath;
            string patinfopath = cmdlinecontext.PatternInfoPath;
            //Read HardIp Info
            HardIpParaData paraData = new HardIpParaData(EnumBlock.Dvdc);
            HardIpInputData hardIpInputData = new HardIpInputData(paraData);
            var cppReader = new PostProcessSheet();
            var cppsetups = new List<CPPSetup>();
            foreach (string file in Directory.GetFiles(cmdlinecontext.OutputDir))
            {
                if (Regex.IsMatch(Path.GetFileNameWithoutExtension(file), "CPP", RegexOptions.IgnoreCase))
                {
                    cppReader.Read(file);
                    cppsetups.AddRange(cppReader.Setups);
                }
            }
            TemplateGeneratorMain tempGen = new TemplateGeneratorMain(hardIpInputData);

            tempGen.Cpps = cppsetups.ToDictionary(p => p.Pattern, p => p.Datas);

            TestProgram.NonIgxlSheetsList.SheetList.Clear();
            tempGen.WorkFlow(paraData, true, null);

            #region library loading
            string targetVBTDir = Path.Combine(LocalSpecs.TarFolder, "Library");

            #region Generate VBT_LIB_CPP_Main.bas and DSP_CPP_Tool.bas for CPPs
            if (tempGen.Cpps.Any())
            {
                if (!Directory.Exists(targetVBTDir))
                {
                    Directory.CreateDirectory(targetVBTDir);
                }
                // ProximaW release for next Json syntax 
                if (LocalSpecs.CppDspControl == "Y")
                {
                    var efuseBitDef = new LoaderEfuseBitDef(cmdlinecontext.EfuseBitDefinitionPath, "");
                    if (!string.IsNullOrEmpty(cmdlinecontext.EfuseBitDefinitionPath))
                    {
                        string projectFilepath = Path.Combine(Directory.GetCurrentDirectory(), "Settings", $"EfuseScriptDescription_default.xml");
                        EfuseScriptConfig config = EfuseAlgorithmCheck.LoadEfuseConfigFile(projectFilepath);
                        efuseBitDef.Parse(config);
                    }

                    List<List<string>> cppFunContent = new List<List<string>>();
                    VbtLibCppMain cppFunctions = new VbtLibCppMain();
                    VbtLibCppMain.Write(cppFunContent, tempGen.Cpps.Values.ToList(), efuseBitDef.BitDefTable.Rows);

                    for (int i = 0; i < cppFunContent.Count(); i++)
                    {
                        string vbtname = "VBT_LIB_CPP_Main_" + i;
                        VbtLibCppMain.AddHeader(cppFunContent[i], vbtname);
                        ClassVbt.WriteVBFile(Path.Combine(targetVBTDir, vbtname + ".bas"), cppFunContent[i]);
                    }

                    List<List<string>> cppToolContent = new List<List<string>>();
                    DSPCppTool cppTools = new DSPCppTool();
                    DSPCppTool.Write(cppToolContent, tempGen.Cpps.Values.ToList());

                    for (int i = 0; i < cppToolContent.Count(); i++)
                    {
                        string vbtname = "DSP_CPP_Tool_" + i;
                        DSPCppTool.AddHeader(cppToolContent[i], vbtname);
                        ClassVbt.WriteVBFile(Path.Combine(targetVBTDir, vbtname + ".bas"), cppToolContent[i]);
                    }
                }
            }
            #endregion

            var basmain = new BasMain();
            if (Directory.Exists(LocalSpecs.BasLibraryFolder))
            {
                (List<Function> functions, _) = basmain.WorkFlow(targetVBTDir);
                TestProgram.VbtFunctionLib.AddVbtFunctionRange(functions);
            }
            #endregion

            Response.Report("Test plan template done...");
            Response.Report("Start Generate Test program...");

            Response.Report("Generating Instance ...");
            List<InstanceSheet> instSheets = new InstanceGenerator(hardIpInputData).GenInst(tempGen.HardIpInputData.PlanDic);

            Response.Report(string.Format("Generating Testflow ..."));
            List<SubFlowSheet> flowSheets = new DvdcFlowGenerator(hardIpInputData).GenFlow(tempGen.HardIpInputData.PlanDic);

            Response.Report("Generating BinTable ...");
            using (var ep = new ExcelPackage(new FileInfo(LocalSpecs.TestPlanFileName)))
            {
                BinNumberSingleton.Instance.Initialize(ep.Workbook);
            }

            BinTableSheet binTableSheet = new BinTableSheetGenerator(hardIpInputData).GenBinTable(tempGen.HardIpInputData.PlanDic);
            TestProgram.IgxlWorkBk.AddBinTblSheet(FolderStructure.TarDir, binTableSheet);

            //Add Init Flag Flow
            Response.Report("Generating InitFlag flow ...");
            SubFlowSheet initFlow = new InitFlagGenerator().GenInitFlag(hardIpInputData.PlanDic, binTableSheet);
            TestProgram.IgxlWorkBk.AddSubFlowSheet(LocalSpecs.TarFolder, initFlow);

            //Add SubFlow
            foreach (SubFlowSheet flowSheet in flowSheets)
            {
                //Delete jobs if All jobs are enable and replace job name in job column by actual job in config
                DataConvertor.FilterFlowJobs(flowSheet);
                if (flowSheet.Rows.Count > 0)
                {
                    TestProgram.IgxlWorkBk.AddSubFlowSheet(LocalSpecs.TarFolder, flowSheet);
                }
            }

            //Add Instance sheet
            foreach (InstanceSheet instSheet in instSheets)
            {
                if (instSheet.Rows.Count > 0)
                {
                    TestProgram.IgxlWorkBk.AddInsSheet(LocalSpecs.TarFolder, instSheet);
                }
            }
            Response.Report("Generating Print all sheets ...");
            TestProgram.IgxlWorkBk.PrintAllSheets("10.20");
            Response.Report("Generating Sheets done ...");
        }
        public PatternResult CreatPatternContent(CmdlineContext cmdlinecontext, PatternGenBusiness patgen, PatternItem patternFile, MappingItem mappingItem, Logger logger)
        {
            PatternResult atpfile;

            logger.Info("Load Registers ...");
            Dictionary<string, Register> registermap = PatternGenBusiness.LoadRegisters(cmdlinecontext.RegisterMapPath);
            if (!PatternGenBusiness.IsCheck)
            {
                PatternGenBusiness.IsCheck = true;
            }

            string[] silicons = cmdlinecontext.SiliconVersion.Split(',');
            if (silicons.Length != 2)
            {
                return new PatternResult();
            }

            if (_pinmap == null && !string.IsNullOrEmpty(cmdlinecontext.PinMapPath))
            {
                _pinmap = new ReadPinMapSheet().GetSheet(cmdlinecontext.PinMapPath);
            }
            HashSet<string> pinsName = _pinmap.PinList.Select(pin => pin.PinName).Where(name => !string.IsNullOrWhiteSpace(name)).ToHashSet(StringComparer.OrdinalIgnoreCase);
            logger.Info("Print Pattern Atp ...");
            atpfile = patgen.PrintPatternAtp(
                patternFile.Address, mappingItem, registermap, logger, cmdlinecontext.OutputDir, silicons[0],
                silicons[1], cmdlinecontext.PreSetupPatterns, pinsName, cmdlinecontext.IsR16,
                cmdlinecontext.IsFullSweep, cmdlinecontext.AddrFor64InBin);
            return atpfile;
        }

        public static void CompilePat(CmdlineContext cmdlinecontext, List<PatternFile> files, string outputDir, PatternConv patconv, Logger logger)
        {
            logger.Info("Zip Atp Files ...");

            foreach (PatternFile file in files)
            {
                string fullPath = Path.Combine(outputDir, file.FileName);
                PatternConv.ZipFile(fullPath);
            }

            logger.Info("Compile to Pattern and Zip ...");

            Parallel.ForEach(files, file =>
            {
                var localConv = new PatternConv();

                string fullPath = Path.Combine(outputDir, file.FileName);

                PatternConv.CompileToPat(fullPath, cmdlinecontext.PinMapPath);
                PatternConv.ZipFile(fullPath.Replace(".atp", ".pat"));
            });

            logger.Info("Get Pattern Info ...");

            var infos = new ConcurrentBag<string>();

            Parallel.ForEach(files, file =>
            {
                var localConv = new PatternConv();

                string fullPath = Path.Combine(outputDir, file.FileName);

                string key = "temp.txt";
                PatternConv.GetPatternInfo(cmdlinecontext.BenchLogPath, fullPath);

                string path = Path.Combine(
                    cmdlinecontext.BenchLogPath,
                    Path.GetFileNameWithoutExtension(file.FileName),
                    "temp.txt"
                );

                if (!File.Exists(path))
                {
                    key = "HardIPInfoOffline.txt";
                    path = path.Replace("temp.txt", key);
                }

                string newpath = path.Replace(
                    key,
                    Path.GetFileNameWithoutExtension(file.FileName) + ".txt"
                );

                infos.Add(newpath);

                if (File.Exists(path))
                {
                    File.Move(path, newpath);
                }
            });
        }
    }
}
