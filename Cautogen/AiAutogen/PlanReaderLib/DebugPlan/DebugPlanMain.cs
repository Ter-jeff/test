using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility;
using Cautogen.AutoCZ.CharPostProcessor.Utility.VbtModuleManager;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;
using Cautogen.common.ReaderWriter.Reader.InputReader;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using DebugPlanReaderLib.DebugPlan.DigSrc;
using DebugPlanReaderLib.DebugPlan.Mapping;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LogLib.Utility;

using OfficeOpenXml;

namespace DebugPlanReaderLib.DebugPlan
{
    public class DebugPlanMain
    {
        private const string Cz = "CZ";

        private const string ConPatternDashboard = "pattern_dashboard";
        private const string ConProcessCondition = "Process_Condition";

        private const string DebugLvccVminBoundary = "Debug_LVCC_VminBoundary";
        private const string EnableDftlhfcDebug = "Enable_DFTLHFC_Debug";
        private const string EnableFailLogDebug = "Enable_Faillog_Debug";

        public List<AiTestPlanSheet> AiTestPlanSheets = new List<AiTestPlanSheet>();
        private List<String> _aiTestPlanEnableWords = new List<string>();
        private string _testerID = string.Empty;
        private string _site0IndxFile = string.Empty;

        public List<Error> Errors = new List<Error>();
        public string InputFile;

        public PatternListSheet PatternListSheet;
        private ProcessConditionSheet ProcessCondition;
        private IgxlDataReader _igxlData;
        private string _patternFolder;
        private string _patInfoPath;
        private string _outputDir;
        private List<HardIpReference> _patInfoData;
        private bool _isCSharp = false;


        public string TesterID { get { return _testerID; } }
        public string Site0IndxFile { get { return _site0IndxFile; } }
        public List<String> AiTestPlanEnableWords { get { return _aiTestPlanEnableWords; } }

        public DebugPlanMain(string inputFile, string patternFolder, string igxlProgram = "", bool isCSharp = false, string outputDir = "")
        {
            InputFile = inputFile;
            _patternFolder = patternFolder;
            _isCSharp = isCSharp;
            _outputDir = outputDir;
            if (!string.IsNullOrEmpty(patternFolder))
            {
                //_patInfoPath = Directory.GetFiles(_patternFolder, "*", SearchOption.AllDirectories).ToList()
                //                .FirstOrDefault(x => Regex.IsMatch(Path.GetFileName(x), @".*HARDIP.*Info", RegexOptions.IgnoreCase));

                var hardipInfoList = Directory.GetFiles(_patternFolder, "*hardip*Info*.txt", SearchOption.AllDirectories);

                foreach (var hardipInfo in hardipInfoList)
                {
                    if (hardipInfoList.Count() > 0 && _patInfoData == null)
                        _patInfoData = new List<HardIpReference>();
                    _patInfoData.AddRange(new PatInfoReader(hardipInfo).Read());
                }

                //_patInfoData = string.IsNullOrEmpty(_patInfoPath) ? new List<HardIpReference>() : new PatInfoReader(_patInfoPath).Read();
            }

            _igxlData = !string.IsNullOrEmpty(igxlProgram) ? new IgxlDataReader(igxlProgram) : null;
        }

        public void Read()
        {
            using (var package = new ExcelPackage(new FileInfo(InputFile)))
            {
                var dashboard = package.Workbook.Worksheets[ConPatternDashboard];
                if (dashboard == null)
                {
                    ErrorReportManager.AddError(
                        AutoAiErrorType.E_MissingSheet_01,
                        ConPatternDashboard,
                        0,
                        0,
                        [ConPatternDashboard]);
                }
                else
                {
                    var patternListReader = new PatternListReader();
                    PatternListSheet = patternListReader.ReadSheet(dashboard);
                }

                var processCondition = package.Workbook.Worksheets[ConProcessCondition];
                if (processCondition == null)
                {
                    ErrorReportManager.AddError(
                        AutoAiErrorType.E_MissingSheet_01,
                        ConProcessCondition,
                        0,
                        0,
                        [ConProcessCondition]);
                }
                else
                {
                    var processConditionReader = new ProcessConditionReader();
                    ProcessCondition = processConditionReader.ReadSheet(processCondition);
                    _testerID = ProcessCondition.Tester;
                }

                foreach (var sheet in package.Workbook.Worksheets)
                {
                    if (sheet.Cells[1, 1].Value != null)
                    {
                        var sheetType = sheet.Cells[1, 1].Value.ToString().Split(';').First();
                        if (sheetType.Equals("AITestPlanSheet", StringComparison.CurrentCultureIgnoreCase))
                        {
                            var aITestPlanReader = new AiTestPlanReader();
                            AiTestPlanSheets.Add(aITestPlanReader.ReadSheet(sheet));

                            if (sheet.Cells[2, 1].Value != null)
                            {
                                var sheetEnableWords = sheet.Cells[2, 1].Value.ToString().Trim();
                                if (!string.IsNullOrEmpty(sheetEnableWords))
                                {
                                    var enableWordsList = sheetEnableWords.Split(';')
                                                            .Where(x => !string.IsNullOrEmpty(x.Trim()))
                                                            .Select(y => y.Trim()).ToList();
                                    _aiTestPlanEnableWords.AddRange(enableWordsList);
                                    _aiTestPlanEnableWords = _aiTestPlanEnableWords.Distinct().ToList();
                                }
                            }
                        }
                    }
                    if (sheet.Name.Equals("site0_index"))
                    {
                        var uuid = Guid.NewGuid().ToString();

                        _site0IndxFile = Path.Combine(_outputDir, $@"site0_index{uuid}.csv");
                        sheet.ExportSheet(_site0IndxFile, ",");

                    }
                }
                SetUniquePatSetName(AiTestPlanSheets);
                var allPlanPatterns = AiTestPlanSheets.SelectMany(x => x.Rows).SelectMany(x => x.Patterns);
                foreach (var planPattern in allPlanPatterns)
                {
                    var match = PatternListSheet.Rows.FirstOrDefault(x => x.Pattern.ToUpper() == planPattern.Name.ToUpper());
                    if (match != null)
                    {
                        planPattern.Version = match.PatternVersion.ToUpper().Replace(match.Pattern.ToUpper() + "_", "");
                    }
                }
                if (_igxlData != null)
                {
                    var referenceData = new MappingResult(_igxlData.InstanceSheets, _igxlData.PatSetsSheets);
                    referenceData.GenMappingIntoInstance(AiTestPlanSheets);

                    if (_patInfoData != null && _patInfoData.Any())
                    {
                        var digSrcHandler = new DigSrcHandler(_patInfoData, _isCSharp);
                        digSrcHandler.GenDigSrcIntoInstance(AiTestPlanSheets);
                    }

                }
            }
        }

        public DfcListSheet GenDfcList()
        {
            var dfcListSheet = new DfcListSheet("DFC_List");

            dfcListSheet.AddRow(new DfcRow("Test Instance"));

            dfcListSheet.AddRows(
                AiTestPlanSheets.SelectMany(x => x.Rows)
                .Where(x => x.EnumDataLoggingSettingType == EnumDataLoggingSettingType.Dfc)
                .Select(x => new DfcRow(x.GetTestName())).ToList()
                );

            return dfcListSheet;
        }

        public bool CheckAll(string patternFolder = null, string timeFolder = null)
        {
            Errors.Clear();
            foreach (var aiTestPlanSheet in AiTestPlanSheets)
            {
                aiTestPlanSheet.ClearErrors();
            }

            PatternListSheet.ClearErrors();
            ProcessCondition.ClearErrors();

            #region pre action

            var timeSets = new List<string>();
            var dcSpecs = new List<string>();
            var acSpecs = new List<string>();
            var patterns = new List<string>();
            var patternsHash = new HashSet<string>();
            var pins = new List<string>();
            var acSymbols = new List<string>();
            var checkTimeSet = false;
            var checkDcSpec = false;
            var checkAcSpec = false;
            var checkPattern = false;
            var checkPin = false;
            var checkMapping = false;
            var checkPatternInfo = false;

            if (Directory.Exists(patternFolder) && _patInfoData.Any())
            {
                checkPatternInfo = true;
            }

            if (Directory.Exists(patternFolder))
            {
                var gzs = new HashSet<string>(Directory.GetFiles(patternFolder, "*.gz", SearchOption.AllDirectories));
                foreach (var gz in gzs)
                {
                    var name = Regex.Replace(gz, ".gz$", "", RegexOptions.IgnoreCase);
                    name = Regex.Replace(name, ".atp$", "", RegexOptions.IgnoreCase);
                    name = Regex.Replace(name, ".pat$", "", RegexOptions.IgnoreCase);
                    name = Regex.Replace(name, ".patx$", "", RegexOptions.IgnoreCase);
                    patterns.Add(Path.GetFileName(name));
                }
                patternsHash = new HashSet<string>(patterns);
                checkPattern = true;
            }

            if (Directory.Exists(timeFolder))
            {
                var timeSetFiles =
                    new List<string>(Directory.GetFiles(timeFolder, "*.txt", SearchOption.AllDirectories));
                timeSets.AddRange(timeSetFiles.Select(Path.GetFileNameWithoutExtension).ToList());
                checkTimeSet = true;
            }

            if (_igxlData != null)
            {
                timeSets.AddRange(_igxlData.TimeSetBasicSheetNames);
                dcSpecs.AddRange(_igxlData.DcSpecSheets.SelectMany(x => x.CategoryList).Distinct().ToList());
                acSpecs.AddRange(_igxlData.AcSpecSheets.SelectMany(x => x.CategoryList).Distinct().ToList());
                pins = _igxlData.PinMapSheets.SelectMany(x => x.GetAllPins()).Distinct().ToList();
                acSymbols = _igxlData.AcSpecSheets.SelectMany(x => x.Rows).Where(x => !x.IsBackup)
                    .Select(x => x.Symbol).Distinct().ToList();
                checkTimeSet = true;
                checkDcSpec = true;
                checkAcSpec = true;
                checkPin = true;
                checkMapping = true;
            }

            #endregion

            #region check AiTestPlanSheets

            foreach (var aiTestPlanSheet in AiTestPlanSheets)
            {
                aiTestPlanSheet.Check();
                if (checkTimeSet)
                {
                    aiTestPlanSheet.CheckTimeSet(timeSets.Distinct().ToList());
                }

                if (checkDcSpec)
                {
                    aiTestPlanSheet.CheckDcSpec(dcSpecs);
                }

                if (checkAcSpec)
                {
                    aiTestPlanSheet.CheckAcSpec(acSpecs);
                }

                if (checkPattern && PatternListSheet != null)
                {
                    aiTestPlanSheet.CheckPattern(patternsHash, PatternListSheet);
                }

                if (checkPin)
                {
                    aiTestPlanSheet.CheckPins(pins, acSymbols);
                }

                //if (checkMapping)
                //{
                //    aiTestPlanSheet.CheckMapping();
                //}
                if (checkPatternInfo)
                {
                    aiTestPlanSheet.CheckDigSrc(_patternFolder, _patInfoData);

                }

                Errors.AddRange(aiTestPlanSheet.GetErrors());
            }

            #endregion

            //20230626 Si said didn't check folder by dashboard 
            //#region check pattern_dashboard

            //if (PatternListSheet != null && !string.IsNullOrEmpty(timeFolder))
            //{
            //    if (checkTimeSet & checkPattern)
            //    {
            //        PatternListSheet.CheckPatternTimeSet(timeFolder, patternsHash);
            //    }

            //    Errors.AddRange(PatternListSheet.Errors);
            //}

            //#endregion

            return Errors.Any(x => x.ErrorLevel == EnumErrorLevel.Error);
        }

        public List<PatSetSubRow> GenPatSetSubRows(string patternFolder)
        {
            var patSetSubRows = new List<PatSetSubRow>();
            var usedPatterns = AiTestPlanSheets.SelectMany(x => x.Rows).SelectMany(x => x.Patterns)
                .Select(x => x.OriName.ToUpper()).Distinct().ToList();
            var patterns = new List<string>(Directory.GetFiles(patternFolder, "*.gz", SearchOption.AllDirectories))
                .ToList();
            foreach (var usedPattern in usedPatterns)
            {
                var patName = usedPattern + ".PAT.GZ";
                if (PatternListSheet.Rows.Exists(x =>
                        x.Pattern.Equals(usedPattern, StringComparison.CurrentCultureIgnoreCase)))
                {
                    var name = PatternListSheet.Rows.Find(x =>
                        x.Pattern.Equals(usedPattern, StringComparison.CurrentCultureIgnoreCase)).PatternVersion;
                    patName = name + ".PAT.GZ";
                }

                if (patterns.Exists(x => Path.GetFileName(x).Equals(patName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    var find = patterns.Find(x =>
                        Path.GetFileName(x).Equals(patName, StringComparison.CurrentCultureIgnoreCase));
                    var patternFolderName = Path.GetFileName(patternFolder);
                    var fileValue = find.Replace(patternFolder, @".\" + patternFolderName);
                    fileValue = fileValue.ToUpper().Replace(@".ATP.GZ", "").Replace(".PAT.GZ", "");

                    var moduleName = "";
                    var atpContent = "";
                    var patInforReader = new PatPatInforReader();
                    if (new PatInfoCmd().ConvertByArgs(find, ref atpContent, "-hdr -switches"))
                    {
                        var vmVectorName = patInforReader.VmVectorReader(atpContent.Split(new[] { '\n' }).ToList());
                        if (vmVectorName.Split(',').Count() == 2)
                            moduleName = vmVectorName.Split(',').Last();
                    }

                    if (string.IsNullOrEmpty(moduleName))
                    {
                        continue;
                    }

                    var patSetSubRow = new PatSetSubRow
                    {
                        Comment = "New for CZ",
                        PatternFileName = fileValue + ".PAT:" + moduleName.ToUpper()
                    };
                    patSetSubRows.Add(patSetSubRow);
                }
            }

            return patSetSubRows;
        }

        private void SetUniquePatSetName(List<AiTestPlanSheet> aiTestPlanSheets)
        {
            var patSetNames = new List<string>();
            foreach (var aiTestPlanSheet in aiTestPlanSheets)
                foreach (var row in aiTestPlanSheet.Rows)
                {
                    if (patSetNames.Contains(row.TestInstanceName, StringComparer.CurrentCultureIgnoreCase))
                    {
                        row.PatSetName = row.SheetName + "_Row" + row.RowNum + "_" + row.TestInstanceName;
                    }
                    else
                    {
                        row.PatSetName = row.TestInstanceName;
                        patSetNames.Add(row.TestInstanceName);
                    }
                }
        }

        public PatSetSheet GenPatSetAllSheet(string patternFolder, string patOrPatx)
        {
            //var patExt = string.Format("*{0}.gz", patOrPatx);
            patOrPatx = ".patx";
            var patExt = string.Format("*{0}*", "pat");

            var patterns = new List<string>(Directory.GetFiles(patternFolder, patExt, SearchOption.AllDirectories))
                .ToList();
            LogHelper.Debug($"Patterns Quantity: {patterns.Count}");

            if (patterns.Count < PatternListSheet.Rows.Count)
            {
                patExt = string.Format("*{0}", "pat");
                patterns = new List<string>(Directory.GetFiles(patternFolder, patExt, SearchOption.AllDirectories)).ToList();
            }

            LogHelper.Debug($"Patterns Quantity: {patterns.Count}");


            const string patSetsAllCz = "PatSets_All_" + Cz;
            var patSetSheet = new PatSetSheet(patSetsAllCz);
            var rows = AiTestPlanSheets.SelectMany(x => x.Rows)
                .Where(x => Regex.IsMatch(x.UseNotUse, "^use", RegexOptions.IgnoreCase)).ToList();
            var usedPatterns = rows.SelectMany(x => x.Patterns).Select(x => x.OriName)
                .Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList();
            foreach (var usedPattern in usedPatterns)
            {
                if (patSetSheet.Rows.Any(x => string.Equals(x.PatSetName, usedPattern, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }
                var patternDate = new PatternDate(usedPattern);
                var patGz = patternDate.OriName;

                LogHelper.Debug($"Pattern List Sheet Count: {PatternListSheet.Rows.Count}");
                LogHelper.Debug($"Pattern Name: {patGz}");

                if (PatternListSheet.Rows.Exists(x => x.Pattern.Equals(patternDate.OriName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    patGz = PatternListSheet.Rows
                        .Find(x => x.Pattern.Equals(patternDate.OriName, StringComparison.CurrentCultureIgnoreCase))
                        .PatternVersion;
                }

                var fileValue = "";

                var candidateExtSearch = new List<string> { ".pat", ".patx", ".pat.GZ", ".patx.GZ" };
                var patternFolderName = @"PATTERN";

                foreach (var candidateExt in candidateExtSearch)
                {
                    string searchedPatternName = $"{patGz}{candidateExt}";
                    string searchedPatternPath = patterns.FirstOrDefault(x => Path.GetFileName(x).Equals(searchedPatternName, StringComparison.CurrentCultureIgnoreCase));

                    if (!string.IsNullOrEmpty(searchedPatternPath))
                    {
                        fileValue = searchedPatternPath.Replace(patternFolder, @".\" + patternFolderName);
                        fileValue = fileValue.ToUpper().Replace(".GZ", "");
                        fileValue = fileValue + ":" + Path.GetFileNameWithoutExtension(searchedPatternPath.ToUpper());
                        break;
                    }
                }
                if (string.IsNullOrEmpty(fileValue))
                {
                    LogHelper.Info($"Pattern Not Found: {patGz}");

                }

                var patSet = new PatSet();
                patSet.PatSetName = patternDate.Name;
                var patSetRow = new PatSetRow();
                patSetRow.File = fileValue;
                patSetRow.Burst = "No";
                patSet.AddRow(patSetRow);
                patSetSheet.AddRow(patSet);
            }

            return patSetSheet;
        }

        public PatSetSheet GenPatSetSheet()
        {
            const string patSetsAllCz = "PatSets_" + Cz;
            var patSetSheet = new PatSetSheet(patSetsAllCz);
            foreach (var aiTestPlanSheet in AiTestPlanSheets)
                foreach (var row in aiTestPlanSheet.Rows)
                {
                    if (!Regex.IsMatch(row.UseNotUse, "^use", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }

                    var patSet = new PatSet();
                    patSet.PatSetName = row.PatSetName;
                    foreach (var pattern in row.Patterns)
                    {
                        var patSetRow = new PatSetRow();
                        patSetRow.File = pattern.Name;
                        patSetRow.Burst = "No";
                        patSet.AddRow(patSetRow);
                    }

                    patSetSheet.AddRow(patSet);
                }

            return patSetSheet;
        }

        public List<BasFile> GenBas(string execEnableWord, string testProgram)
        {
            var basFiles = new List<BasFile>();
            var sheetName = "VBT_LIB_PV.bas";
            var basFile = new BasFile(sheetName);
            var execEnableWords = execEnableWord.Split(',').ToList();
            //var igxlManagerMain = new IgxlManagerMain();
            //var totalEnableWords = igxlManagerMain.GetEnables(testProgram);
            List<string> totalEnableWords = Cautogen.AutoCZ.CharPostProcessor.GenBas.GetEnables(testProgram);
            var rows = new List<BasRow>
            {
                new (){ Text = "Attribute VB_Name = \"" + Path.GetFileNameWithoutExtension(sheetName) + "\""},
                new (){ Text = SetEnableWords(execEnableWords, totalEnableWords)},
                new (){ Text = PrintEnableWords(totalEnableWords)}
            };
            basFile.AddRows(rows);
            basFiles.Add(basFile);
            return basFiles;
        }

        private string SetEnableWords(List<string> execEnableWords, List<string> totalEnableWords)
        {
            var codeText = "Public Sub SetEnableWords()" + "\r\n";

            foreach (var enableWord in totalEnableWords)
            {
                var flag = execEnableWords.Exists(
                    x => x.Equals(enableWord, StringComparison.CurrentCultureIgnoreCase));
                codeText += string.Format("  tl_ExecSetEnableWord \"{0}\", {1}", enableWord, flag) + "\r\n";
            }

            codeText += "End Sub\r\n";
            return codeText;
        }

        private string PrintEnableWords(List<string> totalEnableWords)
        {
            var codeText = "Public Sub PrintEnableWords()" + "\r\n";
            foreach (var enableWord in totalEnableWords)
            {
                codeText += string.Format(
                    "  If (tl_ExecGetEnableWord(\"{0}\")) Then TheExec.Datalog.WriteComment \"{0}:\" + CStr(tl_ExecGetEnableWord(\"{0}\"))",
                    enableWord) + "\r\n";
            }

            codeText += "End Sub\r\n";
            return codeText;
        }

        public InstanceSheet GenInstSheet(VbtFunctionLib vbtFunctionLib, AcSpecSheet acSpecSheet, string doPayloadAfterSRMDSSCFlag, List<LevelSheet> levelSheets, PatSetSheet patSetCz, bool useNewCharLib = false)
        {
            string instCz = "Inst_" + Cz;
            string vbtName = VbtFunctionLib.FunctionalCharName;
            var instanceSheet = new InstanceSheet(instCz);
            var doPayloadAfterSRMDSSC = doPayloadAfterSRMDSSCFlag.Equals("True", StringComparison.OrdinalIgnoreCase);

            //Todo change generating flow for c# or VBT by isCSharp flag.

            if (_isCSharp)
            {
                vbtName = "FuncTestCharMain";
                foreach (var aiTestPlanSheet in AiTestPlanSheets)
                {
                    foreach (var row in aiTestPlanSheet.Rows)
                    {
                        var vbtFunction = vbtFunctionLib.GetFunctionByName(vbtName);

                        if (!Regex.IsMatch(row.UseNotUse, "^use", RegexOptions.IgnoreCase))
                        {
                            continue;
                        }


                        var instanceRow = new InstanceRow();
                        instanceRow.ColumnA = row.SetColumnA();

                        instanceRow.VbtName = vbtFunction.FullFunctionName;
                        instanceRow.VbtType = ".NET";
                        instanceRow.TestName = row.GetTestName();
                        instanceRow.ArgList = vbtFunction.Parameters;
                        instanceRow.DcCategory = row.DcCategory;
                        instanceRow.DcSelector = row.DcSelector;
                        instanceRow.TimeSets = GetTimeSet(row);
                        instanceRow.AcCategory = GetAc(row);
                        //instanceRow.AcCategory = !string.IsNullOrEmpty(instanceRow.AcCategory) ? instanceRow.AcCategory : acSpecSheet.GetAcByTimeSet(instanceRow.TimeSets);
                        instanceRow.AcSelector = "Typ";
                        instanceRow.PinLevels = GetPinLevels(row, levelSheets);
                        instanceRow.Args = vbtFunction.ArgList;
                        instanceSheet.AddRow(instanceRow);

                        SetCSharpParametersForNewTChar(vbtFunction, row, doPayloadAfterSRMDSSC, patSetCz);


                    }
                }


            }
            return instanceSheet;
        }
        private string GetPinLevels(AiTestPlanRow row, List<LevelSheet> levelSheets)
        {
            var block = row.GetBlock();
            var bincutLevels = string.Format("Levels_{0}_BinCut", block);
            var normalLevels = string.Format("Levels_{0}", block);
            if (row.DcLevelsMapping.Count != 1)
            {
                if (row.DcCategory.Equals("Bincut_X_X_X", StringComparison.OrdinalIgnoreCase))
                {
                    if (levelSheets.Any(x => x.Name.Equals(bincutLevels, StringComparison.OrdinalIgnoreCase)))
                    {
                        return bincutLevels;
                    }
                }
                return normalLevels;
            }
            else
            {
                if (row.DcLevelsMapping.First().Split(';').Length > 1)
                    return row.DcLevelsMapping.First().Split(';')[1];
            }
            return normalLevels;
        }
        private string GetTimeSet(AiTestPlanRow row)
        {
            //if (!string.IsNullOrEmpty(row.Timeset))
            //{
            //    return row.Timeset;
            //}

            if (row.TimesetMapping.Any())
            {
                return row.TimesetMapping.FirstOrDefault();
            }

            var timeSets = row.GetTimeSetsByPayloads(PatternListSheet);
            if (timeSets.Any())
            {
                return string.Join(",", timeSets.Distinct());
            }

            return "";
        }

        private string GetAc(AiTestPlanRow row)
        {
            if (row.AcCategoryMapping.Any())
            {
                return row.AcCategoryMapping.FirstOrDefault();
            }

            return "";
        }

        private void SetVbtParameters(VbtFunction vbtFunction, AiTestPlanRow aiTestPlanRow, bool doPayloadAfterSRMDSSC)
        {
            doPayloadAfterSRMDSSC = aiTestPlanRow.GetBlock() == "Scan" && doPayloadAfterSRMDSSC;
            var digSrcBitSizeArr = new string[15];
            var digSrcSegArr = new string[15];
            var digSrcPinArr = new string[15];
            var digSrcEQArr = new string[15];

            var writePayloads = new List<PatternDate>();

            for (var index = 0; index < aiTestPlanRow.Inits.Count; index++)
            {
                if (index > 9)
                    break;
                var init = aiTestPlanRow.Inits[index];
                var skipAfterSRMDSSC = doPayloadAfterSRMDSSC && init.Name.ToUpper().Contains("SRMDSSC");

                var patternName = "";
                if (index == 9 && !skipAfterSRMDSSC)
                {
                    patternName = string.Join(",", aiTestPlanRow.Inits.Skip(9).Select(x => x.Name));
                    foreach (var init10 in aiTestPlanRow.Inits.Skip(9))
                    {
                        init10.Index = "10";
                        init10.SubIndex = "INIT10";
                        if (!string.IsNullOrEmpty(init10.DigSrcBitSize) && !init10.SelsramDigSrc)
                        {
                            digSrcBitSizeArr[index] = init10.DigSrcBitSize;
                            digSrcSegArr[index] = init10.DigSrcEQ;
                            digSrcPinArr[index] = init10.DigSrcPin;
                            digSrcEQArr[index] = init10.DigSrcBits;
                        }
                    }
                }
                else
                {
                    patternName = init.Name;
                    init.Index = (index + 1).ToString();
                    init.SubIndex = "INIT" + (index + 1);
                    if (!init.SelsramDigSrc)
                    {
                        digSrcBitSizeArr[index] = init.DigSrcBitSize;
                        digSrcSegArr[index] = init.DigSrcEQ;
                        digSrcPinArr[index] = init.DigSrcPin;
                        digSrcEQArr[index] = init.DigSrcBits;
                    }
                }
                vbtFunction.SetParamValue("Init_Patt" + (index + 1), patternName);
                if (skipAfterSRMDSSC)
                {
                    writePayloads.AddRange(aiTestPlanRow.Inits.Skip(index + 1));
                    break;
                }
            }
            writePayloads.AddRange(aiTestPlanRow.Payloads);

            for (var index = 0; index < writePayloads.Count; index++)
            {
                if (index > 4)
                    break;
                var payload = writePayloads[index];
                var patternName = "";
                if (index == 4)
                {
                    patternName = string.Join(",", writePayloads.Skip(4).Select(x => x.Name));
                    foreach (var payload5 in writePayloads.Skip(4))
                    {
                        payload5.Index = "15";
                        payload5.SubIndex = "PL5";
                        if (!string.IsNullOrEmpty(payload5.DigSrcBitSize) && !payload5.SelsramDigSrc)
                        {
                            digSrcBitSizeArr[10 + index] = payload5.DigSrcBitSize;
                            digSrcSegArr[10 + index] = payload5.DigSrcEQ;
                            digSrcPinArr[10 + index] = payload5.DigSrcPin;
                            digSrcEQArr[10 + index] = payload5.DigSrcBits;
                        }
                    }
                }
                else
                {
                    patternName = payload.Name;
                    payload.Index = (index + 1 + 10).ToString();
                    payload.SubIndex = "PL" + (index + 1);
                    if (!payload.SelsramDigSrc)
                    {
                        digSrcBitSizeArr[10 + index] = payload.DigSrcBitSize;
                        digSrcSegArr[10 + index] = payload.DigSrcEQ;
                        digSrcPinArr[10 + index] = payload.DigSrcPin;
                        digSrcEQArr[10 + index] = payload.DigSrcBits;
                    }
                }

                vbtFunction.SetParamValue("PayLoad_Patt" + (index + 1), patternName);
            }
            vbtFunction.SetParamValue("DigSrc_BitSize", string.Join(",", digSrcBitSizeArr));
            vbtFunction.SetParamValue("DigSrc_Seg", string.Join(",", digSrcSegArr));
            vbtFunction.SetParamValue("DigSrc_DigSrcPin", string.Join(",", digSrcPinArr));
            vbtFunction.SetParamValue("digSrc_EQ", "'" + string.Join(",", digSrcEQArr));

            //force condition
            var forceCondition = new List<string>();
            foreach (var pin in aiTestPlanRow.Pins.Where(x => x.IsForce && x.Type == "Pin"))
            {
                forceCondition.Add(pin.Name + ":" + "V" + ":" + pin.ForceValue);
            }
            if (!string.IsNullOrEmpty(aiTestPlanRow.USL_LSL))
                forceCondition.Add(aiTestPlanRow.USL_LSL);

            vbtFunction.SetParamValue("Interpose_PrePat", string.Join(";", forceCondition));

            vbtFunction.SetParamValue("Power_Run_Scenario", !string.IsNullOrEmpty(aiTestPlanRow.PowerRunScenario) ? aiTestPlanRow.PowerRunScenario : "init_NV_pl_Sweep");
            vbtFunction.SetParamValue("Wait", GetWaitTime(""));
            vbtFunction.SetParamValue("BlockType", "");

            vbtFunction.SetParamValue("PatternTimeout", "30");
            vbtFunction.SetParamValue("SELSRAM_DSSC", aiTestPlanRow.SelsramDssc);
            vbtFunction.SetParamValue("Vbump", "True");
        }

        private void SetVbtParametersForNewTChar(VbtFunction vbtFunction, AiTestPlanRow aiTestPlanRow, bool doPayloadAfterSRMDSSC, PatSetSheet patSetCz)
        {
            SetupInitPlPatSetForNewTChar(vbtFunction, aiTestPlanRow, doPayloadAfterSRMDSSC, patSetCz);
            vbtFunction.SetParamValue("DigSrc_BitSize", string.Join(",", aiTestPlanRow.Patterns.Where(x => !string.IsNullOrEmpty(x.DigSrcBitSizeWithSubIndex)).Select(x => x.DigSrcBitSizeWithSubIndex)));
            vbtFunction.SetParamValue("DigSrc_Seg", string.Join(",", aiTestPlanRow.Patterns.Where(x => !string.IsNullOrEmpty(x.DigSrcSegWithSubIndex)).Select(x => x.DigSrcSegWithSubIndex)));
            vbtFunction.SetParamValue("DigSrc_DigSrcPin", string.Join(",", aiTestPlanRow.Patterns.Where(x => !string.IsNullOrEmpty(x.DigSrcPinWithSubIndex)).Select(x => x.DigSrcPinWithSubIndex)));
            vbtFunction.SetParamValue("digSrc_EQ", string.Join(",", aiTestPlanRow.Patterns.Where(x => !string.IsNullOrEmpty(x.DigSrcEQWithSubIndex)).Select(x => x.DigSrcEQWithSubIndex)));

            //force condition
            var forceCondition = new List<string>();
            foreach (var pin in aiTestPlanRow.Pins.Where(x => x.IsForce && x.Type == "Pin"))
            {
                forceCondition.Add(pin.Name + ":" + "V" + ":" + pin.ForceValue);
            }
            if (!string.IsNullOrEmpty(aiTestPlanRow.USL_LSL))
                forceCondition.Add(aiTestPlanRow.USL_LSL);

            vbtFunction.SetParamValue("Interpose_PrePat", string.Join(";", forceCondition));

            vbtFunction.SetParamValue("Power_Run_Scenario", !string.IsNullOrEmpty(aiTestPlanRow.PowerRunScenario) ? aiTestPlanRow.PowerRunScenario : "init_NV_pl_Sweep");
            //vbtFunction.SetParamValue("Wait", GetWaitTime(""));
            vbtFunction.SetParamValue("BlockType", "");

            vbtFunction.SetParamValue("PatternTimeout", "30");
            vbtFunction.SetParamValue("SELSRAM_DSSC", aiTestPlanRow.SelsramDssc);
            vbtFunction.SetParamValue("Vbump", "True");
        }

        private void SetCSharpParametersForNewTChar(VbtFunction vbtFunction, AiTestPlanRow aiTestPlanRow, bool doPayloadAfterSRMDSSC, PatSetSheet patSetCz)
        {
            SetupInitPlPatSetForNewTChar(vbtFunction, aiTestPlanRow, doPayloadAfterSRMDSSC, patSetCz, true);
            ConvertCSharpPatternDigSrc(vbtFunction, aiTestPlanRow, null);

            //force condition
            var forceCondition = new List<string>();
            foreach (var pin in aiTestPlanRow.Pins.Where(x => x.IsForce && x.Type == "Pin"))
            {
                forceCondition.Add(pin.Name + ":" + "V" + ":" + pin.ForceValue);
            }
            if (!string.IsNullOrEmpty(aiTestPlanRow.USL_LSL))
                forceCondition.Add(aiTestPlanRow.USL_LSL);

            vbtFunction.SetParamValue("interposePrePat", string.Join(";", forceCondition));
            vbtFunction.SetParamValue("powerRunScenario", !string.IsNullOrEmpty(aiTestPlanRow.PowerRunScenario) ? aiTestPlanRow.PowerRunScenario : "init_NV_pl_Sweep");

        }

        private void SetupInitPlPatSetForNewTChar(VbtFunction vbtFunction, AiTestPlanRow aiTestPlanRow, bool doPayloadAfterSRMDSSC, PatSetSheet patSetCz, bool isCSharp = false)
        {
            doPayloadAfterSRMDSSC = aiTestPlanRow.GetBlock() == "Scan" && doPayloadAfterSRMDSSC;
            var writeInits = new List<string>();
            var writePayloads = new List<string>();
            var adjustPayloads = new List<PatternDate>();
            var adjustPatIndexDict = new Dictionary<string, string>();
            var adjustIndex = 1;
            for (var index = 0; index < aiTestPlanRow.Inits.Count; index++)
            {
                var init = aiTestPlanRow.Inits[index];
                var skipAfterSRMDSSC = doPayloadAfterSRMDSSC && init.Name.ToUpper().Contains("SRMDSSC");
                writeInits.Add(init.Name);

                if (skipAfterSRMDSSC)
                {
                    adjustPayloads.AddRange(aiTestPlanRow.Inits.Skip(index + 1));
                    break;
                }
            }
            adjustPayloads.AddRange(aiTestPlanRow.Payloads);

            for (var index = 0; index < adjustPayloads.Count; index++)
            {
                var payload = adjustPayloads[index];
                writePayloads.Add(payload.Name);
                adjustPatIndexDict[payload.SubIndex] = "PL" + adjustIndex.ToString();
                adjustIndex++;
            }

            if (writeInits.Count > 0)
            {
                var initPatSet = new PatSet();
                initPatSet.PatSetName = aiTestPlanRow.GetTestName().ToUpper() + "_INIT";
                writeInits.ForEach(x => initPatSet.AddRow(new PatSetRow()
                {
                    PatternSet = initPatSet.PatSetName,
                    Burst = "NO",
                    File = x.ToUpper(),
                    Comment = string.Format("{0}:Row{1}", aiTestPlanRow.SheetName, aiTestPlanRow.RowNum)
                }));
                patSetCz.AddRow(initPatSet);
                if (isCSharp)
                    vbtFunction.SetParamValue("initPatset", initPatSet.PatSetName);

                else
                    vbtFunction.SetParamValue("INIT_Patset", initPatSet.PatSetName);
            }

            if (writePayloads.Count > 0)
            {
                var plPatSet = new PatSet();
                plPatSet.PatSetName = aiTestPlanRow.GetTestName().ToUpper() + "_PL";
                writePayloads.ForEach(x => plPatSet.AddRow(new PatSetRow()
                {
                    PatternSet = plPatSet.PatSetName,
                    Burst = "NO",
                    File = x.ToUpper(),
                    Comment = string.Format("{0}:Row{1}", aiTestPlanRow.SheetName, aiTestPlanRow.RowNum)
                }));
                patSetCz.AddRow(plPatSet);
                if (isCSharp)
                    vbtFunction.SetParamValue("plpatset", plPatSet.PatSetName);
                else
                    vbtFunction.SetParamValue("PL_Patset", plPatSet.PatSetName);

            }

            UpdatePatSubIndex(aiTestPlanRow, adjustPatIndexDict);
        }

        private void UpdatePatSubIndex(AiTestPlanRow aiTestPlanRow, Dictionary<string, string> adjustPatIndexDict)
        {
            foreach (var patternDate in aiTestPlanRow.Inits)
            {
                if (adjustPatIndexDict.ContainsKey(patternDate.SubIndex))
                    patternDate.SubIndex = adjustPatIndexDict[patternDate.SubIndex];
            }
            foreach (var patternDate in aiTestPlanRow.Payloads)
            {
                if (adjustPatIndexDict.ContainsKey(patternDate.SubIndex))
                    patternDate.SubIndex = adjustPatIndexDict[patternDate.SubIndex];
            }
        }

        private string GetWaitTime(string retention)
        {
            if (retention == "")
            {
                return ",,,,,,,,,,,,,,";
            }

            if (Regex.Matches(retention, ",").Count == 4)
            {
                return ",,,,,,,,,," + retention;
            }

            double time;
            if (double.TryParse(retention, out time))
            {
                return ",,,,,,,,,," + time + ",,,,";
            }

            return retention;
        }

        public SubFlowSheet GenFlowSheet(bool genTMPS = false)
        {
            const string flowCz = "Flow_" + Cz;
            var flowSheet = new SubFlowSheet(flowCz);

            if (!_isCSharp)
            {
                flowSheet.AddHeaderRow(Cz, "");
                flowSheet.AddRow(new FlowRow { Opcode = "Test", Parameter = "PrintEnableWords_Header" });
                flowSheet.AddRow(new FlowRow { Opcode = "Test", Parameter = "PrintEnableWords" });
                flowSheet.AddRow(new FlowRow { Opcode = "Test", Parameter = "PrintEnableWords_Footer" });
                flowSheet.AddRow(new FlowRow { Opcode = "nop", Enable = EnableFailLogDebug });
                flowSheet.AddRow(new FlowRow { Opcode = "nop", Enable = "Enable_DFTLHFC_Debug" });
                flowSheet.AddRow(new FlowRow { Opcode = "Test", Parameter = "Enable_Charz_mode" });
                flowSheet.AddRow(new FlowRow { Opcode = "flag-clear", Parameter = "F_Shmoo_Alarm" });

            }
            else
            {
                flowSheet.AddRow(new FlowRow { Opcode = "Print", Parameter = "\"====AutoAI====\"" });
                flowSheet.AddRow(new FlowRow { Opcode = "nop", Enable = EnableFailLogDebug });
                flowSheet.AddRow(new FlowRow { Opcode = "nop", Enable = "Enable_DFTLHFC_Debug" });
                flowSheet.AddRow(new FlowRow { Opcode = "Test", Parameter = "Enable_Charz_mode" });
                flowSheet.AddRow(new FlowRow { Opcode = "flag-clear", Parameter = "F_Shmoo_Alarm" });

            }


            var currentDataLoggingSettingType = EnumDataLoggingSettingType.Na;
            foreach (var aiTestPlanSheet in AiTestPlanSheets)
            {
                foreach (var row in aiTestPlanSheet.Rows)
                {
                    if (!Regex.IsMatch(row.UseNotUse, "^use", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }

                    var type = row.EnumDataLoggingSettingType;
                    if (type != currentDataLoggingSettingType)
                    {
                        if (type == EnumDataLoggingSettingType.Na)
                        {
                            flowSheet.AddRows(GenEnableOfTest());
                        }
                        else if (type == EnumDataLoggingSettingType.Fc)
                        {
                            flowSheet.AddRows(GenEnableOfFc());
                        }
                        else if (type == EnumDataLoggingSettingType.Dfc)
                        {
                            flowSheet.AddRows(GenEnableOfDfc());
                        }
                    }

                    var flowRowTest = new FlowRow();
                    flowRowTest.ColumnA = row.SetColumnA();
                    flowRowTest.Opcode = "Test";
                    flowRowTest.Parameter = row.GetTestName();
                    flowSheet.AddRow(flowRowTest);

                    if (row.EnumAiType == EnumAiType.Shmoo1D ||
                        row.EnumAiType == EnumAiType.Shmoo2D)
                    {
                        var flowRowShmoo = new FlowRow();
                        flowRowShmoo.ColumnA = row.SetColumnA();
                        flowRowShmoo.Opcode = "characterize";
                        flowRowShmoo.Parameter = row.Parameter;
                        flowSheet.AddRow(flowRowShmoo);
                        var binShmooAlarmFlowRow = new FlowRow();
                        binShmooAlarmFlowRow.ColumnA = row.SetColumnA();
                        binShmooAlarmFlowRow.Opcode = "BinTable";
                        binShmooAlarmFlowRow.Parameter = "Bin_Central_Gating_Rule_By_Flow";
                        flowSheet.AddRow(binShmooAlarmFlowRow);
                    }
                    if (genTMPS && row.GetTmpsFlowName() != "Flow_TMPS")
                    {
                        var flowRowTMPS = new FlowRow();
                        flowRowTMPS.ColumnA = row.SetColumnA();
                        flowRowTMPS.Opcode = "call";
                        flowRowTMPS.Parameter = row.GetTmpsFlowName();
                        flowSheet.AddRow(flowRowTMPS);
                    }
                    currentDataLoggingSettingType = type;
                }
            }
            if (!_isCSharp)
            {
                flowSheet.AddFooterRow(Cz, "");
            }
            flowSheet.AddReturnRow();
            return flowSheet;
        }

        private List<FlowRow> GenEnableOfTest()
        {
            var flowRows = new List<FlowRow>();
            flowRows.Add(new FlowRow { Opcode = "disable-flow-word", Parameter = DebugLvccVminBoundary });
            flowRows.Add(new FlowRow { Opcode = "disable-flow-word", Parameter = EnableDftlhfcDebug });
            flowRows.Add(new FlowRow { Opcode = "disable-flow-word", Parameter = EnableFailLogDebug });
            return flowRows;
        }

        private List<FlowRow> GenEnableOfFc()
        {
            var flowRows = new List<FlowRow>();
            flowRows.Add(new FlowRow { Opcode = "enable-flow-word", Parameter = DebugLvccVminBoundary });
            flowRows.Add(new FlowRow { Opcode = "disable-flow-word", Parameter = EnableDftlhfcDebug });
            flowRows.Add(new FlowRow { Opcode = "disable-flow-word", Parameter = EnableFailLogDebug });
            return flowRows;
        }

        private List<FlowRow> GenEnableOfDfc()
        {
            var flowRows = new List<FlowRow>();
            flowRows.Add(new FlowRow { Opcode = "enable-flow-word", Parameter = DebugLvccVminBoundary });
            flowRows.Add(new FlowRow { Opcode = "enable-flow-word", Parameter = EnableDftlhfcDebug });
            flowRows.Add(new FlowRow { Opcode = "enable-flow-word", Parameter = EnableFailLogDebug });
            return flowRows;
        }

        public CharSheet GenCharSheet(PinMapSheet currentPinMapSheet, AcSpecSheet currentAcSpecSheet, PortMapSheet currentPortMapSheet, string allSuspendDatalogFalse)
        {
            const string charCz = "Char_" + Cz;
            var charSheet = new CharSheet(charCz);
            foreach (var aiTestPlanSheet in AiTestPlanSheets)
            {
                foreach (var row in aiTestPlanSheet.Rows)
                {
                    if (!Regex.IsMatch(row.UseNotUse, "^use", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }

                    if (row.EnumAiType == EnumAiType.Shmoo1D)
                    {
                        var charSetup = new CharSetup();
                        charSetup.ColumnA = row.SetColumnA();
                        charSetup.SetupName = row.CharName;
                        charSetup.TestMethod = "Retest";
                        var pins = row.Pins.Where(x => x.IsSearch).ToList();
                        var pin = row.Pins.First(x => x.IsSearch);
                        if (!string.IsNullOrEmpty(row.Order))
                        {
                            var pinList = new List<Pin>();
                            foreach (var order in row.Order.Split(';'))
                            {
                                if (pins.Exists(x => x.Name.Equals(order, StringComparison.CurrentCultureIgnoreCase)))
                                {
                                    pinList.Add(
                                        pins.Find(x => x.Name.Equals(order, StringComparison.CurrentCultureIgnoreCase)));
                                }
                            }

                            if (pinList.Count == 1)
                            {
                                pin = pinList.First();
                            }
                        }
                        var testMethod = row.GetTestMethods().Any() ? row.GetTestMethods().First() : null;
                        charSetup.CharSteps.Add(CreateShmoo(row.CharName, pin,
                            testMethod, CharSetupConst.ModeXShmoo,
                            currentPinMapSheet, currentAcSpecSheet, currentPortMapSheet, row.FailCycleEachPoint));
                        foreach (var trackingPin in pins.Where(x => !x.Name.Equals(pin.Name, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            charSetup.CharSteps.Add(CreateShmooTracking(row.CharName, trackingPin,
                                CharSetupConst.ModeXShmoo,
                                currentPinMapSheet, currentAcSpecSheet, currentPortMapSheet));
                        }
                        charSheet.AddRow(charSetup);
                    }
                    else if (row.EnumAiType == EnumAiType.Shmoo2D)
                    {
                        var charSetup = new CharSetup();
                        charSetup.ColumnA = row.SetColumnA();
                        charSetup.SetupName = row.CharName;
                        charSetup.TestMethod = "Retest";
                        var pins = row.Pins.Where(x => x.IsSearch).ToList();
                        if (!string.IsNullOrEmpty(row.Order))
                        {
                            var pinList = new List<Pin>();
                            foreach (var pin in row.Order.Split(';'))
                            {
                                if (pins.Exists(x => x.Name.Replace("_", "").Equals(pin, StringComparison.CurrentCultureIgnoreCase)))
                                {
                                    pinList.Add(
                                        pins.Find(x => x.Name.Replace("_", "").Equals(pin, StringComparison.CurrentCultureIgnoreCase)));
                                }
                            }

                            if (pinList.Count == 2)
                            {
                                pins = pinList;
                            }
                        }

                        var testMethods = row.GetTestMethods();
                        TestMethod testMethodsX = null;
                        TestMethod testMethodsY = null;
                        if (testMethods != null)
                        {
                            testMethodsX = testMethods.First();
                            testMethodsY = testMethods.First();
                            if (testMethods.Count == 2)
                            {
                                testMethodsY = testMethods[1];
                            }
                        }

                        var pinX = pins.ElementAt(0);
                        var pinY = pins.ElementAt(0);
                        if (pins.Count == 2)
                        {
                            pinY = pins.ElementAt(1);
                        }

                        charSetup.CharSteps.Add(CreateShmoo(row.CharName, pinX,
                            testMethodsX, CharSetupConst.ModeXShmoo,
                            currentPinMapSheet, currentAcSpecSheet, currentPortMapSheet, row.FailCycleEachPoint));

                        foreach (var trackingPin in row.Pins.Where(x => x.IsSearch
                                                                    && !x.Name.Equals(pinX.Name, StringComparison.CurrentCultureIgnoreCase)
                                                                    && !x.Name.Equals(pinY.Name, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            charSetup.CharSteps.Add(CreateShmooTracking(row.CharName, trackingPin,
                                CharSetupConst.ModeXShmoo,
                                currentPinMapSheet, currentAcSpecSheet, currentPortMapSheet));
                        }

                        charSetup.CharSteps.Add(CreateShmoo(row.CharName, pinY,
                            testMethodsY, CharSetupConst.ModeYShmoo,
                            currentPinMapSheet, currentAcSpecSheet, currentPortMapSheet, row.FailCycleEachPoint));
                        charSheet.AddRow(charSetup);
                    }
                }
            }
            if (allSuspendDatalogFalse.Equals("True", StringComparison.OrdinalIgnoreCase))
            {
                //charSheet.CharSetups.ForEach(x => x.CharSteps.ForEach(y => y.SuspendDataLog = "FALSE"));
            }
            return charSheet;
        }

        private CharStep CreateShmoo(string charName, Pin pin, TestMethod testMethod, string method,
            PinMapSheet currentPinMapSheet, AcSpecSheet currentAcSpecSheet, PortMapSheet currentPortMapSheet, string suspendDatalog)
        {
            var stepName = pin.Name + "_" + method.Replace(" ", "_");
            var setup = new CharStep(charName, stepName);
            setup.Mode = method;
            var forceType = GetForceType(pin.Name, currentPinMapSheet, currentAcSpecSheet);
            setup.ParameterType = GetParameterTypeGlobalSpec(forceType);
            setup.ParameterName = setup.ParameterType == CharSetupConst.ParameterTypeAcSpec ? pin.Name : pin.ShmooName;

            if (setup.ParameterType == CharSetupConst.ParameterTypeAcSpec && pin.Name.IndexOf("Freq_VAR", StringComparison.CurrentCultureIgnoreCase) != -1)
            {
                var portName = Regex.Replace(pin.Name, @"Freq_VAR", @"Port", RegexOptions.IgnoreCase);
                if (currentPortMapSheet.Rows.Any(x => string.Equals(x.PortName, portName, StringComparison.OrdinalIgnoreCase)))
                {
                    setup.PrePoint = "CoreTestLibrary.Char.FunctionalTestCharMain.FreerunclkSetXY";
                    setup.PrePointArguments = string.Format("{0},{1},{2}", setup.Mode[0], portName, pin.Name);
                    setup.PostPoint = "CoreTestLibrary.Char.FunctionalTestCharMain.FreerunclkStop";
                    setup.PostPointArguments = portName;
                }
            }

            setup.RangeCalcField = CharSetupConst.RangeCalcFieldSteps;
            if (forceType == EnumForceType.Frequency)
            {
                string start;
                pin.Start.TryConvertToFreq(out start);
                setup.RangeFrom = start;
                string stop;
                pin.Stop.TryConvertToFreq(out stop);
                setup.RangeTo = stop;
                string step;
                pin.Step.TryConvertToFreq(out step);
                setup.RangeStepSize = step;
            }
            else
            {
                string start;
                pin.Start.TryConvertToVolt(out start);
                setup.RangeFrom = start;
                string stop;
                pin.Stop.TryConvertToVolt(out stop);
                setup.RangeTo = stop;
                string step;
                pin.Step.TryConvertToVolt(out step);
                setup.RangeStepSize = step;
            }

            setup.AlgorithmName = CharSetupConst.AlgorithmNameLinear;
            if (testMethod != null)
            {
                if (testMethod.Name.Equals(CharSetupConst.AlgorithmNameJump, StringComparison.CurrentCultureIgnoreCase))
                {
                    setup.AlgorithmName = testMethod.Name;
                    if (!string.IsNullOrEmpty(testMethod.Arguments))
                    {
                        setup.AlgorithmArguments = testMethod.Arguments;
                    }
                }
            }

            setup.PostSetupArguments = "CorePower," + pin.Name;
            //setup.PostStepFunction = CharStepConst.PostStepFunctionPrintShmooInfo;
            setup.PostSetup = _isCSharp ? "CoreTestLibrary.Char.FunctionalTestCharMain.PrintShmooInfoMain" : ConstData.InterPostPostStep;

            if (setup.ParameterType != CharSetupConst.ParameterTypeAcSpec)
            {
                setup.ApplyToPins = pin.Name;
                setup.ApplyToPinExecMode = "Simultaneous";
            }

            setup.AxisExecutionOrder = "X-Y[-Z]";
            setup.OutputFormat = "Enhanced";
            setup.OutputDestinationsTextFile = "Disable";
            setup.OutputDestinationsSheet = "Disable";
            setup.OutputSuspendDatalog = suspendDatalog.ToUpper() == "TRUE" ? "FALSE" : "TRUE";
            setup.OutputDestinationsDatalog = "Enable";
            setup.OutputDestinationsImmediateWin = "Disable";
            setup.OutputDestinationsOutputWin = "Disable";
            return setup;
        }

        private CharStep CreateShmooTracking(string charName, Pin pin, string method,
            PinMapSheet currentPinMapSheet, AcSpecSheet currentAcSpecSheet, PortMapSheet currentPortMapSheet)
        {
            var stepName = pin.Name + "_" + method.Replace(" ", "_") + "_Tracking";
            var setup = new CharStep(charName, stepName);
            setup.Mode = method;
            var forceType = GetForceType(pin.Name, currentPinMapSheet, currentAcSpecSheet);
            setup.ParameterType = GetParameterTypeGlobalSpec(forceType);
            setup.ParameterName = setup.ParameterType == CharSetupConst.ParameterTypeAcSpec ? pin.Name : pin.ShmooName;

            if (setup.ParameterType == CharSetupConst.ParameterTypeAcSpec && pin.Name.IndexOf("Freq_VAR") != -1)
            {
                var portName = Regex.Replace(pin.Name, @"Freq_VAR", @"Port", RegexOptions.IgnoreCase);
                if (currentPortMapSheet.Rows.Any(x => string.Equals(x.PortName, portName, StringComparison.OrdinalIgnoreCase)))
                {
                    setup.PrePoint = "freerunclk_set_XY";
                    setup.PrePointArguments = string.Format("{0},{1},{2}", setup.Mode[0], portName, pin.Name);
                    setup.PostPoint = "freerunclk_stop";
                    setup.PostPointArguments = portName;
                }
            }

            if (forceType == EnumForceType.Frequency)
            {
                string start;
                pin.Start.TryConvertToFreq(out start);
                setup.RangeFrom = start;
                string stop;
                pin.Stop.TryConvertToFreq(out stop);
                setup.RangeTo = stop;
            }
            else
            {
                string start;
                pin.Start.TryConvertToVolt(out start);
                setup.RangeFrom = start;
                string stop;
                pin.Stop.TryConvertToVolt(out stop);
                setup.RangeTo = stop;
            }
            if (setup.ParameterType != CharSetupConst.ParameterTypeAcSpec)
            {
                setup.ApplyToPins = pin.Name;
                setup.ApplyToPinExecMode = "Simultaneous";
            }
            return setup;
        }

        private string GetParameterTypeGlobalSpec(EnumForceType forceType)
        {
            if (forceType == EnumForceType.Voltage)
            {
                return CharSetupConst.ParameterTypeGlobalSpec;
            }

            if (forceType == EnumForceType.Frequency)
            {
                return CharSetupConst.ParameterTypeAcSpec;
            }

            return CharSetupConst.ParameterTypeGlobalSpec;
        }

        private EnumForceType GetForceType(string name, PinMapSheet currentPinMapSheet, AcSpecSheet currentAcSpecSheet)
        {
            if (currentPinMapSheet != null && currentPinMapSheet.IsPinExist(name))
            {
                return EnumForceType.Voltage;
            }

            if (currentAcSpecSheet != null && currentAcSpecSheet.IsSymbolExist(name))
            {
                return EnumForceType.Frequency;
            }

            return EnumForceType.Voltage;
        }

        private static void ConvertCSharpPatternDigSrc(VbtFunction vbtFunction, AiTestPlanRow planItem, Dictionary<string, string> ArgPatternIndexConvertedDsscDict)
        {
            var patternList = planItem.Patterns.Where(x => !string.IsNullOrEmpty(x.DigSrcPin)).ToList();
            if (!patternList.Any())
            {
                return;
            }

            string sourcePin = patternList.Select(x => x.DigSrcPin).FirstOrDefault();
            string[] digSrcEquation = new string[planItem.Patterns.Count];
            List<string> digSrcAssignment = [];

            int sgmtIdx = 0;

            foreach (var pattern in patternList)
            {
                if (!int.TryParse(pattern.Index, out int patternIndex))
                {
                    continue;
                }

                List<string> equations = [];

                var sourceSeqList = pattern.DigSrcSeg.Split(';');
                foreach (var sourceSeq in sourceSeqList)
                {
                    equations.Add($"sgmt{sgmtIdx}");

                    var sgmt = sourceSeq.Split('=')[1];

                    if (Regex.IsMatch(sgmt, @"SELSRAM", RegexOptions.IgnoreCase))
                    {
                        var selsrmDSSC = planItem.SelsramDssc;
                        if (!string.IsNullOrEmpty(selsrmDSSC))
                        {
                            var selsrmBits = Regex.Replace(selsrmDSSC, @"selsrm", "", RegexOptions.IgnoreCase);
                            var selsrmBitsReplace = Regex.Replace(selsrmBits, @"s", "", RegexOptions.IgnoreCase);

                            if (!string.IsNullOrEmpty(selsrmBitsReplace))
                            {
                                sgmt = $"sgmt{sgmtIdx}=SELSRAM({string.Join(",", selsrmBits.ToArray())})";
                            }
                            else
                            {
                                sgmt = $"sgmt{sgmtIdx}=SELSRAM()";
                            }
                            digSrcAssignment.Add(sgmt);
                        }
                    }
                    else
                    {
                        digSrcAssignment.Add($"sgmt{sgmtIdx}={sgmt}");
                    }

                    sgmtIdx++;
                }

                digSrcEquation[patternIndex - 1] = string.Join('+', equations);
            }

            vbtFunction.SetParamValue("digSrcPin", sourcePin);
            vbtFunction.SetParamValue("digSrcEquation", string.Join('|', digSrcEquation));
            vbtFunction.SetParamValue("digSrcAssignment", string.Join(';', digSrcAssignment));
        }

        private static string ConvertBinStr(string binStr, HardIpReference patInfo, bool IsReverseData)
        {
            if (Regex.IsMatch(binStr, @"^\d+$"))
                return binStr;

            var bitStrArray = patInfo.SendBitStr.Split('+');
            //Comment for EMA remapping
            //if (!Regex.IsMatch(binStr, @"sgmt\d+"))
            //{
            //    var EmaTarget =
            //        LocalSpecs.EmaMappingItems.FirstOrDefault(
            //            p => p.Pattern.Equals(patInfo.Payload, StringComparison.OrdinalIgnoreCase));
            //    if (EmaTarget != null)
            //    {
            //        var data = EmaTarget.GetCaseData(binStr);
            //        binStr = data;
            //    }
            //}
            var defaultBinStr = Regex.Match(binStr, @"sgmtdef(?<value>\d+)", RegexOptions.IgnoreCase).Groups["value"].ToString();
            var sgmtSets = Regex.Split(binStr, @"(sgmt\d+)", RegexOptions.IgnoreCase);

            var dicBin = new Dictionary<string, string>();

            //sgmtName: sgmt[0-9], sgmtStr: srcData
            for (var i = 0; i < sgmtSets.Length; i++)
            {
                var sgmtName = Regex.Match(sgmtSets[i], @"sgmt\d+", RegexOptions.IgnoreCase).ToString().ToLower();

                if (sgmtName == "" || i + 1 >= sgmtSets.Length)
                    continue;

                var sgmtStr = "";

                if (Regex.IsMatch(sgmtSets[i + 1], @"f[01S]+", RegexOptions.IgnoreCase))
                    sgmtStr = sgmtSets[i + 1].Split('f')[1];

                else if (Regex.IsMatch(sgmtSets[i + 1], @"g[0-9a-fA-F]+", RegexOptions.IgnoreCase))
                {
                    var hexArray =
                        Regex.Match(sgmtSets[i + 1], @"g(?<data>[0-9a-fA-F]+)").Groups["data"].ToString().ToCharArray();

                    sgmtStr = (
                        from hex in hexArray
                        select Convert.ToInt32(hex.ToString(CultureInfo.InvariantCulture), 16)
                            into value
                        select Convert.ToString(value, 2)
                                into hexSrc
                        select hexSrc.PadLeft(4, '0')
                        ).Aggregate(sgmtStr, (current, hexSrc) => current + hexSrc);
                }
                dicBin.Add(sgmtName, sgmtStr);
            }

            int defaultValue;
            if (!Int32.TryParse(defaultBinStr, out defaultValue))
                defaultValue = 0;

            //ConstData.DefaultSigsrcValue = "sgmt_default=" + defaultValue;
            var newBinStr = "";
            if (!string.IsNullOrEmpty(patInfo.SendBitStr))
            {
                foreach (var sgmt in bitStrArray)
                {
                    try
                    {
                        var sgmtName = sgmt.Split('_')[0].ToLower();
                        var sgmtLength = sgmt.Split('_')[1];
                        var bitCount = Convert.ToInt32(sgmtLength);

                        if (dicBin.ContainsKey(sgmtName))
                        {
                            var data = "";
                            if (bitCount > dicBin[sgmtName].Length)
                                data = dicBin[sgmtName].PadLeft(bitCount, '0');

                            else if (bitCount < dicBin[sgmtName].Length)
                                data = dicBin[sgmtName].Substring(dicBin[sgmtName].Length - bitCount, bitCount);

                            else
                                data = dicBin[sgmtName];

                            if (IsReverseData)
                            {
                                var chararray = data.ToCharArray();
                                Array.Reverse(chararray);
                                data = new string(chararray);
                            }
                            newBinStr += data;
                        }
                        else
                        {
                            var str = Convert.ToString(defaultValue, 2);
                            var newStr = str.PadLeft(bitCount, '0');
                            newBinStr += newStr;
                        }
                    }
                    catch (Exception e)
                    {
                        ;
                    }
                }
            }
            else
            {
                var dataFromCharPlan = IsReverseData
                    ? string.Join("", dicBin.Values.Select(p => new string(p.Reverse().ToArray())))
                    : string.Join("", dicBin.Values.Select(p => p));
                newBinStr = dataFromCharPlan;
            }
            return newBinStr;
        }
    }
}
