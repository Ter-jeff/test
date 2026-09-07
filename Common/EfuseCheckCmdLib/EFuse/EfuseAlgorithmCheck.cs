using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.Extension;

using EfuseCheckCmdLib.AlgorithmCheck;
using EfuseCheckCmdLib.EFuse.EFuseApp;

using IgxlLib;

using OfficeOpenXml;

using TestPlanLib.Efuse;
using TestPlanLib.Settings;

namespace EfuseCheckCmdLib.EFuse
{
    public partial class EfuseAlgorithmCheck
    {
        [GeneratedRegex("CFG_Condition_Table", RegexOptions.IgnoreCase)]
        private static partial Regex CfgConditionTableFileRegex();

        [GeneratedRegex("Config_table", RegexOptions.IgnoreCase)]
        private static partial Regex ConfigTableFileRegex();

        [GeneratedRegex("EFUSE_BitDef_Table", RegexOptions.IgnoreCase)]
        private static partial Regex BitDefTableFileRegex();

        public static string CurrentPrj { get; private set; } = null!;
        public static Dictionary<string, int> JobFlow { get; private set; } = null!;
        public static bool IsBdfCheck { get; private set; }
        public static string TestplanFile { get; private set; } = null!;
        public static string WaferIdFile { get; private set; } = null!;
        public string BitDefFile;
        public string InputFolder;
        public string CfgTableFile;
        public string TestProgramFile;
        public List<string> StdfFile;
        public string OutputPath;
        public bool Bin1Only = false;
        public Action<string, string> AppendRichText;

        public EfuseAlgorithmCheck(Action<string, string> richTextBoxAppend, string inputFolder,
            string bitDefFile, string cfgTableFile, string testplanFile, string tpfile, string waferIdFile, List<string> stdfFile, bool bdfCheck, string outputPath = "")
        {
            AppendRichText = richTextBoxAppend;
            BitDefFile = bitDefFile;
            InputFolder = inputFolder;
            CfgTableFile = cfgTableFile;
            StdfFile = stdfFile;
            TestProgramFile = tpfile;
            IsBdfCheck = bdfCheck;
            TestplanFile = testplanFile;
            WaferIdFile = waferIdFile;
            OutputPath = outputPath;
        }

        public void WorkFlow(bool isMergeResult)
        {
            Dictionary<string, string> projects = GetProjectFromSettings();
            XParseDatalog.BitDefRef = null;
            if (!string.IsNullOrEmpty(WaferIdFile))
            {
                AppendRichText("Start Generate Prober Hex Code ...", "Blue");
                EFuseMainScript.CheckProberHexCode(WaferIdFile);
                AppendRichText("Generate Prober Hex Code Done...", "Blue");
                return;
            }
            if (!string.IsNullOrEmpty(TestProgramFile))
            {
                string outputProg = Path.Combine(Path.GetDirectoryName(TestProgramFile)!, "exportProg");
                var exportMain = new IgxlManager();
                IgxlManager.ExportWorkBook(TestProgramFile, outputProg);
                if (string.IsNullOrEmpty(BitDefFile))
                {
                    string? bitdef = Directory.GetFiles(outputProg).FirstOrDefault(p => Path.GetFileNameWithoutExtension(p).EqualsIgnoreCase("EFUSE_BitDef_Table"));
                    if (bitdef != null)
                    {
                        BitDefFile = bitdef;
                    }
                }
                if (string.IsNullOrEmpty(CfgTableFile))
                {
                    string? cfg = Directory.GetFiles(outputProg).FirstOrDefault(p => CfgConditionTableFileRegex().IsMatch(Path.GetFileNameWithoutExtension(p)));
                    if (cfg != null)
                    {
                        CfgTableFile = cfg;
                    }
                }
            }
            else if (!string.IsNullOrEmpty(TestplanFile))
            {
                AppendRichText($"Load Testplan File{TestplanFile}: ", "Blue");
                var inputExcel = new ExcelPackage(new FileInfo(TestplanFile));
                if (inputExcel != null)
                {
                    AppendRichText("Start Generate BDF & Config ...", "Blue");
                    EpWorkbook.TestPlanWorkbook = inputExcel.Workbook;
                    new EFuseMainScript().NonIgxlSheetProcess(true);

                    var tpFileInfo = new FileInfo(TestplanFile);
                    CfgTableFile = Directory.GetFiles(tpFileInfo.DirectoryName!).FirstOrDefault(p => ConfigTableFileRegex().IsMatch(Path.GetFileNameWithoutExtension(p)))!;

                    BitDefFile = Directory.GetFiles(tpFileInfo.DirectoryName!).FirstOrDefault(p => BitDefTableFileRegex().IsMatch(Path.GetFileNameWithoutExtension(p)))!;

                    AppendRichText($"Generate BDF Done {BitDefFile}: ", "Green");
                    AppendRichText($"Generate Config Done {CfgTableFile}: ", "Green");

                    if (ErrorReportManager.GetErrorList().Count > 0)
                    {
                        string extension = tpFileInfo.Extension;
                        string errorReportFile = TestplanFile.Replace(extension, "_Real" + VersionControl.Timestamp + extension);
                        File.Copy(TestplanFile, errorReportFile);
                        using (var pkg = new ExcelPackage(new FileInfo(errorReportFile)))
                        {
                            ErrorReportManager.GenErrorReport(pkg, "ErrorReport");
                            pkg.Save();
                        }
                        AppendRichText($"Generate Error Report {errorReportFile}: ", "Blue");
                    }
                }
            }

            CurrentPrj = ParseProjectCode(projects, InputFolder);

            string basicConfig =
                Path.Combine(Directory.GetCurrentDirectory(), "Settings", "Basic", $"Basic_Configure_{CurrentPrj}.xlsx");
            if (!File.Exists(basicConfig))
            {
                basicConfig =
                    Path.Combine(Directory.GetCurrentDirectory(), "Settings", "Basic", "Basic_Configure_default.xlsx");
            }
            AppendRichText($"Load Setting File{basicConfig}: ", "Blue");
            JobFlow = AccessJobFlow(basicConfig);
            AccessHarverstFieldName(basicConfig);

            string projectFilepath = Path.Combine(Directory.GetCurrentDirectory(), "Settings", $"EfuseScriptDescription_{CurrentPrj}.xml");
            if (!File.Exists(projectFilepath))
            {
                //Copy from default
                File.Copy(Path.Combine(Directory.GetCurrentDirectory(), "Settings", $"EfuseScriptDescription_default.xml"), projectFilepath);
            }
            AppendRichText($"Load EfuseScript Config File{projectFilepath}: ", "Blue");
            EfuseScriptConfig config = LoadEfuseConfigFile(projectFilepath);
            // read BDF File and push to datalog property, temp code
            AppendRichText("Load BitDefinition File...", "Blue");
            string stdFileFirst = "";
            if (StdfFile.Count > 0)
            {
                stdFileFirst = StdfFile[0];
            }

            var efuseBitDef = new LoaderEfuseBitDef(BitDefFile!, stdFileFirst);
            efuseBitDef.Parse(config);
            if (!string.IsNullOrEmpty(stdFileFirst))
            {
                AppendRichText("Load Stdf File...", "Blue");
                XParseDatalog.StdfReader = new EFuseStdfReader1(stdFileFirst, [.. efuseBitDef.BitDefTable.HipList.Keys]);
                XParseDatalog.StdfReader.WorkFlow();
            }
            XParseDatalog.BitDefRef = efuseBitDef;
            if (InputFolder.Length != 0)
            {

                //public xParseDatalog(string inDir, TextBox textBox, string efuseBitDefPath)

                var exeObj = new XParseDatalog(config, InputFolder, StdfFile, OutputPath) { Bin1Only = Bin1Only };
                if (string.IsNullOrEmpty(CfgTableFile) && Path.GetExtension(BitDefFile).EqualsIgnoreCase(".xlsx"))
                {
                    CfgTableFile = BitDefFile;
                }

                if (!string.IsNullOrEmpty(CfgTableFile))
                {
                    EfuseCfgTableReader1.WorkFlow(CfgTableFile);
                }

                AppendRichText("Parsing datalog...", "Blue");
                exeObj.Execute(isMergeResult, config, AppendRichText);
            }

        }

        //Access Project from Settings.ini
        private static Dictionary<string, string> GetProjectFromSettings()
        {
            var projects = new Dictionary<string, string>();
            string iniPath = Path.Combine(LocalSpecs.SettingFolder, "Settings", "Settings.ini");
            if (!File.Exists(iniPath))
            {
                Console.WriteLine("Settings.ini not found");
                return [];
            }
            var myIni = new IniFile(iniPath);
            //project code
            string[] projectCodes = myIni.Read("ProjectCode").Split('|');
            foreach (string sProject in projectCodes)
            {
                string project = myIni.Read(sProject);
                if (project.Length > 0 && !projects.ContainsKey(sProject[..3]))
                {
                    projects.Add(sProject[..3], project);
                }
            }
            //product code
            string[] productCodes = myIni.Read("ProductCode").Split('|');
            foreach (string sProduct in productCodes)
            {
                string project = myIni.Read(sProduct);
                if (project.Length > 0 && !projects.ContainsKey(sProduct))
                {
                    projects.Add(sProduct, project);
                }
            }
            return projects;
        }

        //Access Project from datalog
        private static string ParseProjectCode(Dictionary<string, string> projects, string dir)
        {
            string[] datalogs = Directory.GetFiles(dir);
            string key = "";
            foreach (string datalog in datalogs)
            {
                using var sr = new StreamReader(datalog);
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine()!;
                    if (string.IsNullOrEmpty(line))
                    {
                        continue;
                    }

                    if (line.ContainsIgnoreCase("prog name:"))
                    {
                        string name = line.Split(':')[1];
                        key = name.Split('_').Length != 0 ? name.Split('_')[0].Trim() : "";
                        break;
                    }
                }
                sr.Close();
            }
            foreach (string prjkey in projects.Keys)
            {
                if (Regex.IsMatch(key, prjkey, RegexOptions.IgnoreCase))
                {
                    key = prjkey;
                    break;
                }
            }
            return key.Length != 0 && projects.TryGetValue(key, out string? project) ? project : "default";
        }

        //Access JobFlow From Basic_Configure_Project file
        private static Dictionary<string, int> AccessJobFlow(string path)
        {
            var jobflow = new Dictionary<string, int>();
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using ExcelPackage ep = new ExcelPackage(fs);
                ExcelWorksheet worksheet = ep.Workbook.Worksheets["JobMapping"];
                JobMapSheet jobMapSheet = SettingStatic.JobMapSheet;
                Dictionary<string, List<string>> jobs = jobMapSheet.JobMapDictionary;
                int i = 0;
                foreach (string job in jobs.Keys)
                {
                    foreach (string substage in jobs[job])
                    {
                        jobflow.Add(substage, i);
                        i++;
                    }
                }
            }

            return jobflow;
        }

        private static void AccessHarverstFieldName(string path)
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var ep = new ExcelPackage(fs);
            ExcelWorksheet worksheet = ep.Workbook.Worksheets["HarvestFieldForEfuseCheck"];
            if (worksheet == null)
            {
                return;
            }

            var reader = new HarvestFieldReader();
            reader.ReadFlow(worksheet);
        }

        public static EfuseScriptConfig LoadEfuseConfigFile(string filepath)
        {
            return EfuseScriptConfigReader.LoadXml(filepath);
        }
    }
}
