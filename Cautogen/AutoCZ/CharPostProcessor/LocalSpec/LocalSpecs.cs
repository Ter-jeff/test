using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.AutoCZ.CharPostProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure.IgxlData.Prog;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure.ShmooData;
using Cautogen.AutoCZ.CharPostProcessor.Utility.TestNumManager;
using Cautogen.common.IgxlProgramLib.IgxlProgramParser;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;
using Cautogen.Utility;

using CommonLib.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPostProcessor.LocalSpec
{
    public class LocalSpecs
    {
        public static string TimeSetFolder { get; set; }
        public static string PatternFolder { get; set; }
        public static string PatListFile { get; set; }

        public static string Project;
        public static IGXL IgxlConfig;
        public static bool GenAssignSiteVar;
        public static bool ProgramUpdateOnly;
        public static string OutputFolder;
        public static List<string> AllModuleSheets;
        public static List<string> AllCommonSheets;
        public static List<CharPlanSheet> CharPlanSheets { get; set; }
        public static List<string> PatternsInCharPlan;
        public static Dictionary<string, Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure.HardIpReference> PatInfoList;
        public static Dictionary<string, PatternData> PatternDatas;
        public static IgxlProgram TestProgram;
        public static Dictionary<string, string> FileStructure;
        public static List<string> PreserveFileList;
        public static List<string> FileLoadList;
        public static ProgInfo ProgInfo;
        public static List<ShmooSetup> AllShmooSetups;
        public static InputParam InputParam;
        public static Dictionary<string, SubFlowSheet> HipTempFlowDic;
        public static Dictionary<string, List<string>> ProgFlowDic; // {charPlanSheetName: [prodFlowSheetName related to the charPlanSheet]}
        public static Dictionary<string, string> PatternDic;
        public static Dictionary<string, FlowRow> AllFlowStepsDic;
        public static PostMessageWriter MessageWriter;
        public static ExcelPackage PostReportWriter;
        public static Dictionary<string, Dictionary<string, List<FlowRow>>> UseLimitDict;
        public static PatSetSheet PatSetAll;
        public static Dictionary<string, List<string>> HtolAndTtr;
        public static bool IsDirectory;
        public static double ExportVersion;
        public static Dictionary<string, CompileITem> CompileITemDic;
        public static List<EmaMappingItem> EmaMappingItems;
        public static SubFlowSheet MainFlowSheets;
        public static ExcelWorksheet DFCSheet;
        public static AdaptiveCooling AdaptiveCooling;
        public static string CurrentJob;
        public static IProgress<string> Progress;
        public static ExcelWorksheet OptionalTimesettings;
        public static DigSrcReg DigSrcReg;
        public static List<IIgxlSheet> GenSheets;
        public static List<string> GenOthers;

        public static string TimeStamp
        {
            get { return TimeContext.Now.ToString("yyyyMMdd_HHmmss"); }
        }

        public static void SetAllKeptSheets()
        {
            AllModuleSheets = new List<string>();
            string moduleFolder = Path.Combine(OutputFolder, ConstData.ModuleFolder);
            foreach (string fileName in Directory.GetDirectories(moduleFolder)
                .SelectMany(folder => Directory.GetFiles(folder)
                .Select(Path.GetFileNameWithoutExtension)
                .Where(fileName => !AllModuleSheets.Contains(fileName))))
            {
                AllModuleSheets.Add(fileName);
            }

            AllCommonSheets = new List<string>();
            string commonFolder = Path.Combine(OutputFolder, ConstData.CommonFolder);
            if (!Directory.Exists(commonFolder))
            {
                commonFolder = Path.Combine(OutputFolder, ConstData.ModuleFolder);
            }

            foreach (string fileName in Directory.GetDirectories(commonFolder)
                .SelectMany(folder => Directory.GetFiles(folder)
                .Select(Path.GetFileNameWithoutExtension)
                .Where(fileName => !AllCommonSheets.Contains(fileName))))
            {
                AllCommonSheets.Add(fileName);
            }
        }

        public static void Reset(InputParam param, IProgress<string> progess = null)
        {
            TimeSetFolder = param.TimeSetFolder;
            PatternFolder = param.PatternFolder;
            PatListFile = param.PatListFile;
            CurrentJob = param.JobName;
            MainFlowSheets = null;
            InputParam = param;
            Project = param.ProjectName;
            AllShmooSetups = new List<ShmooSetup>();
            HipTempFlowDic = new Dictionary<string, SubFlowSheet>();
            ProgFlowDic = new Dictionary<string, List<string>>();
            PatternDic = new Dictionary<string, string>();
            AllFlowStepsDic = new Dictionary<string, FlowRow>();
            PatternsInCharPlan = new List<string>();
            ProgInfo = new ProgInfo();
            OutputFolder = param.OutputFolder;
            //FileUtility.CleanDir(param.OutputFolder);
            MessageWriter = new PostMessageWriter(param.OutputFolder, progess);
            PostReportWriter = new ExcelPackage(new FileInfo(Path.Combine(param.OutputFolder, string.Format("{0}_{1}.xlsx", "PostProcessorReport", TimeContext.Now.ToString("yyyy-MM-dd HHmmss")))));
            HtolAndTtr = new Dictionary<string, List<string>>
            {
                {ConstData.Htol, new List<string>()},
                {ConstData.Ttr, new List<string>()}
            };
            CompileITemDic = new Dictionary<string, CompileITem>();
            FileLoadList = new List<string>();
            TestNumMain.Reset(param.TNumStart);
            if (!string.IsNullOrEmpty(param.ProgWorkBookPath))
            {
                FileAttributes attr = File.GetAttributes(param.ProgWorkBookPath);
                GenAssignSiteVar = param.GenAssignSiteVar;
                IsDirectory = attr.HasFlag(FileAttributes.Directory);

            }
            ExportVersion = param.ExportVersion == "" ? 8.3 : double.Parse(param.ExportVersion);
            _PrintOutInformation(param);

            if (param.IgxlProgram != null)
            {
                TestProgram = param.IgxlProgram;
            }
            else
            {
                TestProgram = new IgxlProgram(param.JobName);
            }

            ProgramUpdateOnly = false;
            DFCSheet = null;
            Progress = null;
            GenSheets = new List<IIgxlSheet>();
            GenOthers = new List<string>();
        }



        private static void _PrintOutInformation(InputParam param)
        {
            MessageWriter.WriteLine("Start Time: " + TimeStamp);
            MessageWriter.WriteLine("General Information");
            MessageWriter.WriteLine("============================================================================");
            MessageWriter.WriteLine("Test Program Path: " + param.ProgWorkBookPath);
            MessageWriter.WriteLine("Test Program File or Directory: " + (IsDirectory ? "Directory" : "File"));
            MessageWriter.WriteLine("Char Plan Intermedia File: " + param.CharFile);
            MessageWriter.WriteLine("Char Setup Output Folder: " + param.OutputFolder);
            MessageWriter.WriteLine("HardIP Information File: " + param.PatInfoFile);
            MessageWriter.WriteLine("Start TNum Number: " + param.TNumStart);
            MessageWriter.WriteLine(param.GenCharNotUse
                ? "Generate Test NoUsed in CharPlan"
                : "Not Generate Test NoUsed in CharPlan");
            MessageWriter.WriteLine(param.GenPatNotUse
                ? "Generate Test NotUsed in PatList"
                : "Not Generate Test NotUsed in PatList");
            MessageWriter.WriteLine(param.GenFlowProdFlow
                ? "Generate Flow Step NoUsed in Production Flow"
                : "Not Generate Flow Step NoUsed in Production Flow");
            MessageWriter.WriteLine("=============================================================================");
        }


    }
    public class PostMessageWriter
    {
        private IProgress<string> _progress;
        private StreamWriter _writer;
        public PostMessageWriter(string logPath, IProgress<string> progess)
        {
            _progress = progess;
            _writer = new StreamWriter(Path.Combine(logPath, "PostProcessor.log"), true);
        }

        public void WriteLine(string msg)
        {
            if (_progress != null)
            {
                _progress.Report(msg);
            }

            _writer.WriteLine(msg);
        }

        public void Close() { _writer.Close(); }
    }
}
