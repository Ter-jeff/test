using System;
using System.Collections.Generic;
using System.IO;

using Automation.Static;

using BinCutScriptLib.Reader;

using EfuseCheckCmdLib.EFuse;
using EfuseCheckCmdLib.EFuse.EFuseApp;

using LogLib.Static;
using LogLib.Utility;

using NLog;

namespace AutogenCommandLine.Tools.CheckScript
{
    public class EFuse
    {
        private static string _inDir;
        private static string _bitDef;
        private static string _cfg;
        private static string _testProg;
        private static string _outputDir;
        private static string _testPlanFile;

        public EFuse(string inDir, string bitDef, string cfg, string testPlanFile, string testProg, string outputDir)
        {
            _inDir = inDir;
            _bitDef = bitDef;
            _cfg = cfg;
            _outputDir = outputDir;
            _testPlanFile = testPlanFile;
            _testProg = testProg;
        }

        [STAThread]
        public static void RunEfuse()
        {
            DataLogReader.CheckDatalog(_inDir, out bool csFlag);
            if (!csFlag)
            {
                EfuseStatic.OutputPath = Path.Combine(_outputDir, "Report", "Efuse");
                LogHelper.SetNLog(Path.Combine(_outputDir, "Log", "Efuse.log"));
                if (string.IsNullOrEmpty(_bitDef) && string.IsNullOrEmpty(_cfg) && string.IsNullOrEmpty(_testProg))
                {
                    AppendRichText("Has multiple input cfg, bdf and test program", "Blue");
                    AppendRichText("The execution result is : Fail.", "Blue");
                    return;
                }
                try
                {
                    RunValidationStandAlone();
                }
                catch (Exception e)
                {
                    EfuseStatic.Result = EfuseCheckResultType.Exception;
                    AppendRichText(e.StackTrace, "Blue");
                }
                switch (EfuseStatic.Result)
                {
                    case EfuseCheckResultType.Pass:
                        AppendRichText("The execution result is : Pass.", "Blue");
                        break;
                    case EfuseCheckResultType.Fail:
                        AppendRichText("The execution result is : Fail.", "Blue");
                        break;
                    case EfuseCheckResultType.Exception:
                        AppendRichText("The execution result is : Exception.", "Blue");
                        break;
                    default:
                        AppendRichText("The execution result is : Exception.", "Blue");
                        break;
                }
            }
            else //C#
            {
                _ = new EfuseAlgorithmCheckCs(_inDir, _bitDef, _cfg, _outputDir);
                EfuseAlgorithmCheckCs.WorkFlow();
            }
            Response.Report("Finished EFuse Check Script...");
        }

        private static void RunValidationStandAlone()
        {
            HarvestFieldReader.HarvestFieldList.Clear();
            string waferIdFile = "";
            var stdfFile = new List<string>();
            EfuseAlgorithmCheck exeCheckObj = new EfuseAlgorithmCheck(AppendRichText, _inDir, _bitDef, _cfg, _testPlanFile, _testProg, waferIdFile, stdfFile, false, EfuseStatic.OutputPath);
            exeCheckObj.WorkFlow(false);
        }

        public static void AppendRichText(string text, string color)
        {
            Logger logger = LogManager.GetCurrentClassLogger();
            logger.Trace(text);
        }
    }
}
