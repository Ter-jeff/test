using System;
using System.Diagnostics;
using System.IO;

using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;

using Teradyne.Oasis.IGLinkBase;

namespace Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor
{
    public class GenIgxlProg
    {
        public static string IgLinkClPath { get; set; }

        public string GenIgxlProgram(string projectName, string mainFlow, string inFolder, string outFolder)
        {
            string outSubFolder = Path.Combine(outFolder, "IGLink");
            Directory.CreateDirectory(outSubFolder);

            string igLinkProjectPath = Path.Combine(outSubFolder, projectName + ".igxlProj");  //ex: D:\CAutoGen\IGLink\Myst.igxlProj

            var objProject = new DeviceProject
            {
                Name = "ProjectTemple",
                SaveAsXLS = false,
                SheetOrder = SheetOrderPreference.Alphabetically
            };


            var objSubProgram = new SubProgram
            {
                Name = projectName,
                JobNames = projectName,
                MainFlow = projectName + ":" + mainFlow,
                GenerateJobListSheet = false
            };

            objProject.SubPrograms.Add(objSubProgram);

            // add files into subProgram
            foreach (string srcFile in Directory.GetFiles(inFolder, "*.*", SearchOption.AllDirectories))
            {
                if (srcFile.ToLower().IndexOf(".txt", StringComparison.Ordinal) != -1)
                {
                    objSubProgram.Add(GetSheetInfo(srcFile, ""));
                }
                else if (srcFile.ToLower().IndexOf(".bas", StringComparison.Ordinal) != -1)
                {
                    objSubProgram.Add(GetVbInfo(srcFile, ""));
                }
                else if (srcFile.ToLower().IndexOf(".cls", StringComparison.Ordinal) != -1)
                {
                    objSubProgram.Add(GetVbInfo(srcFile, ""));
                }
            }

            //STEP3. Create IGLinkObjectExt -> 產生IG-Link Project
            //              GenIGXLProgram  -> 產生IGEX程式
            DeviceProject.SaveProjectCfg(objProject);

            var igLinkProjObj = new IGLinkObjectExt(igLinkProjectPath);  //ex: D:\aaa\IGLink\MYST.igxlProj

            string switchFlag = " -e ";
            string ext = ".xlsm";
            GenerationMode mode = GenerationMode.Excel;

            if (!(LocalSpecs.ExportVersion < 9.0))
            {
                switchFlag = " -g ";
                ext = ".igxl";
                mode = GenerationMode.IGXL;
            }

            string igexProramPath = Path.Combine(outSubFolder, projectName + ext);
            var genIgxlObj = new GenIGXLProgram(igLinkProjObj, null, s => { }, mode, outFolder, projectName + ext, "");

            //3.1 Generate IG-Link project and iG-Excel program
            GenerateIgxlProg(igLinkProjectPath, igexProramPath, objSubProgram.Name, switchFlag);

            return File.Exists(igexProramPath) ? igexProramPath : "";
        }

        public Sheet GetSheetInfo(string filepath, string refPath)
        {
            /* use to assign Sheet.Source(path) */
            var nSheet = new Sheet();
            string tmpStr;

            if (refPath != "")
            {
                tmpStr = filepath.ToLower().IndexOf(refPath.ToLower(), StringComparison.Ordinal) != -1
                    ? filepath.Substring(refPath.Length + 1, filepath.Length - refPath.Length - 1)
                    : filepath;
            }
            else
            {
                tmpStr = filepath;
            }

            nSheet.Source = tmpStr;

            return nSheet;
        }

        public VBFile GetVbInfo(string filepath, string refPath)
        {
            /* use to assign Sheet.Source(path) */
            var nVbModule = new VBFile();
            string tmpStr;
            if (refPath != "")
            {
                tmpStr = filepath.ToLower().IndexOf(refPath.ToLower(), StringComparison.Ordinal) != -1
                    ? filepath.Substring(refPath.Length + 1, filepath.Length - refPath.Length - 1)
                    : filepath;
            }
            else
            {
                tmpStr = filepath;
            }

            nVbModule.Source = tmpStr;
            return nVbModule;
        }

        public void GenerateIgxlProg(string igLinkProjectName, string outputFile
            , string subProgramName = "", string switchName = " -e ", string jobName = "")
        {
            string option = "";
            string igLinkCl = "";
            string oasisRootFolder = Environment.GetEnvironmentVariable("OASISROOT");
            if (oasisRootFolder != null)
            {
                igLinkCl = Path.Combine(oasisRootFolder, "IGLinkCL.exe");
                if (!File.Exists(igLinkCl))
                {
                    if (!string.IsNullOrEmpty(IgLinkClPath))
                    {
                        igLinkCl = Path.Combine(IgLinkClPath, "IGLink", "IGLinkCL.exe");
                    }
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(IgLinkClPath))
                {
                    igLinkCl = Path.Combine(IgLinkClPath, "IGLink", "IGLinkCL.exe");
                }
#pragma warning disable Ter402 // Windows-only IGXL installation path check
                else if (File.Exists(@"C:\Program Files (x86)\Teradyne\Oasis\IGLinkCL.exe"))
                {
                    igLinkCl = @"C:\Program Files (x86)\Teradyne\Oasis\IGLinkCL.exe";
                }
#pragma warning restore Ter402
            }

            if (!File.Exists(igLinkCl))
            {
                return;
            }

            if (subProgramName != "")
            {
                option = "-i " + "\"" + igLinkProjectName + "\"" + " -s " + "\"" + subProgramName + "\"" + switchName + "\"" + outputFile + "\"";
            }
            else if (jobName != "")
            {
                option = "-i " + "\"" + igLinkProjectName + "\"" + " -j " + "\"" + subProgramName + "\"" + switchName + "\"" + outputFile + "\"";
            }

            RunCmd(igLinkCl, option);
        }

        public void RunCmd(string cmdstr, string argment)
        {
            var nProcess = new Process();
            var startInfo = new ProcessStartInfo
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Normal,
                FileName = cmdstr,
                Arguments = argment
            };

            nProcess.StartInfo = startInfo;
            nProcess.Start();
            string result = nProcess.StandardOutput.ReadToEnd();
            nProcess.WaitForExit();
            if (!result.Equals(""))
            {
                //Pipe out error message string in "result"
                //Response.Report(_progress, "Meet an error in generating IGXL program. " + result, 100, MessageLevel.Error);
            }
        }
    }
}
