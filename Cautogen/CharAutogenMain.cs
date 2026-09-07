using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Automation.Static;

using Cautogen.AutoCZ.CharPostProcessor;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPreProcessor;
using Cautogen.AutoCZ.CharPreProcessor.InputReader;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using LogLib.Static;
using LogLib.Utility;

namespace Cautogen
{
    public class CharAutogenMain
    {
        private string job = "";
        private string project = "";
        private string channel = "";
        private string projectPatternFolder = "";
        private string charPlan = "";
        private string patternInfo = "";
        private string patternDashboard = "";
        private string outputDirectory = "";
        private string baseTestProgram = "";
        private string csLibraryPath = "";
        private bool genpatnotuse = false;
        private bool gencharnotuse = false;
        private bool isCheckOnly = false;

        public CharAutogenMain()
        {
        }


        private void UpdateHardIPInfo(string projectPatternFolder)
        {
            LogHelper.Info("Updating HardIP Info......");
            string infoFile = patternInfo;
            string outputPath = Path.Combine(outputDirectory, "hardip_info.log");
            var hardIpAllFiles = new List<string>();
            new PatternListInputReader(patternDashboard).Read();

            string hardIpInfoFolder = Path.Combine(projectPatternFolder, "Hard_IP", "CSV_HardIP_Info");

            //Decide where is pattern info file from
            if (string.IsNullOrEmpty(infoFile) && Directory.Exists(hardIpInfoFolder))
            {
                hardIpAllFiles = Directory.GetFiles(hardIpInfoFolder, "HardIP_AutoGen_Info_All*").ToList();
            }
            else if (Path.GetFileName(infoFile).Equals("hardip_info.log", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
            else if (Directory.Exists(projectPatternFolder))
            {
                hardIpAllFiles = Directory.GetFileSystemEntries(projectPatternFolder, "HardIP_AutoGen_Info_All*", SearchOption.AllDirectories).ToList();
            }
            else if (File.Exists(projectPatternFolder))
            {
                hardIpAllFiles = new List<string>() { projectPatternFolder };
            }
            else
            {
                hardIpAllFiles = new List<string>() { hardIpInfoFolder };
            }

            // Read Pattern_List.csv
            Dictionary<string, string> lfullPatNameTimeSetDict =
                new Dictionary<string, string>();
            lfullPatNameTimeSetDict = PatternListInputReader.PatternList.Values.GroupBy(p => p.GetFullName())
                .ToDictionary(p => p.Key, p => p.Select(q => q.TimesetVersion).First());
            Dictionary<string, string> fullPatNameTimeSetDict = lfullPatNameTimeSetDict;

            Dictionary<string, string> lGenericfullNameDict =
                new Dictionary<string, string>();
            lGenericfullNameDict =
                PatternListInputReader.PatternList.Values.GroupBy(p => p.PatternName.ToUpper())
                    .ToDictionary(p => p.Key, p => p.First().GetFullName());

            var hardIpInfoAllDict = new Dictionary<string, Dictionary<string, string>>();

            // Read HardIP_Info_All
            foreach (string hardIpAllFile in hardIpAllFiles)
            {
                //Update HardIP info all and replace to modified file for further analysis
                if (!string.IsNullOrEmpty(hardIpAllFile) || File.Exists(hardIpAllFile))
                {
                    // Generate hardip_info.log
                    ReadHardIPInfoAll(hardIpInfoAllDict, hardIpAllFile, lGenericfullNameDict);
                }
            }
            GenerateHardIP_PatInfo(fullPatNameTimeSetDict, hardIpInfoAllDict, outputPath, "N/A");

            patternInfo = outputPath;
            //txt_PatternInfoPath.Invoke((MethodInvoker)delegate
            //{
            //    txt_PatternInfoPath.Text = outputPath;
            //});
        }

        private void ReadHardIPInfoAll(Dictionary<string, Dictionary<string, string>> hardIPInfoAll, string hardIpAllFile, Dictionary<string, string> PatListCSV)
        {
            if (File.Exists(hardIpAllFile))
            {

                string line;
                string genericPatName = "";
                string version = "";
                string hardipInfo = "";
                StreamReader file = null;
                int capbit = 0;
                int sendbit = 0;

                try
                {
                    file = new StreamReader(hardIpAllFile);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(string.Format(ex.ToString()));
                    string backupFile = Path.Combine(@"K:\Log\HardIP_Info_All", Path.GetFileName(hardIpAllFile));
                    if (File.Exists(backupFile))
                    {
                        file = new StreamReader(backupFile);
                    }
                }
                finally
                {
                    if (file != null)
                    {
                        while ((line = file.ReadLine()) != null)
                        {
                            if (line.IndexOf("GenericPat:=") != -1)
                            {
                                genericPatName = line.Substring(line.IndexOf(":=") + 2).ToUpper();
                                if (!PatListCSV.ContainsKey(genericPatName))
                                {
                                    genericPatName = "";
                                }
                            }

                            if (string.IsNullOrEmpty(genericPatName))
                            {
                                continue;
                            }

                            if (line.IndexOf("Version:=") != -1)
                            {
                                version = line.Substring(line.IndexOf(":=") + 2);
                            }
                            else if (line.IndexOf("<HardIP_Info_Token>") != -1)
                            {
                                // new genericPat
                                if (genericPatName.Contains("OTP"))
                                {
                                    version = "";
                                    hardipInfo = "";
                                    continue;
                                }
                                if (!hardIPInfoAll.ContainsKey(genericPatName))
                                {
                                    Dictionary<string, string> verInfo = new Dictionary<string, string>();
                                    verInfo.Add(version, hardipInfo);
                                    hardIPInfoAll.Add(genericPatName, verInfo);
                                }
                                // existed generic Pat new Version
                                else
                                {
                                    if (!hardIPInfoAll[genericPatName].ContainsKey(version))
                                    {
                                        hardIPInfoAll[genericPatName].Add(version, hardipInfo);
                                    }
                                }
                                // clean up variables
                                genericPatName = "";
                                version = "";
                                hardipInfo = "";
                                capbit = 0;
                                sendbit = 0;
                            }
                            else if (line != "")
                            {
                                if (line.IndexOf("Cap Bit:") != -1)
                                {
                                    capbit = Convert.ToInt32(line.Substring(line.IndexOf(":") + 1));
                                }
                                if (line.IndexOf("Send Bit:") != -1)
                                {
                                    sendbit = Convert.ToInt32(line.Substring(line.IndexOf(":") + 1));
                                }

                                //if (capbit < 50000 && sendbit < 50000)
                                //{
                                // modify by JN for separated DSSC bit if DSSC bit is bigger than 24   // C651 AP team STDF will trcade
                                // 32:d2_pll_lock => 24:d2_pll_lock_1,8:d2_pll_lock_2 
                                if (line.IndexOf("DSSC_OUT,") != -1)
                                {
                                    line = "DSSC_OUT," + SplitDsscBy24(line); // rewrite dssc line
                                }
                                hardipInfo += line + "\r\n";
                                //}
                            }
                        }
                        file.Close();
                    }
                }
            }
        }

        private static string SplitDsscBy24(string dsscOutLine)
        {
            var modDssc = new List<string>();
            var rgexDSSC = new Regex(@"(?<dssc>(?<bit>\d+):(?<reg>.*))", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            foreach (string dsscOut in dsscOutLine.Split(',').ToList())
            {
                foreach (Match m in rgexDSSC.Matches(dsscOut))
                {
                    int bits = Convert.ToInt32(m.Groups["bit"].Value);
                    if (bits > 24)
                    {
                        int quotient = bits / 16;
                        int reminder = bits - (16 * quotient);
                        for (int i = 1; i <= quotient; i++)
                        {
                            string tempdssc = "16:" + m.Groups["reg"].Value + "_" + i;
                            modDssc.Add(tempdssc);
                        }
                        quotient++;

                        if (reminder != 0)
                        {
                            string reminderdssc = reminder + ":" + m.Groups["reg"].Value + "_" + quotient;
                            modDssc.Add(reminderdssc);
                        }
                    }
                    else
                    {
                        modDssc.Add(m.Groups["dssc"].Value);
                    }
                }
            }
            return string.Join(",", modDssc);
        }

        private void GenerateHardIP_PatInfo(Dictionary<string, string> fullPatNameTimeSetDict, Dictionary<string, Dictionary<string, string>> HardIPInfoAllDict, string outputPath, string patternListFile)
        {
            int count = 1;
            string tempUpdateStr = "";
            var builderList = new List<StringBuilder>();
            var updateStr = new StringBuilder();
            updateStr.AppendLine("############################################################################################################");
            updateStr.AppendLine("## " + patternListFile);
            updateStr.AppendLine("## Greated By Hard_IP Auto Gen_Core ");
            updateStr.AppendLine("############################################################################################################");

            foreach (string fullPatName in fullPatNameTimeSetDict.Keys)
            {
                if (fullPatName.Equals("N/A"))
                {
                    continue;
                }

                string genericPatName = GetGenericPatName(fullPatName);
                string patVersion = GetVersion(fullPatName).ToUpper();
                if (HardIPInfoAllDict.ContainsKey(genericPatName))
                {

                    if (HardIPInfoAllDict[genericPatName].ContainsKey(patVersion))
                    {
                        tempUpdateStr = HardIPInfoAllDict[genericPatName][patVersion];
                        tempUpdateStr = tempUpdateStr.Replace("1)===", count.ToString() + ")===");
                        tempUpdateStr = tempUpdateStr.Replace("Tset: NA", "Tset: " + fullPatNameTimeSetDict[fullPatName]);
                        updateStr.AppendLine(tempUpdateStr);
                        count++;
                    }

                }
                if (updateStr.Length >= 10000000)
                {
                    builderList.Add(updateStr);
                    updateStr = new StringBuilder();
                }
            }
            //if (builderList.Count == 0) 
            builderList.Add(updateStr);

            UpdateFileSafely(outputPath, builderList);
        }

        private static string GetGenericPatName(string fullPatternName)
        {
            // fullPatternName: "D:\\HardIP_Info\\miner\\PP_MINA0_A_FULP_AN_AA03_MEA_JTG_FRQ_ALLFV_SI_MEAS4_1_A0_1510070021.ATP"
            // genericPatName:  "PP_MINA0_A_FULP_AN_AA03_MEA_JTG_FRQ_ALLFV_SI_MEAS4"
            string patternName = Path.GetFileName(fullPatternName);
            string[] split = patternName.Split(new char[] { '_' });

            string genericPatName = "";
            for (int i = 0; i < split.Length - 3; i++)
            {
                genericPatName += split[i];
                if (i < split.Length - 4)
                { genericPatName += "_"; }
            }
            return genericPatName.ToUpper();
        }

        private static string GetVersion(string fullPatternName)
        {
            try
            {


                // fullPatternName: "D:\\HardIP_Info\\miner\\PP_MINA0_A_FULP_AN_AA03_MEA_JTG_FRQ_ALLFV_SI_MEAS4_1_A0_1510070021.ATP"
                // genericPatName:  "1_A0_1510070021"
                string patternName = Path.GetFileName(fullPatternName);
                string[] split = patternName.Split(new char[] { '_' });

                string versionStr = "";
                for (int i = split.Length - 3; i < split.Length; i++)
                { versionStr += split[i] + '_'; }
                versionStr = versionStr.Substring(0, versionStr.Length - 1); // remove "_"

                // remove ".ATP"
                if (versionStr.IndexOf(".") != -1)
                {
                    versionStr = versionStr.Substring(0, versionStr.IndexOf('.'));
                }

                return versionStr;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(string.Format(ex.ToString()));
                return "";
            }
        }

        private static void UpdateFileSafely(string filePath, IEnumerable<StringBuilder> builderList)
        {
            CheckFileExist(filePath);

            // Copy file to file.tmp
            string tempFilePath = Path.ChangeExtension(filePath, ".tmp");
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
            //File.Copy(filePath, tempFilePath);

            //// Update .tmp file
            //if (useAppend)
            //{
            //    FileStream fs = new FileStream(tempFilePath, FileMode.Open, FileAccess.ReadWrite);
            //    fs.Seek(0, SeekOrigin.End);

            //    StreamWriter sw = new StreamWriter(fs);
            //    sw.WriteLine(writeLines);
            //    sw.Close();
            //    fs.Close();
            //}
            //else
            //{ // overwrite file
            //File.WriteAllText(tempFilePath, writeLines.ToString());
            //}

            // create and write tmp file
            FileStream fileStream = File.Create(tempFilePath);
            var streamWriter = new StreamWriter(fileStream);
            foreach (StringBuilder writeLines in builderList)
            {
                fileStream.Seek(0, SeekOrigin.End);
                streamWriter.WriteLine(writeLines);
            }
            streamWriter.Close();
            fileStream.Close();
            // Delet the original file
            File.Delete(filePath);

            // Rename .tmp file to the original file
            File.Move(tempFilePath, filePath);
        }

        private static bool CheckFileExist(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                { return true; }
                else
                {
                    File.Create(filePath).Close();
                    return false;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(string.Format(ex.ToString()));
                return false;
            }
        }

        //public override ICommandLineOptions ValidateInput(ICommandLineOptions options)
        //{
        //    var arguments = (CharAutogenOptions)options;

        //    job = arguments.JobName;
        //    project = arguments.CurrentProjectName;
        //    channel = arguments.ChannelMap;
        //    projectPatternFolder = arguments.PatternFolder;
        //    if (!Directory.Exists(projectPatternFolder))
        //    {
        //        throw new Exception($"Cannot Found Pattern Folder : {projectPatternFolder}");
        //    }
        //    charPlan = arguments.CharPlan;


        //    if (!File.Exists(charPlan) && !Directory.Exists(charPlan))
        //    {

        //        throw new Exception($"Cannot Found Char Plan : {charPlan}");

        //    }
        //    patternInfo = arguments.PatternInfo;
        //    if (!Directory.Exists(patternInfo))
        //    {
        //        throw new Exception($"Cannot Found Pattern Info File Folder : {patternInfo}");
        //    }
        //    patternDashboard = arguments.PatternDashboard;
        //    if (!File.Exists(patternDashboard))
        //    {
        //        throw new Exception($"Cannot Found Pattern Dashboard : {projectPatternFolder}");

        //    }
        //    outputDirectory = arguments.OutputDirectory;

        //    baseTestProgram = arguments.BaseTestProgram;
        //    if (!File.Exists(baseTestProgram))
        //    {
        //        throw new Exception($"Cannot Found IGXL Test Program : {baseTestProgram}");

        //    }
        //    csLibraryPath = arguments.CSLibraryPath;
        //    if (!Directory.Exists(csLibraryPath))
        //    {
        //        throw new Exception($"Cannot Found CSharp solution folder : {csLibraryPath}");
        //    }


        //    return options;
        //}

        public void Execute()
        {
            //var arguments = (CharAutogenOptions)options;
            //job = LocalSpecs.AllJobs.FirstOrDefault();
            job = LocalSpecs.AllJobs.Count == 0 ? LocalSpecs.CurrentJob : LocalSpecs.AllJobs.FirstOrDefault();
            project = LocalSpecs.CurrentProject;
            channel = LocalSpecs.DefaultChannelMap;
            projectPatternFolder = LocalSpecs.PatternFolder;
            charPlan = LocalSpecs.CharPlanFileName;
            patternInfo = LocalSpecs.HardIpInfoFileName;
            patternDashboard = LocalSpecs.PatternListCsvFileName;
            outputDirectory = LocalSpecs.TarFolder;
            baseTestProgram = string.IsNullOrEmpty(LocalSpecs.BaseTestProgram) ? "" : LocalSpecs.BaseTestProgram;
            csLibraryPath = LocalSpecs.CsLibraryFolder;
            genpatnotuse = false;
            gencharnotuse = false;

            isCheckOnly = LocalSpecs.CheckOnly;

            Response.Report($"Argument CurrentProjectName {project}");
            Response.Report($"Argument JobName {job}");
            Response.Report($"Argument ChannelMap {channel}");
            Response.Report($"Argument CharPlan {charPlan}");
            Response.Report($"Argument PatternInfo {patternInfo}");
            Response.Report($"Argument PatternDashboard {patternDashboard}");
            Response.Report($"Argument BaseTestProgram {baseTestProgram}");
            Response.Report($"Argument CSLibraryPath {csLibraryPath}");
            Response.Report($"Argument PatternFolder {projectPatternFolder}");
            Response.Report($"Argument OutputDirectory {outputDirectory}");
            Response.Report($"Argument GenPatternNotUse {genpatnotuse}");
            Response.Report($"Argument GenCharNotUse {gencharnotuse}");

            Response.Report("Create Output folder...");
            //var postProcessOutFolder = Path.Combine(outputDirectory, "char_result");

            //if (!Directory.Exists(postProcessOutFolder))
            //    Directory.CreateDirectory(postProcessOutFolder);

            if (!string.IsNullOrEmpty(patternInfo))
            {
                UpdateHardIPInfo(patternInfo);
            }
            else
            {
                UpdateHardIPInfo(projectPatternFolder);
            }

            #region Pre-Process
            var newParam = new ParamData
            {
                JobName = job,
                CharPlanFile = charPlan,
                PatinfoFile = patternInfo,
                PatListFile = patternDashboard,
                TarDic = outputDirectory,
                ProjectName = project,
                NvIsChecked = false,
                BaseProgram = baseTestProgram,
                ShmooPowerPinHightoLow = true,
                CharPreCheckForNewTChar = true,
                CommandLineMode = true,
            };
            var preProcessor = new PreProcessFunc(newParam);
            preProcessor.Run(isCheckOnly);
            #endregion

            #region Post-Process
            if (!isCheckOnly)
            {
                var postProcessor = new PostProcessorFunc(new InputParam
                {
                    GenFlowProdFlow = true,
                    GenCharNotUse = gencharnotuse,
                    GenPatNotUse = genpatnotuse,
                    GenTxtOnly = false,
                    GenTNum = true,
                    GenTmpsOnFlow = true,
                    GenPmode = true,
                    GenPatSub = true,
                    GenReadEcid = true,
                    CharPlan = charPlan,
                    PatListFile = patternDashboard,
                    CharFile = preProcessor.ImFilePath,
                    JobName = job,
                    ChannelMapName = channel,
                    OutputFolder = outputDirectory,
                    PatInfoFile = patternInfo,
                    ProgWorkBookPath = baseTestProgram,
                    ExportVersion = "9.0",
                    ProjectName = project,
                    PostSettings = "",
                    TimeSetFolder = Path.Combine(projectPatternFolder, "TimeSet"),
                    PatternFolder = projectPatternFolder,
                    FlowTmpsName = "TMPS",
                    TNumStart = "10000",
                    IgnoreHfLimits = false,
                    RunPayloadAfterSelsram = false,
                    GenCSharp = true,
                    CsLibraryPath = csLibraryPath,
                    IgxlProgram = newParam.IgxlProgram,

                });

                string enableWord;
                postProcessor.Execute(out enableWord);
            }
            #endregion
            Response.Report("Finished !!!");

        }
    }
}
