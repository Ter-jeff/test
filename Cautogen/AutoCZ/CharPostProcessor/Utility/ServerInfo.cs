using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility
{
    public class SubrPatInfo
    {
        public List<string> Subroutine { get; set; }
        public string VmVector { get; set; }

        public SubrPatInfo()
        {
            Subroutine = new List<string>();
            VmVector = null;
        }
    }

    public static class ServerInfo
    {
        public static Dictionary<string, string> PatlistDict = new Dictionary<string, string>();
        public static List<string> Errorlist = new List<string>();

        //public static void SelectPatterns(string Project, Dictionary<string, string> PatListToFind, int totalPatterns, string outputfileDir, string patSetDir, bool ftpStru, bool isUfpPat, IProgress<ProgressStatus> progress = null)
        //{
        //    var args = new ProgressStatus { Percentage = 0, EnableMsg = true, Level = MessageLevel.EndPoint, Result = "Copying Patterns from Server to Local. ..." }; //To Report Progress
        //    if (progress != null) progress.Report(args);
        //    var copyCount = 0;
        //    /*var allDrives = DriveInfo.GetDrives();
        //    var drivePath = string.Empty;
        //    foreach (var d in allDrives)
        //    {
        //        if (d.IsReady) { if (d.VolumeLabel == "S_BACKUP") drivePath = d.Name; }
        //    }

        //    if (Environment.MachineName == "STDP-SVN-S") drivePath = drivePath + @"\VPN\";

        //    drivePath = Path.Combine(drivePath, Project);*/
        //    PatlistDict = PatListToFind;
        //    Dictionary<string, string> DicPatPath = null;
        //    #region //Check for PMIC
        //    if (LocalSpecs.TSMC_ENV && DicPatPath == null)
        //    {
        //        ProjectConfigSingleton.Instance().LoadProjectConfig(LocalSpecs.IniFileName);
        //        var deviceType = ProjectConfigSingleton.Instance().GetConfigSettingValue("Device", "Type");
        //        if (deviceType == "PMIC" || ftpStru)
        //        {
        //            //var allDrives = DriveInfo.GetDrives();
        //            var dbPath = LocalSpecs.PatternPath + @"\Log\CompiledPat\" + LocalSpecs.CurrentProject.ToLower() + ".db";
        //            //var filename = LocalSpec.LocalSpecs.CurrentProject.ToLower() + "_CompiledPat.csv";
        //            DicPatPath = DownloadList.GetDownloadList(dbPath);
        //        }
        //    }
        //    foreach (var item in PatlistDict)
        //    {
        //        var fullPath = string.Empty;
        //        if (Regex.IsMatch(item.Key, @"_DUPLICATE.?", RegexOptions.IgnoreCase))
        //        {
        //            var patname = item.Key;
        //            patname = Regex.Replace(patname, @"_DUPLICATE.+|_DUPLICATE", "", RegexOptions.IgnoreCase);
        //            fullPath = Path.Combine(patSetDir, item.Value, patname);
        //        }
        //        else
        //        {
        //            fullPath = isUfpPat ? Path.Combine(patSetDir, "patx", item.Value, item.Key) : Path.Combine(patSetDir, item.Value, item.Key);
        //        }

        //    #endregion
        //        if (DicPatPath != null)
        //            CopyToLocal(fullPath, item.Key, @".\Pattern\" + DicPatPath[item.Key.ToUpper()], Project, outputfileDir, isUfpPat);
        //        else
        //            CopyToLocal(fullPath, item.Key, item.Value, Project, outputfileDir, isUfpPat);
        //        copyCount += 1;
        //        args.Percentage = Convert.ToInt32((double)copyCount * 100 / (double)totalPatterns);
        //        args.EnableMsg = false;
        //        args.Result = "";
        //        progress.Report(args);

        //    }
        //    Thread.Sleep(50);
        //    if (Errorlist.Count > 0)
        //    {
        //        args.Percentage = 0;
        //        args.EnableMsg = true;
        //        args.Level = MessageLevel.Error;
        //        args.Result = "Found Missing Patterns in Server... Generating Report";
        //        if (progress != null) progress.Report(args);
        //        Thread.Sleep(50);
        //        //var targetPath = Path.Combine(outputfileDir);
        //        var errorfile = "Errors_Patterns_not_found_Summary.txt";
        //        var errorpath = Path.Combine(outputfileDir, errorfile);
        //        if (!Directory.Exists(outputfileDir)) Directory.CreateDirectory(outputfileDir);
        //        if (!File.Exists(errorpath))
        //        {
        //            var file = File.Create(errorpath);
        //            file.Close();

        //        }
        //        TextWriter tw = new StreamWriter(errorpath);

        //        foreach (var s in Errorlist)
        //            tw.WriteLine(s);

        //        tw.Close();
        //        //File.WriteAllLines(errorpath, Errorlist);

        //        args.Percentage = 100;
        //        args.EnableMsg = true;
        //        args.Level = MessageLevel.Error;
        //        args.Result = "Missing patterns report is generated... Check " + errorpath;
        //        if (progress != null) progress.Report(args);
        //        Thread.Sleep(50);

        //    }
        //    else
        //    {
        //        args.Percentage = 100;
        //        args.EnableMsg = true;
        //        args.Result = "All the patterns are copied. No missing patterns";
        //        if (progress != null) progress.Report(args);
        //        Thread.Sleep(50);
        //    }
        //    args.Percentage = 100;
        //    args.EnableMsg = true;
        //    args.Level = MessageLevel.EndPoint;
        //    args.Result = "Copy to local done";
        //    if (progress != null) progress.Report(args);
        //    Thread.Sleep(50);


        //}
        //public static void CopyToLocal(string sourcePath, string PatName, string Folder, string Project, string Outputfolder, bool isUfpPat)
        //{
        //    //var targetPath = Path.Combine(Outputfolder, Project,Folder); m9 request don't combin project name
        //    var targetPath = Path.Combine(Outputfolder, Folder);

        //    var patSufix = isUfpPat ? ".PATX.gz" : ".PAT.gz";
        //    var fullPath1 = sourcePath + ".ATP.gz";
        //    var fullPath2 = sourcePath + patSufix;
        //    var serverinfo1 = new FileInfo(fullPath1);
        //    var serverinfo2 = new FileInfo(fullPath2);
        //    var patname = PatName;
        //    if (serverinfo1.Exists && serverinfo1.Directory != null)
        //    {
        //        if (!Directory.Exists(targetPath))
        //        {
        //            Directory.CreateDirectory(targetPath);
        //        }
        //        patname = patname + ".ATP.gz";
        //        targetPath = Path.Combine(targetPath, patname);
        //        if (!File.Exists(targetPath))
        //        {
        //            File.Copy(fullPath1, targetPath, true);
        //        }
        //    }
        //    else if (serverinfo2.Exists && serverinfo2.Directory != null)
        //    {
        //        if (!Directory.Exists(targetPath))
        //        {
        //            Directory.CreateDirectory(targetPath);
        //        }
        //        patname = patname + patSufix;
        //        targetPath = Path.Combine(targetPath, patname);
        //        if (!File.Exists(targetPath))
        //        {
        //            File.Copy(fullPath2, targetPath, true);
        //        }

        //    }
        //    else
        //    {

        //        Errorlist.Add(patname + " is not found in Server in " + Project + " project " + Folder + " folder.");


        //    }



        //}

        public static Dictionary<string, SubrPatInfo> ReadHardIpInfoAll(string hardipInfoFile)
        {
            var patternInfoAll = new Dictionary<string, SubrPatInfo>();

            if (!File.Exists(hardipInfoFile))
            {
                return patternInfoAll;
            }

            string line;
            string genericPatName = "";
            string version = "";
            var patterninfo = new SubrPatInfo();
            var file = new StreamReader(hardipInfoFile);
            while ((line = file.ReadLine()) != null)
            {
                if (line.StartsWith("GenericPat:=", StringComparison.CurrentCultureIgnoreCase))
                {
                    genericPatName = line.Substring(line.IndexOf(":=") + 2);
                }
                else if (line.StartsWith("Version:=", StringComparison.CurrentCultureIgnoreCase))
                {
                    version = line.Substring(line.IndexOf(":=") + 2);
                }
                else if (line.StartsWith("<HardIP_Info_Token>", StringComparison.CurrentCultureIgnoreCase))
                {
                    // new Pat
                    string fullpattern = genericPatName + "_" + version;
                    if (!patternInfoAll.ContainsKey(fullpattern.ToUpper()))
                    {
                        patternInfoAll.Add(fullpattern.ToUpper(), patterninfo);
                    }
                    // clean up variables
                    genericPatName = "";
                    version = "";
                    patterninfo = new SubrPatInfo();

                }
                else if (line != "")
                {
                    if (line.StartsWith("Subr:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        // "Subr: A-Z | a-z | 0-9 | _
                        //  Subr: ht_skya0_c_fulp_an_aa02_mea_jtg_imx_allfv_0000000000_si_t4p10_srm_meas1, ht_skya0_c_fulp_an_aa02_mea_jtg_imx_allfv_0000000000_si_t4p10_srm_meas2
                        //Regex m = new Regex(@"(Subr:)((\s)+)(?<Subroutine>(([A-Z]|[a-z]|[0-9]|[_])+))");
                        //Match match = m.Match(line);
                        //if (match.Success)
                        //{
                        //    patterninfo.Subroutine = match.Groups["Subroutine"].Value;
                        //}
                        string[] subList = line.Split(':')[1].Trim().Split(',');
                        foreach (string sub in subList)
                        {
                            patterninfo.Subroutine.Add(sub);
                        }
                    }

                    if (line.StartsWith("VM_Vector:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        // "VM_Vector: A-Z | a-z | 0-9 | _
                        var m1 = new Regex(@"(VM_Vector:)((\s)+)(?<VM_vector>(([A-Z]|[a-z]|[0-9]|[_])+))");
                        Match match = m1.Match(line);
                        if (match.Success)
                        {
                            patterninfo.VmVector = match.Groups["VM_vector"].Value;
                        }
                    }
                }
            }
            file.Close();
            return patternInfoAll;
        }
    }
}

