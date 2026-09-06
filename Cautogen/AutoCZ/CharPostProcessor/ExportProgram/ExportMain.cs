using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.InputReader;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;

using Teradyne.Oasis.IGData;

namespace Cautogen.AutoCZ.CharPostProcessor.ExportProgram
{
    public class ExportMain
    {
        public static void Export(string progPath, string outputFolder, bool genTxtOnly, bool isUsedUpdateTp = false)
        {
            if (Path.GetExtension(progPath) != ".xlsm")
            {
                // for xlsm >= 9.0
                _ExportFor90(progPath, outputFolder, genTxtOnly, isUsedUpdateTp);
                return;
            }

            LocalSpecs.MessageWriter.WriteLine("Export prog and initial IG-Link structure... ");

            string exportfolder = Path.Combine(outputFolder, "exportProg");
            ExportWorkBookCmd(progPath, exportfolder);

            LocalSpecs.FileLoadList = AccessFileList(exportfolder);
            if (genTxtOnly)
            {
                return;
            }

            LocalSpecs.FileStructure = _GetFileStructureFromIGLinkManifest(outputFolder, isUsedUpdateTp);
            if (LocalSpecs.FileStructure.Count == 0)
            {
                return;
            }

            // copy the export files into ig-link folder structure
            foreach (string src in Directory.GetFiles(exportfolder))
            {
                string fileName = Path.GetFileName(src);
                if (fileName == null || !LocalSpecs.FileStructure.TryGetValue(fileName, out string dst))
                {
                    continue;
                }

                string dstFolder = Path.GetDirectoryName(dst);
                if (dstFolder == null)
                {
                    continue;
                }

                // copy file if ...
                if (ConstData.KeptModuleList.Exists(
                    a => Regex.IsMatch(dstFolder, a, RegexOptions.IgnoreCase)) ||  // 1. is kept moudle
                    dstFolder.Contains(ConstData.CommonFolder) ||                  // 2. is in preserve fils list
                    PostSettings.PreservedSheets.Contains(fileName) ||             // 3. folder name is HardIP$ and filename is *_IDS
                    (Regex.IsMatch(dstFolder, "HardIP$", RegexOptions.IgnoreCase) && Regex.IsMatch(fileName, "_IDS", RegexOptions.IgnoreCase)))
                {
                    CheckFolderExist(dstFolder);
                    File.Copy(src, dst, true);
                }
            }
            LocalSpecs.SetAllKeptSheets();
        }

        private static void _ExportFor90(string progPath, string outputFolder, bool genTxtOnly, bool isusedUpdateFlag = false)
        {
            LocalSpecs.MessageWriter.WriteLine("Export prog and initial IG-Link structure... ");

            string exportfolder = Path.Combine(outputFolder, "exportProg");
            ExportWorkBookCmd(progPath, exportfolder);
            //ExportWorkbook.Export(progPath, exportfolder);
            LocalSpecs.FileLoadList = AccessFileList(exportfolder);
            if (genTxtOnly)
            {
                return;
            }

            LocalSpecs.FileStructure = _GetFileStructureFromIGLinkManifest(outputFolder, isusedUpdateFlag);
            if (LocalSpecs.FileStructure.Count == 0)
            {
                return;
            }

            // copy the export files into ig-link folder structure
            foreach (string src in Directory.GetFiles(exportfolder))
            {
                try
                {
                    string fileName = Path.GetFileName(src);
                    if (fileName == null || !LocalSpecs.FileStructure.TryGetValue(fileName, out string dst))
                    {
                        continue;
                    }

                    string dstFolder = Path.GetDirectoryName(dst);
                    if (dstFolder == null)
                    {
                        continue;
                    }

                    // copy file if ...
                    if (isusedUpdateFlag)
                    {
                        CheckFolderExist(dstFolder);
                        File.Copy(src, dst, true);
                    }
                    else if (ConstData.KeptModuleList.Exists(
                        a => Regex.IsMatch(dstFolder, a, RegexOptions.IgnoreCase)) || // 1. is kept moudle
                             dstFolder.Contains(ConstData.CommonFolder) || // 2. is in preserve fils list
                             PostSettings.PreservedSheets.Contains(fileName) ||
                             // 3. folder name is HardIP$ and filename is *_IDS
                             (Regex.IsMatch(dstFolder, "HardIP$", RegexOptions.IgnoreCase) &&
                              Regex.IsMatch(fileName, "_IDS", RegexOptions.IgnoreCase)))
                    {
                        CheckFolderExist(dstFolder);
                        File.Copy(src, dst, true);
                    }
                }
                catch (Exception e)
                {
                    if (LocalSpecs.MessageWriter != null)
                    {
                        LocalSpecs.MessageWriter.WriteLine("Generating IG-XL program failed " + e.Message);
                    }
                }

            }

            LocalSpecs.SetAllKeptSheets();
        }

        public static void CopyExportFile(string outputFolder, bool isUsedUpdateTp = false)
        {
            LocalSpecs.FileStructure = _GetFileStructureFromIGLinkManifest(outputFolder, isUsedUpdateTp);
            if (LocalSpecs.FileStructure.Count == 0)
            {
                return;
            }

            string exportfolder = Path.Combine(outputFolder, "exportProg");

            // copy the export files into ig-link folder structure
            _ = Directory.GetFiles(exportfolder).ToList();
            foreach (string src in Directory.GetFiles(exportfolder))
            {
                string fileName = Path.GetFileName(src);
                if (fileName == null || !LocalSpecs.FileStructure.TryGetValue(fileName, out string dst))
                {
                    continue;
                }

                if (Regex.IsMatch(dst, "tmps", RegexOptions.IgnoreCase))
                {
                }
                string dstFolder = Path.GetDirectoryName(dst);
                if (dstFolder == null)
                {
                    continue;
                }

                // copy file if ... 
                //if (ConstData.KeptModuleList.Exists(
                //    a => Regex.IsMatch(dstFolder, a, RegexOptions.IgnoreCase)) ||  // 1. is kept moudle
                //    dstFolder.Contains(ConstData.CommonFolder) ||                  // 2. is in preserve fils list
                //    PostSettings.PreservedSheets.Contains(fileName) ||             // 3. folder name is HardIP$ and filename is *_IDS
                //    (Regex.IsMatch(dstFolder, "HardIP$", RegexOptions.IgnoreCase) && Regex.IsMatch(fileName, @"_IDS", RegexOptions.IgnoreCase))||
                //    Regex.IsMatch(dst,"tmps",RegexOptions.IgnoreCase))
                //{
                //    CheckFolderExist(dstFolder);
                //    File.Copy(src, dst, true);
                //}
                CheckFolderExist(dstFolder);
                File.Copy(src, dst, true);
            }
            LocalSpecs.SetAllKeptSheets();
        }

        public static void ExportWorkBookCmd(string testprogramname, string exportfolder)
        {
            CheckFolderExist(exportfolder);
            string option = "-w \"" + testprogramname + "\" -d \"" + exportfolder + "\"";
            // assueme ExportWorkbook is in the PATH
            RunCmd("ExportWorkbook", option);
        }

        private static void RunCmd(string cmd, string argment = "")
        {
            var nProcess = new Process();
            var startInfo = new ProcessStartInfo
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                FileName = cmd,
                Arguments = argment
            };
            nProcess.StartInfo = startInfo;
            nProcess.Start();
            nProcess.WaitForExit();
        }

        private static void CheckFolderExist(string folder)
        {
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
        }

        private static Dictionary<string, string> _GetFileStructureFromIGLinkManifest(string outputFolder, bool isUsedUpdateTp)
        {
            string regCustomPath = @"trunk(?<cus>.*)[/\\][(Common)|(Module)]";
            var fileStructure = new Dictionary<string, string>();
            string manifestfile = Path.Combine(outputFolder, "exportProg", "IGLinkManifest.txt");
            if (!File.Exists(manifestfile))
            {
                return fileStructure;
            }

            List<string> allfileInfo = File.ReadAllLines(manifestfile).ToList();
            string iglinkFolder = Path.GetDirectoryName(allfileInfo[0].Split('\t')[1].Split(',')[0]);
            for (int i = 6; i < allfileInfo.Count; i++)
            {
                List<string> line = allfileInfo[i].Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                if (line.Count != 3 && !(line.Count == 4 && line[3] == ""))
                {
                    continue;
                }

                if (line[0] == "Generated by IG-Link")
                {
                    continue;
                }

                string sourceFile = line[1].Replace("%20", " ");
                sourceFile = !string.IsNullOrEmpty(iglinkFolder) ? sourceFile.Replace(iglinkFolder, "").TrimStart('\\', '/') : sourceFile;
                string filePath = Path.Combine(outputFolder, "IGLink",
                                            //Regex.IsMatch(sourceFile, "^trunk", RegexOptions.IgnoreCase) ? "IGLink" : Path.Combine("IGLink", "trunk"),
                                            sourceFile);
                string fileName = line[2];
                if (Regex.IsMatch(filePath, regCustomPath, RegexOptions.IgnoreCase))
                {
                    _ = Regex.Match(filePath, regCustomPath, RegexOptions.IgnoreCase).Groups["cus"].Value;

                    try
                    {
                        //if (!string.IsNullOrEmpty(cusPath))
                        //filePath = filePath.Replace(cusPath, "");
                    }
                    catch (Exception)
                    {
                    }
                }
                if (fileName.Contains("JobList") && !isUsedUpdateTp)
                {
                    continue;
                }

                fileStructure[fileName] = filePath;
            }
            return fileStructure;
        }

        public static List<string> AccessFileList(string path)
        {
            var sheetTypeToLoad = new List<string>
            {
                nameof(Sheet.SheetTypes.DTBintablesSheet),
                nameof(Sheet.SheetTypes.DTChanMap),
                nameof(Sheet.SheetTypes.DTFlowtableSheet),
                nameof(Sheet.SheetTypes.DTJobListSheet),
                nameof(Sheet.SheetTypes.DTTestInstancesSheet),
                nameof(Sheet.SheetTypes.DTPatternSetSheet),
                nameof(Sheet.SheetTypes.DTPinMap),
                nameof(Sheet.SheetTypes.DTTimesetSheet),
                nameof(Sheet.SheetTypes.DTTimesetBasicSheet),
                nameof(Sheet.SheetTypes.DTDCSpecSheet),
            };

            string[] files = Directory.GetFiles(path);
            return (from file in files
                    let header = _ReadSheetHeader(file)
                    where sheetTypeToLoad.FirstOrDefault(p => Regex.IsMatch(header, p, RegexOptions.IgnoreCase)) != null
                    select file).ToList();
        }

        private static string _ReadSheetHeader(string file)
        {
            var reader = new StreamReader(file);
            string header = reader.ReadLine();
            reader.Close();
            return header ?? "";
        }
    }
}
