using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.common.IgxlDataExtension;

using IgxlLib.Enums;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

namespace Cautogen.AutoCZ.CharPostProcessor
{
    public class GenBas
    {
        public void WorkFlow(string execEnableWord, string progWorkBookPath, string newEnableWord)
        {
            List<string> execEnableWords = execEnableWord.Split(',').ToList();
            execEnableWords.AddRange(newEnableWord.Split(','));
            string fileName = Path.Combine(Path.Combine(LocalSpecs.OutputFolder, ConstData.CzFolder), "VBT_LIB_PV.bas");
            List<string> totalEnableWords = GetEnables(progWorkBookPath);
            totalEnableWords.AddRange(newEnableWord.Split(','));

            var basSheet = new BasFile("VBT_LIB_PV.bas");
            var contentLines = new List<string>();
            contentLines.Add("Attribute VB_Name = \"" + Path.GetFileNameWithoutExtension(fileName) + "\"");
            contentLines.Add(SetEnableWords(execEnableWords, totalEnableWords));
            contentLines.Add(PrintEnableWords(totalEnableWords));
            basSheet.Write(fileName);
            LocalSpecs.GenSheets.Add(basSheet);
        }

        public static List<string> GetEnables(string text)
        {
            var enables = new List<string>();
            if (!File.Exists(text))
            {
                return enables;
            }

            using (ZipArchive archive = ZipFile.OpenRead(text))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    string sheetName = Path.GetFileNameWithoutExtension(entry.FullName);
                    if (sheetName != null)
                    {
                        string firstLine = "";
                        using (Stream stream = entry.Open())
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            firstLine = reader.ReadLine();
                        }

                        if (firstLine != null)
                        {
                            EnumSheetType type = IgxlLoaderHelpers.GetIgxlSheetType(firstLine);
                            if (type == EnumSheetType.DTFlowtableSheet)
                            {
                                using (Stream processingStream = entry.Open())
                                {
                                    enables.AddRange(new ReadFlowSheet().GetEnables(processingStream, sheetName));
                                }
                            }
                        }
                    }
                }
            }
            enables = enables.SelectMany(x => Regex.Split(x, @"\&|\||!|\(|\)")).Select(x => x.Trim()).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList();
            return enables.OrderBy(x => x).ToList();
        }

        private string SetEnableWords(List<string> execEnableWords, List<string> totalEnableWords)
        {
            string codeText = "Public Sub SetEnableWords()" + "\r\n";

            foreach (string enableword in totalEnableWords)
            {
                bool flag = execEnableWords.Exists(
                        x => x.Equals(enableword, StringComparison.CurrentCultureIgnoreCase));
                codeText += $"  tl_ExecSetEnableWord \"{enableword}\" ,{flag}" + "\r\n";
            }
            codeText += "End Sub\r\n";
            return codeText;
        }

        private string PrintEnableWords(List<string> totalEnableWords)
        {
            string codeText = "Public Sub PrintEnableWords()" + "\r\n";
            foreach (string enableword in totalEnableWords)
            {
                codeText += string.Format("  If (tl_ExecGetEnableWord(\"{0}\")) Then TheExec.Datalog.WriteComment \"{0}:\" + CStr(tl_ExecGetEnableWord(\"{0}\"))", enableword) + "\r\n";
            }

            codeText += "End Sub\r\n";
            return codeText;
        }

    }
}
