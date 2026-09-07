using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Static;

namespace Automation.GenerateIgxl.PostAction.VersionTracing
{
    public static class OutputHashTable
    {
        public static List<string[]> CreateOutputHashTable()
        {
            var table = new List<string[]>
            {
                new[] { "Output Filepath", "Output File Type", "Checksum" }
            };
            Dictionary<string, string> sheets = TestProgram.IgxlWorkBk.AllIgxlSheetsDicWithType;
            sheets.Concat(TestProgram.T0TxIgxlWorkBk.AllIgxlSheetsDicWithType);
            sheets.Concat(TestProgram.SubProgIgxlWorkBk.AllIgxlSheetsDicWithType);
            foreach (KeyValuePair<string, string> sheet in sheets)
            {
                string outputPath = sheet.Key.Replace(FolderStructure.DirIgLink + Path.DirectorySeparatorChar, "");
                string outputFileType = sheet.Value;
                string checksum = HashCalculator.CalculateMd5(sheet.Key + ".txt");
                table.Add(new[] { outputPath, outputFileType, checksum });
            }
            return table;
        }
    }
}
