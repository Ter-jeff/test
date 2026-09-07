using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Static;

namespace Automation.GenerateIgxl.PostAction.VersionTracing
{
    public static class VersionTracingMain
    {
        private const string ExecInfoSheetName = "ExecInfo";
        private const string VersionsSheetName = "Versions";
        private const string OutputHashesSheetName = "output_hashes";

        public static void WorkFlow()
        {
            WriteExecInfoTable();
            WriteOutputHashTable();
            WriteVersionsTable();
        }

        private static void WriteExecInfoTable()
        {
            List<string[]> execInfoTable = ExecInfoTable.CreateExecInfoTable();
            WriteSheet(ExecInfoSheetName, execInfoTable);
        }

        private static void WriteOutputHashTable()
        {
            List<string[]> outputHashTable = OutputHashTable.CreateOutputHashTable();
            WriteSheet(OutputHashesSheetName, outputHashTable);
        }

        private static void WriteVersionsTable()
        {
            List<string[]> versionsTable = VersionsTable.CreateVersionTable();
            WriteSheet(VersionsSheetName, versionsTable);
        }

        private static void WriteSheet(string sheetName, IEnumerable<string[]> table)
        {
            // Convert table to TSV
            string[] lines = table.Select(row => string.Join("\t", row)).ToArray();

            // Write to file
            string filePath = Path.Combine(FolderStructure.DirReference, sheetName + ".txt");
            if (!Directory.Exists(FolderStructure.DirReference))
            {
                Directory.CreateDirectory(FolderStructure.DirReference);
            }
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            File.WriteAllLines(filePath, lines.ToArray());
            LocalSpecs.SelectedFiles.Add(filePath);

            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirReference, sheetName);
        }
    }
}
