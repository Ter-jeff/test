using System.Collections.Generic;
using System.IO;

using Automation.Static;

using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.CPM;

namespace Automation.GenerateIgxl.Scan.CPM
{
    internal class CpmTable
    {
        public void WorkFlow(EfuseCpmSheet efuseCpm)
        {
            var cpmTable = new ExcelPackage(new FileInfo(Path.Combine(FolderStructure.DirCpm, "CPM_Table.xlsx")));
            ExcelWorksheet workSheet = cpmTable.Workbook.AddSheet("CPM_Table");
            var headers = new List<string> { "Flag1", "Flag2" };
            var efuseNames = new List<string>();
            var flagStatus1 = new List<string> { "True", "False" };
            var flagStatus2 = new List<string> { "False", "True" };
            var flags = new List<string>();
            foreach (KeyValuePair<string, int> flag in efuseCpm.FlagColNumber)
            {
                efuseNames.Add(flag.Key);
                flags.Add(flag.Key);
            }
            foreach (EfuseCpmSheetRow row in efuseCpm.Rows)
            {
                headers.Add(row.EfuseBank.Replace("bank_", ""));
                efuseNames.Add(row.CpmEfuseName2);
                flagStatus1.Add(row.FlagsValue[flags[0]]);
                flagStatus2.Add(row.FlagsValue[flags[1]]);
            }
            int col = 1;
            foreach (string header in headers)
            {
                workSheet.Cells[1, col].Value = header;
                col++;
            }
            col = 1;
            foreach (string efuseName in efuseNames)
            {
                workSheet.Cells[2, col].Value = efuseName;
                col++;
            }
            col = 1;
            foreach (string flag in flagStatus1)
            {
                workSheet.Cells[3, col].Value = flag;
                col++;
            }
            col = 1;
            foreach (string flag in flagStatus2)
            {
                workSheet.Cells[4, col].Value = flag;
                col++;
            }

            cpmTable.ExportWorkBook2Txt(FolderStructure.DirCpm);
            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCpm, "CPM_Table");
        }
    }
}
