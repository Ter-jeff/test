using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using LogLib.Utility;

using OfficeOpenXml;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    public class RegAssignGenerator
    {
        private const string SubBlockName = "SubBlockName";
        private const string InstanceName = "InstanceName";

        public RegAssignGenerator(HardIpInputData hardIpInputData)
        {
        }

        public void WorkFlow(List<HardIpRegAssign> hardIpRegAssigns, string outputFolder = "")
        {
            var groupByType = hardIpRegAssigns
                .GroupBy(p => p.IseFuseReadWriteType()
                    ? "_eFuse" : "_" + ((p.SubBlockName ?? "")
                        .Split(new[] { '_' }, 2, StringSplitOptions.RemoveEmptyEntries)
                        .FirstOrDefault() ?? ""))
                .ToDictionary(g => g.Key, g => g.ToList());

            foreach (KeyValuePair<string, List<HardIpRegAssign>> group in groupByType)
            {
                var dt = new DataTable();
                try
                {
                    int colIndex = 0;
                    int maxRowNum = group.Value.Max(p => p.RegAssignList.Count) + 3;
                    for (int i = 0; i < maxRowNum; i++)
                    {
                        dt.Rows.Add(dt.NewRow());
                    }

                    var rows = group.Value.OrderBy(x => x.SubBlockName).ToList();
                    GenerateItemsInDataTable(rows, dt, ref colIndex);
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }

                if (!Directory.Exists(FolderStructure.DirHardIp))
                {
                    Directory.CreateDirectory(FolderStructure.DirHardIp);
                }
                if (string.IsNullOrEmpty(outputFolder))
                {
                    outputFolder = FolderStructure.DirHardIp;
                }
                using (StreamWriter sw = File.CreateText(Path.Combine(outputFolder, NeededSheets.HardIRegAssign + group.Key + ".txt")))
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        sw.WriteLine(string.Join("\t", row.ItemArray));
                    }
                }
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirHardIp, NeededSheets.HardIRegAssign + group.Key);
            }
        }

        public void WorkFlow(ExcelWorkbook wb, List<HardIpRegAssign> hardIpRegAssigns)
        {
            var groupByType = hardIpRegAssigns
                .GroupBy(p => p.IseFuseReadWriteType()
                    ? "_eFuse" : "_" + ((p.SubBlockName ?? "")
                        .Split(new[] { '_' }, 2, StringSplitOptions.RemoveEmptyEntries)
                        .FirstOrDefault() ?? ""))
                .ToDictionary(g => g.Key, g => g.ToList());

            foreach (KeyValuePair<string, List<HardIpRegAssign>> group in groupByType)
            {
                ExcelWorksheet sheet = wb.Worksheets.Add(NeededSheets.HardIRegAssign + group.Key);
                var dt = new DataTable();
                try
                {
                    int colIndex = 0;
                    int maxRowNum = group.Value.Max(p => p.RegAssignList.Count) + 3;
                    for (int i = 0; i < maxRowNum; i++)
                    {
                        dt.Rows.Add(dt.NewRow());
                    }

                    var rows = group.Value.OrderBy(x => x.SubBlockName).ToList();
                    GenerateItemsInDataTable(rows, dt, ref colIndex);
                    sheet.Cells["A1"].LoadFromDataTable(dt, false);
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }

        }

        private void GenerateItemsInDataTable(List<HardIpRegAssign> hardIpRegAssigns, DataTable dt, ref int colIndex)
        {
            foreach (HardIpRegAssign regAssignInfo in hardIpRegAssigns)
            {
                if (!regAssignInfo.RegAssignList.Any())
                {
                    continue;
                }

                WriteData(regAssignInfo, dt, ref colIndex);

                WriteEmptyColumn(dt, ref colIndex);
            }
        }

        private static void WriteEmptyColumn(DataTable dt, ref int colIndex)
        {
            var column = new DataColumn(colIndex.ToString(), typeof(string));
            dt.Columns.Add(column);
            colIndex++;
        }

        private void WriteData(HardIpRegAssign hardIpRegAssign, DataTable dt, ref int colIndex)
        {
            SetRegAssignInfo(hardIpRegAssign, ref dt, ref colIndex);

            int maxColumn;
            if (hardIpRegAssign.Headers != null)
            {
                maxColumn = Math.Max(hardIpRegAssign.Headers.Count, hardIpRegAssign.RegAssignList.Max(x => x.Count));
            }
            else
            {
                maxColumn = hardIpRegAssign.RegAssignList.Max(x => x.Count);
            }

            int rowIndex = 1;
            for (int i = 0; i < maxColumn; i++)
            {
                var column = new DataColumn((colIndex + i).ToString(), typeof(string));
                string header = hardIpRegAssign.Headers == null || i >= hardIpRegAssign.Headers.Count ? "" : hardIpRegAssign.Headers.ElementAt(i);
                dt.Columns.Add(column);
                dt.Rows[rowIndex][colIndex + i] = header;
            }

            rowIndex = 2;
            for (int i = 0; i < hardIpRegAssign.RegAssignList.Count; i++)
            {
                List<string> data = hardIpRegAssign.RegAssignList.ElementAt(i);
                for (int j = 0; j < hardIpRegAssign.RegAssignList[i].Count; j++)
                {
                    try
                    {
                        dt.Rows[rowIndex + i][colIndex + j] = data.ElementAt(j);
                    }
                    catch (Exception)
                    {
                        //pass
                    }
                }
            }
            colIndex += maxColumn;
        }

        private static void SetRegAssignInfo(HardIpRegAssign regAssignInfo, ref DataTable dt, ref int colIndex)
        {
            var column = new DataColumn(colIndex.ToString(), typeof(string));
            dt.Columns.Add(column);
            dt.Rows[0][colIndex] = regAssignInfo.Type;
            dt.Rows[1][colIndex] = regAssignInfo.IseFuseReadWriteType() ? InstanceName : SubBlockName;
            dt.Rows[2][colIndex] = regAssignInfo.SubBlockName;
            colIndex++;
        }
    }
}
