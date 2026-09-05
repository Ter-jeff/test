using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Static;

using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.Static;

namespace Automation.Utility
{
    public static class MergeSheet
    {
        public static void ParseTestSettingSheetToTestPlan(List<string> voltageTbFileName, ExcelWorkbook testPlanWorkbook)
        {
            HashSet<string> existJobs = new HashSet<string>();
            foreach (string table in voltageTbFileName)
            {
                string fileName = Path.GetFileNameWithoutExtension(table);
                string jobFromTable = fileName.Split('_')[4];
                string job = "";
                List<string> jobList = LocalSpecs.AllJobs;
                foreach (string jobName in jobList)
                {
                    if (jobName.Equals(jobFromTable, StringComparison.OrdinalIgnoreCase))
                    {
                        job = jobName;
                    }
                }
                if (string.IsNullOrEmpty(job))
                {
                    continue;
                }
                if (existJobs.Contains(job))
                {
                    throw new Exception(string.Format("There are multiple voltage table for {0}", job));
                }
                else
                {
                    existJobs.Add(job);
                }
                ExcelWorksheet updatedSheet = testPlanWorkbook.Worksheets
                    .Where(s => NeededSheets.IsTestSettingSheetName(s.Name, LocalSpecs.CurrentProject))
                    .ToList()
                    .Find(x => x.Name.Contains(job));
                if (updatedSheet != null)
                {
                    testPlanWorkbook.DeleteSheet(updatedSheet.Name);
                }

                if (Path.GetExtension(table) == ".csv")
                {
                    ExcelWorksheet sheet = testPlanWorkbook.Worksheets.Add("TestSetting_" + job);
                    int index = 0;
                    using (var sr = new StreamReader(table))
                    {
                        while (!sr.EndOfStream)
                        {
                            string line = sr.ReadLine();
                            if (line == "")
                            {
                                continue;
                            }

                            index++;
                            if (line != null)
                            {
                                string[] arr = line.Split(new[] { ',' }, StringSplitOptions.None);
                                int cnt = 0;
                                foreach (string item in arr)
                                {
                                    sheet.Cells[index, 1 + cnt].Value = item;
                                    cnt++;
                                }
                            }
                        }
                    }
                }
                else if (Path.GetExtension(table) == ".xlsx" || Path.GetExtension(table) == ".xlsm")
                {
                    bool found = false;
                    using (var excelPackage = new ExcelPackage(new FileInfo(table)))
                    {
                        foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                        {
                            if (worksheet.Name.ContainsIgnoreCase(job.ToUpper()))
                            {
                                testPlanWorkbook.Worksheets.Add("TestSetting_" + job, worksheet);
                                found = true;
                                break;
                            }
                        }
                        if (!found && excelPackage.Workbook.Worksheets.Count() == 1)
                        {
                            testPlanWorkbook.Worksheets.Add("TestSetting_" + job, excelPackage.Workbook.Worksheets.ElementAt(0));
                        }
                    }
                }
            }
        }
    }
}
