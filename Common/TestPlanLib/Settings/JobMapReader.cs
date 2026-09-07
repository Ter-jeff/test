using System;
using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

using TestPlanLib.Basic;

namespace TestPlanLib.Settings
{
    public class JobMapReader : MySheetReader<JobMapSheet>
    {
        private int _startColIndex = 1;
        private int _startRowIndex = 1;
        private int _endColIndex = 1;
        private int _endRowIndex = 1;

        public override JobMapSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            try
            {
                ExcelWorksheet = excelWorksheet;
                InitIndex(excelWorksheet);
                JobMapSheet jobMapSheet = ReadData();
                return jobMapSheet;
            }
            catch (Exception e)
            {
                throw new Exception("Error in reading job mapping sheet! " + e.StackTrace);
            }
        }

        private void InitIndex(ExcelWorksheet excelWorksheet)
        {
            ExcelCellAddress startAddress = excelWorksheet.Dimension.Start;
            _startColIndex = startAddress.Column;
            _startRowIndex = startAddress.Row;
            ExcelCellAddress endAddress = excelWorksheet.Dimension.End;
            _endColIndex = endAddress.Column;
            _endRowIndex = endAddress.Row;
        }

        private JobMapSheet ReadData()
        {
            JobMapSheet jobMapSheet = new JobMapSheet();
            List<JobTemperatureMap> jobTempMap = [];
            Dictionary<string, List<string>> jobMapDictionary = [];
            for (int j = _startColIndex; j <= _endColIndex; j++)
            {
                int i = _startRowIndex;
                string testSetting = ExcelWorksheet.GetCellValue(i, j).Trim();
                if (testSetting.Length != 0)
                {
                    var jobList = new List<string>();
                    for (i++; i <= _endRowIndex; i++)
                    {
                        string jobName = ExcelWorksheet.GetCellValue(i, j);
                        if (jobName.Length != 0)
                        {
                            if (jobName.Split(':').Length > 1)
                            {
                                string temp = jobName.Split(':')[1];
                                string job = jobName.Split(':')[0];
                                jobList.Add(job);
                                jobTempMap.Add(new JobTemperatureMap(job, temp, testSetting));
                            }
                            else
                            {
                                jobList.Add(jobName);
                            }
                        }
                    }
                    jobMapDictionary.Add(testSetting, jobList);
                }
            }

            jobMapSheet.JobMapDictionary = jobMapDictionary;
            jobMapSheet.JobTempMap = jobTempMap;
            return jobMapSheet;
        }
    }

    public class JobMapSheet
    {
        public Dictionary<string, List<string>> JobMapDictionary { get; set; } = [];
        public List<JobTemperatureMap> JobTempMap { get; set; } = [];
    }
}
