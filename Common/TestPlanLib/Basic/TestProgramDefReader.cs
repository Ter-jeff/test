using System.Collections.Generic;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace TestPlanLib.Basic
{
    public class TestProgramDefReader : MySheetReader<TestProgramDefSheet>
    {
        private const string ConTestprgramName = "TestProgram Name";
        private readonly Dictionary<string, int> _jobIdx = [];
        private int _indexTestProgramName = -1;
        private readonly string _regJob = @"(CP|WLFT|FT|FQA|RMA|EMA)[\d]+";
        private readonly TestProgramDefSheet _testprogramDefSheet = new();

        public override TestProgramDefSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            ReadSheet();

            return _testprogramDefSheet;
        }

        private void ReadSheet()
        {
            var allJobs = new List<string>();
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new TestProgramRow();
                if (_indexTestProgramName != -1)
                {
                    row.ProgramName = ExcelWorksheet.GetCellValue(i, _indexTestProgramName).Trim();
                }
                foreach (KeyValuePair<string, int> jobIdx in _jobIdx)
                {
                    string value = ExcelWorksheet.GetCellValue(i, jobIdx.Value).Trim().ToUpper();
                    string mappingJob = jobIdx.Key.ToUpper();
                    if (string.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    if (!allJobs.Exists(x => x == mappingJob))
                    {
                        allJobs.Add(mappingJob);
                    }

                    row.JobMapping.Add(mappingJob, value);
                }
                if (row.JobMapping.Count != 0)
                {
                    _testprogramDefSheet.Rows.Add(row);
                }
            }
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (Regex.IsMatch(header, _regJob))
                {
                    _jobIdx.Add(header, i);
                    continue;
                }
                if (header.EqualsIgnoreCase(ConTestprgramName))
                {
                    _indexTestProgramName = i;
                    continue;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow;
            int colNum = EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConTestprgramName))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }
    public class TestProgramDefSheet
    {
        public List<TestProgramRow> Rows = [];
        private readonly List<string> _allJobs = [];
        private Dictionary<string, string>? _jobMappingDic;
        public Dictionary<string, string> JobMappingDic
        {
            get
            {
                if (_jobMappingDic == null)
                {
                    _jobMappingDic = [];
                    foreach (TestProgramRow row in Rows)
                    {
                        foreach (KeyValuePair<string, string> item in row.JobMapping)
                        {
                            string realJob = item.Key;
                            string job = item.Value;
                            _jobMappingDic.TryAdd(job, realJob);
                        }
                    }
                }
                return _jobMappingDic;
            }
        }
        public List<string> AllJobs
        {
            get
            {
                if (_allJobs.Count == 0)
                {
                    foreach (TestProgramRow row in Rows)
                    {
                        foreach (KeyValuePair<string, string> job in row.JobMapping)
                        {
                            if (string.IsNullOrEmpty(job.Value))
                            {
                                continue;
                            }

                            if (!_allJobs.Contains(job.Key))
                            {
                                _allJobs.Add(job.Key);
                            }
                        }
                    }
                }
                return _allJobs;
            }
        }
    }

    public class TestProgramRow
    {
        public string ProgramName = "";
        public Dictionary<string, string> JobMapping = [];
        public JobListSheet? JobListSheet { get; set; }
    }
}
