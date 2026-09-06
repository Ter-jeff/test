using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Basic
{
    public class JobInfoReader(List<string> allJobs, bool vreEnable = false) : MySheetReader<JobInfoSheet>
    {
        private const string ConJob = "Job";
        private const string ConType = "TesterType";
        private readonly Dictionary<string, int> _jobIdx = [];
        private int _indexJob = -1;
        private readonly JobInfoSheet _jobInfoSheet = new();
        private readonly List<string> _allJobs = allJobs;
        private readonly bool _vreEnable = vreEnable;
        private int _indexTypeRow = -1;

        public override JobInfoSheet? ReadSheet(ExcelWorksheet excelWorksheet)
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

            return _jobInfoSheet;
        }

        private void ReadSheet()
        {
            foreach (KeyValuePair<string, int> job in _jobIdx)
            {
                var jobInfoRow = new JobInfoRow { JobName = job.Key };
                int jobCol = job.Value;
                string type = ExcelWorksheet.GetCellValue(_indexTypeRow, jobCol).Trim();
                if (!string.IsNullOrEmpty(type))
                {
                    jobInfoRow.TesterType = _vreEnable ? "UFP" : type;
                }

                _jobInfoSheet.Rows.Add(jobInfoRow);
            }
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConJob))
                {
                    _indexJob = i;
                    continue;
                }
                if (_allJobs.Exists(x => x == header))
                {
                    _jobIdx.TryAdd(header, i);
                    continue;
                }
            }
            if (_indexJob != -1)
            {
                for (int i = StartRow; i <= EndRow; i++)
                {
                    string header = ExcelWorksheet.GetCellValue(i, _indexJob).Trim();
                    if (header.EqualsIgnoreCase(ConType))
                    {
                        _indexTypeRow = i;
                        continue;
                    }
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
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConJob))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }
    public class JobInfoRow
    {
        public string JobName = "";
        public string TesterType = "";
    }
    public class JobInfoSheet
    {
        public List<JobInfoRow> Rows = [];
    }
}
