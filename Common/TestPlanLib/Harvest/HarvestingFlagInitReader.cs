using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Harvest
{
    public class HarvestingFlagInitReader : MySheetReader<HarvestingFlagInitSheet>
    {
        private const string ConHeaderSetFlag = "Set Flag";
        private List<string> _allJobs;
        private Dictionary<string, int> _jobHeaderDic = new(StringExtensions.IgnoreCase);
        private int _indexSetFlag = -1;
        public HarvestingFlagInitReader(List<string> allJobs)
        {
            _allJobs = allJobs;
        }
        public override HarvestingFlagInitSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            HarvestingFlagInitSheet harvestingFlagInitSheet = ReadSheet(sheetName);
            harvestingFlagInitSheet.CreateMappingListByJob();

            return harvestingFlagInitSheet;
        }

        private HarvestingFlagInitSheet ReadSheet(string sheetName)
        {
            var harvestingFlagInitSheet = new HarvestingFlagInitSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new HarvestingFlagInitRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexSetFlag != -1)
                {
                    row.SetFlag = ExcelWorksheet.GetCellValue(i, _indexSetFlag).Trim();
                }
                row.JobFuse = _jobHeaderDic.ToDictionary(job => job.Key, job => ExcelWorksheet.GetCellValue(i, job.Value).Trim());

                if (!string.IsNullOrEmpty(row.SetFlag))
                {
                    harvestingFlagInitSheet.Rows.Add(row);
                }
            }
            return harvestingFlagInitSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderSetFlag))
                {
                    _indexSetFlag = i;
                    continue;
                }
                string? matchJob = _allJobs.FirstOrDefault(header.StartsWithIgnoreCase);
                if (matchJob != null)
                {
                    _jobHeaderDic.Add(matchJob, i);
                    continue;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderSetFlag))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public partial class HarvestingFlagInitSheet : MySheet
    {
        private const string Regex1 = "(?<flag>.+)";
        private const string Regex2 = @"(?<cores>\[.+\])";

        [GeneratedRegex(Regex1, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(Regex2, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();
        private static readonly Regex _regex2 = MyRegex();
        private static readonly Regex _regex4 = MyRegex1();

        public List<HarvestingFlagInitRow> Rows { set; get; }
        public List<HarvestingFuseRead> EfuseMappingTable { set; get; } = [];
        public Dictionary<string, string> BankDictionary = [];
        #region Constructor

        public HarvestingFlagInitSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }

        public void CreateMappingListByJob()
        {
            EfuseMappingTable = [];
            var allJobs = Rows.SelectMany(x => x.JobFuse.Keys).Select(x => x).Distinct().ToList();
            foreach (string job in allJobs)
            {
                EfuseMappingTable.Add(new HarvestingFuseRead { Job = job, FuseMapping = Rows.ToDictionary(x => x.SetFlag, x => x.JobFuse[job]) });
            }
        }

        public List<string> GetFlags()
        {
            var flagList = new List<string>();
            foreach (string flagName in Rows.Select(x => x.SetFlag))
            {
                if (_regex2.IsMatch(flagName))
                {
                    string flag = _regex2.Match(flagName).Groups["flag"].ToString();
                    string flagRemoveCode = _regex4.Replace(flag, "");
                    string[] sumflagItem = flag.Split([',', ' '], StringSplitOptions.RemoveEmptyEntries);
                    string sumFlagName = (flag.Replace(",", "_").Replace("[", "_").Replace("]", "_").Replace(":", "_").Replace("..", "_") + "_SUM").Replace("__", "_");
                    GroupItem iterates = flag.GetIterates();
                    foreach (KeyValuePair<int, List<Tuple<string, string>>> iterate in iterates.Groups)
                    {
                        string text = flag;
                        foreach (Tuple<string, string> item in iterate.Value)
                        {
                            text = text.Replace(item.Item1, item.Item2);
                        }

                        flagList.Add(text);
                    }
                    if (iterates.Groups.Count == 0)
                    {
                        flagList.Add(flagName);
                    }
                }
            }

            return flagList;
        }
        #endregion
    }
    public class HarvestingFlagInitRow : MyRow
    {
        public string SetFlag = "";
        public Dictionary<string, string> JobFuse = new(StringExtensions.IgnoreCase);
        #region Constructor
        public HarvestingFlagInitRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion
    }

    public class HarvestingFuseRead
    {
        public string Job = "";
        public Dictionary<string, string> FuseMapping = [];
    }
}
