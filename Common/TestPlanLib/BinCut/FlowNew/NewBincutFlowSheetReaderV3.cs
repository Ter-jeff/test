using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;
using CommonLib.Utility;

using CommonReaderLib;

using OfficeOpenXml;

using TestPlanLib.BinCut.Flow;

namespace TestPlanLib.BinCut.FlowNew
{
    public class NewBincutFlowSheetReaderV3 : MySheetReader<NewBinCutFlowTables>
    {
        private const string ConHeaderBincutFlowVersion1 = "BinCutFlowVersion1.0";
        private const string ConHeaderBiningDomain = "Binning.*";
        private const string ConHeaderTableType = "Type";
        private const string ConHeaderPerformanceMode = "Performance.*";
        private const string ConHeaderSubFlow = "SubFlow";
        private const string ConHeaderComment = "Comment";
        private const string ConEnable = "Enable";
        private const int ConMaxSearchColumn = 5;
        private const int ConMaxSearchRow = 10;

        private bool _version;
        private int _binningIdx;
        private int _tableTypeIdx;
        private int _performanceIdx;
        private int _subFlowIdx;
        private Dictionary<string, int> _jobStageIdx = [];
        private int _enableIdx;
        public Dictionary<string, int> SubFlowIdxTable = [];
        public string BincutType { get; private set; } = "";

        public NewBincutFlowSheetReaderV3(ref bool version)
        {
            _version = version;
        }

        public override NewBinCutFlowTables? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var outputTables = new NewBinCutFlowTables();
            ExcelWorksheet = excelWorksheet;
            EndCol = ExcelWorksheet.Dimension.End.Column;
            EndRow = ExcelWorksheet.Dimension.End.Row;
            BincutType = excelWorksheet.Name;
            _version = ReadHeader();
            if (!_version)
            {
                return outputTables;
            }

            foreach (KeyValuePair<string, int> job in _jobStageIdx)
            {
                NewBinCutFlowTable sheet = ReadSheet(job.Key, job.Value);
                if (sheet.Rows.Count != 0)
                {
                    NewBinCutFlowTable.Check();
                    sheet.SubFlowMappingTable = SubFlowIdxTable;
                    outputTables.Add(sheet);
                }
            }

            return outputTables;
        }

        public List<string> ReadJobCount()
        {
            bool version = ReadHeader();
            if (!version)
            {
                return [];
            }

            return [.. _jobStageIdx.Select(x => x.Key)];
        }

        private bool ReadHeader()
        {
            _jobStageIdx = [];
            bool foundHeader = false;
            for (int i = 1; i < ConMaxSearchColumn; i++)
            {
                if (foundHeader)
                {
                    break;
                }
                for (int j = 1; j < ConMaxSearchRow; j++)
                {
                    string value = ExcelWorksheet.GetCellValue(j, i);
                    if (CellDiff.IsLiked(value, ConHeaderBincutFlowVersion1))
                    {
                        return false;
                    }
                    if (CellDiff.IsLiked(value, ConHeaderTableType))
                    {
                        StartCol = i;
                        StartRow = j;
                        foundHeader = true;
                        break;
                    }
                    if (CellDiff.IsLiked(value, ConHeaderBiningDomain))
                    {
                        StartCol = i;
                        StartRow = j;
                        foundHeader = true;
                        break;
                    }
                    if (CellDiff.IsLiked(value, ConHeaderPerformanceMode))
                    {
                        StartCol = i;
                        StartRow = j;
                        foundHeader = true;
                        break;
                    }
                }
            }
            for (int i = StartCol; i <= EndCol; i++)
            {
                string value = ExcelWorksheet.GetCellValue(StartRow, i);
                if (CellDiff.IsLiked(value, ConHeaderTableType))
                {
                    _tableTypeIdx = i;
                }
                else if (CellDiff.IsLiked(value, ConHeaderBiningDomain))
                {
                    _binningIdx = i;
                }
                else if (CellDiff.IsLiked(value, ConHeaderPerformanceMode))
                {
                    _performanceIdx = i;
                }
                else if (CellDiff.IsLiked(value, ConHeaderSubFlow))
                {
                    _subFlowIdx = i;
                }
                else if (CellDiff.IsLiked(value, ConEnable))
                {
                    _enableIdx = i;
                }
                else if (CellDiff.IsLiked(value, ConHeaderComment))
                {
                    break;
                }
                else
                {
                    _jobStageIdx.Add(value, i);
                }
            }
            return true;
        }

        private NewBinCutFlowTable ReadSheet(string job, int jobCol)
        {
            var bincutFlowTable = new NewBinCutFlowTable
            {
                SheetName = ExcelWorksheet.Name,
                BinningDomainIndex = _binningIdx,
                PerformanceModeIndex = _performanceIdx
            };
            string mergeCell = ExcelWorksheet.MergedCells[StartRow, StartCol];
            if (mergeCell != null)
            {
                bincutFlowTable.StartRowIndex = new ExcelAddress(ExcelWorksheet.MergedCells[StartRow, StartCol]).End.Row;
            }
            else
            {
                bincutFlowTable.StartRowIndex = StartRow;
            }

            bincutFlowTable.JobName = job;
            bincutFlowTable.FinalJob = new Job(job.Split(',').First()).JobType.ToString();

            EnumBinCutTableType tableType = EnumBinCutTableType.Lv;
            string performanceModePreviousRow = "";
            int startRow = StartRow + 1;
            var row = new NewBinCutFlowSheetRow(ExcelWorksheet.Name, bincutFlowTable.FinalJob);
            var subFlows = new List<string>();
            var subFlowsbyBms = new List<(string, bool)>();
            SubFlowIdxTable ??= [];

            for (int i = startRow; i <= EndRow; i++)
            {
                string performanceModeNextRow = ExcelWorksheet.GetCellValue(i, _performanceIdx).Trim();
                performanceModePreviousRow = string.IsNullOrEmpty(performanceModePreviousRow) ? performanceModeNextRow : performanceModePreviousRow;
                if (performanceModePreviousRow != performanceModeNextRow)
                {
                    performanceModePreviousRow = performanceModeNextRow;
                    if (subFlows.Count != 0)
                    {
                        row.SubFlows = [.. subFlows.Where(x => !string.IsNullOrEmpty(x))];
                        row.SubFlowsByType = [.. subFlowsbyBms];
                        bincutFlowTable.Rows.Add(row);
                    }
                    row = new NewBinCutFlowSheetRow(ExcelWorksheet.Name, bincutFlowTable.FinalJob);
                    subFlows = [];
                    subFlowsbyBms = [];
                }

                row.RowNum = i;
                row.BinningDomain = ExcelWorksheet.GetCellValue(i, _binningIdx).Trim();
                row.PerformanceMode = ExcelWorksheet.GetCellValue(i, _performanceIdx).Trim().Split('_')[0];
                row.Enable = ExcelWorksheet.GetCellValue(i, _enableIdx).Trim();
                string tableTypeStr = ExcelWorksheet.GetCellValue(i, _tableTypeIdx).Trim();
                if (tableTypeStr == "HBV")
                {
                    tableType = EnumBinCutTableType.Hv;
                }
                else if (tableTypeStr == "BV")
                {
                    tableType = EnumBinCutTableType.Lv;
                }
                else if (tableTypeStr == "PBC")
                {
                    tableType = EnumBinCutTableType.Post;
                }

                row.TableType = tableType;
                string jobValue = ExcelWorksheet.GetCellValue(i, jobCol);

                string value = ExcelWorksheet.GetCellValue(i, _performanceIdx).Trim() + ":" + ExcelWorksheet.GetCellValue(i, _subFlowIdx).Trim();
                if (!NewBincutFlowSheetRegex.Performance().IsMatch(value) && !value.StartsWithIgnoreCase("Flow_TMPS"))
                {
                    value = row.PerformanceMode + "#" + value;
                }

                if (ExcelWorksheet.GetCellValue(i, jobCol).Contains("F_"))
                {
                    value = value + " " + ExcelWorksheet.GetCellValue(i, jobCol).Split(["(", ")"], StringSplitOptions.RemoveEmptyEntries).Last();
                }

                if (!string.IsNullOrEmpty(value) && !SubFlowIdxTable.ContainsKey(value))
                {
                    if (value.Contains("HIP", StringComparison.OrdinalIgnoreCase))
                    {
                        string valueIlb = NewBincutFlowSheetRegex.Hip().Replace(value, "ILB");
                        SubFlowIdxTable.TryAdd(valueIlb, i);
                        string valueElb = NewBincutFlowSheetRegex.Hip().Replace(value, "ELB");
                        SubFlowIdxTable.TryAdd(valueElb, i);
                    }
                    else
                    {
                        SubFlowIdxTable.Add(value, i);
                    }
                }
                if (string.IsNullOrEmpty(jobValue) || !jobValue.EqualsIgnoreCase("TRUE"))
                {
                    continue;
                }

                if (string.IsNullOrEmpty(value))
                {
                    continue;
                }

                if (value.Contains("HIP", StringComparison.OrdinalIgnoreCase))
                {
                    string valueIlb = NewBincutFlowSheetRegex.Hip().Replace(value, "ILB");
                    subFlows.Add(valueIlb);
                    subFlowsbyBms.Add((valueIlb, true));
                    string valueElb = NewBincutFlowSheetRegex.Hip().Replace(value, "ELB");
                    subFlows.Add(valueElb);
                    subFlowsbyBms.Add((valueElb, true));
                }
                else
                {
                    subFlows.Add(value);
                    subFlowsbyBms.Add((value, false));
                }
            }
            if (subFlows.Count != 0)
            {
                row.SubFlows = [.. subFlows.Where(x => !string.IsNullOrEmpty(x))];
                row.SubFlowsByType = [.. subFlowsbyBms];
                bincutFlowTable.Rows.Add(row);
            }
            return bincutFlowTable;
        }
    }
}
