using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;
using CommonLib.Utility;

using CommonReaderLib;

using OfficeOpenXml;

using TestPlanLib.BinCut.Flow;

namespace TestPlanLib.BinCut.FlowNew
{
    public partial class NewBincutFlowSheetReaderV2 : MySheetReader<NewBinCutFlowTables>
    {
        private const string ConHeaderBincutFlowVersion1 = "BinCutFlowVersion1.0";
        private const string ConHeaderBiningDomain = "Binning.*";
        private const string ConHeaderPerformanceMode = "Performance.*";
        private const string ConHeaderSubFlow = "SubFlow";
        private const string ConHeaderComment = "Comment";
        private const string ConHeaderEnable = "Enable";
        private const int ConMaxSearchColumn = 5;
        private const int ConMaxSearchRow = 10;

        [GeneratedRegex(BinCutFlowTable.RegexPerformance, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        private bool _isMatchVersion;
        private int _binningIdx;
        private int _enableIdx;
        private int _performanceIdx;
        private int _subFlowIdx;
        private Dictionary<string, int> _jobStageIdx = [];
        public Dictionary<string, int> SubFlowIdxTable = [];

        public NewBincutFlowSheetReaderV2(ref bool version)
        {
            _isMatchVersion = version;
        }

        public override NewBinCutFlowTables? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var outputTables = new NewBinCutFlowTables();
            ExcelWorksheet = excelWorksheet;

            GetDimensions();

            _isMatchVersion = ReadHeader();
            if (!_isMatchVersion)
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

        private bool ReadHeader()
        {
            _jobStageIdx = [];
            bool foundHeader = FindStartPostion();
            if (!foundHeader)
            {
                return false;
            }
            for (int i = StartCol; i <= EndCol; i++)
            {
                string value = ExcelWorksheet.GetCellValue(StartRow, i);
                if (CellDiff.IsLiked(value, ConHeaderBiningDomain))
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
                else if (CellDiff.IsLiked(value, ConHeaderEnable))
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

        private bool FindStartPostion()
        {
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

            return foundHeader;
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

            int startRow = StartRow + 1;
            var row = new NewBinCutFlowSheetRow(ExcelWorksheet.Name, bincutFlowTable.FinalJob);
            var subFlows = new List<string>();
            SubFlowIdxTable ??= [];

            ReadData(jobCol, bincutFlowTable, startRow, ref row, ref subFlows);

            if (subFlows.Count != 0)
            {
                row.SubFlows = [.. subFlows.Where(x => !string.IsNullOrEmpty(x))];
                bincutFlowTable.Rows.Add(row);
            }
            return bincutFlowTable;
        }

        private void ReadData(int jobCol, NewBinCutFlowTable newBinCutFlowTable, int startRow, ref NewBinCutFlowSheetRow newBinCutFlowSheetRow, ref List<string> subFlows)
        {
            EnumBinCutTableType tableType = Flow.EnumBinCutTableType.Lv;
            EnumBinCutTableBinType tableBinType = EnumBinCutTableBinType.Bin1;
            string performanceModePreviousRow = "";
            for (int i = startRow; i <= EndRow; i++)
            {
                string performanceModeNextRow = ExcelWorksheet.GetCellValue(i, _performanceIdx).Trim();
                performanceModePreviousRow = string.IsNullOrEmpty(performanceModePreviousRow) ? performanceModeNextRow : performanceModePreviousRow;
                string subFlowName = ExcelWorksheet.GetCellValue(i, _subFlowIdx).Trim();

                if (performanceModePreviousRow != performanceModeNextRow)
                {
                    performanceModePreviousRow = performanceModeNextRow;
                    if (subFlows.Count != 0)
                    {
                        newBinCutFlowSheetRow.SubFlows = [.. subFlows.Where(x => !string.IsNullOrEmpty(x))];
                        newBinCutFlowTable.Rows.Add(newBinCutFlowSheetRow);
                    }
                    newBinCutFlowSheetRow = new NewBinCutFlowSheetRow(ExcelWorksheet.Name, newBinCutFlowTable.FinalJob);
                    subFlows = [];
                }

                newBinCutFlowSheetRow.RowNum = i;
                newBinCutFlowSheetRow.BinningDomain = ExcelWorksheet.GetCellValue(i, _binningIdx).Trim();
                newBinCutFlowSheetRow.PerformanceMode = ExcelWorksheet.GetCellValue(i, _performanceIdx).Trim();
                newBinCutFlowSheetRow.Enable = ExcelWorksheet.GetCellValue(i, _enableIdx).Trim();
                if (_binningIdx != 0 && !string.IsNullOrEmpty(newBinCutFlowSheetRow.BinningDomain))
                {
                    if (newBinCutFlowSheetRow.BinningDomain.StartsWithIgnoreCase("End"))
                    {
                        continue;
                    }

                    EnumBinCutTableType(newBinCutFlowSheetRow.BinningDomain, ref tableType, ref tableBinType);

                    newBinCutFlowSheetRow.TableType = tableType;
                    newBinCutFlowSheetRow.TableBinType = tableBinType;
                }
                else
                {
                    if (newBinCutFlowSheetRow.PerformanceMode.StartsWithIgnoreCase("End"))
                    {
                        continue;
                    }

                    EnumBinCutTableType(newBinCutFlowSheetRow.PerformanceMode, ref tableType, ref tableBinType);

                    newBinCutFlowSheetRow.TableType = tableType;
                    newBinCutFlowSheetRow.TableBinType = tableBinType;
                }
                if (string.IsNullOrEmpty(ExcelWorksheet.GetCellValue(i, jobCol)))
                {
                    continue;
                }

                string value = ExcelWorksheet.GetCellValue(i, _subFlowIdx).Trim();
                if (string.IsNullOrEmpty(value))
                {
                    continue;
                }

                if (!_regex.IsMatch(value) && !value.StartsWithIgnoreCase("Flow_TMPS"))
                {
                    value = newBinCutFlowSheetRow.PerformanceMode + "#" + value;
                }

                if (ExcelWorksheet.GetCellValue(i, jobCol).Contains("F_"))
                {
                    value = value + " " + ExcelWorksheet.GetCellValue(i, jobCol).Split(["(", ")"], StringSplitOptions.RemoveEmptyEntries).Last();
                }

                if (!string.IsNullOrEmpty(value) && !SubFlowIdxTable.ContainsKey(value))
                {
                    SubFlowIdxTable.Add(value, i);
                }
                subFlows.Add(value);
            }
        }

        private static void EnumBinCutTableType(string key, ref EnumBinCutTableType enumBinCutTableType, ref EnumBinCutTableBinType enumBinCutTableBinType)
        {
            if (key.Contains("HVCC", StringComparison.OrdinalIgnoreCase) || key.Contains("HBV", StringComparison.OrdinalIgnoreCase))
            {
                enumBinCutTableType = Flow.EnumBinCutTableType.Hv;
            }

            if (key.Contains("BIN1", StringComparison.OrdinalIgnoreCase))
            {
                enumBinCutTableBinType = EnumBinCutTableBinType.Bin1;
            }
            else if (key.Contains("BINX", StringComparison.OrdinalIgnoreCase))
            {
                enumBinCutTableBinType = EnumBinCutTableBinType.BinX;
            }
            else if (key.Contains("BINY", StringComparison.OrdinalIgnoreCase))
            {
                enumBinCutTableBinType = EnumBinCutTableBinType.BinY;
            }

            if (key.Contains("PBC", StringComparison.OrdinalIgnoreCase))
            {
                enumBinCutTableType = Flow.EnumBinCutTableType.Post;
            }
        }
    }
}
