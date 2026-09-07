using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;
using CommonLib.Utility;

using CommonReaderLib;

using IgxlLib.Enums;

using OfficeOpenXml;

using TestPlanLib.BinCut.Flow;
using TestPlanLib.Utility;

namespace TestPlanLib.BinCut.FlowNew
{
    public partial class BinCutFlowSheetReaderNew(bool isT0Tx = false) : MySheetReader<NewBinCutFlowTables>
    {

        private const string ConHeaderBiningDomain = "Binning.*";
        private const string ConHeaderPerformanceMode = "Performance.*";
        private const string ConHeaderSubFlow = @"SubFlow\d*";

        private const int ConMaxSearchColumn = 5;
        private const int ConMaxSearchRow = 10;

        [GeneratedRegex(ConHeaderSubFlow, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(BinCutFlowTable.RegexPerformance, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();

        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();

        private readonly bool _isT0Tx = isT0Tx;
        private int _conHeaderBiningDomainIndex = -1;
        private int _conHeaderPerformanceModeIndex = -1;

        public override NewBinCutFlowTables? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var outputTables = new NewBinCutFlowTables();

            ExcelWorksheet = excelWorksheet;

            EndCol = ExcelWorksheet.Dimension.End.Column;
            EndRow = ExcelWorksheet.Dimension.End.Row;

            List<SubBinCutSheetInfo> subBinCutSheetInfos = ReadSubSheetInfo();

            foreach (SubBinCutSheetInfo subBinCutSheetInfo in subBinCutSheetInfos)
            {
                if (!_isT0Tx && new Job(subBinCutSheetInfo.JobName.Split(',').First()).JobType == EnumJob.None)
                {
                    continue;
                }

                ReadHeader(subBinCutSheetInfo);
                NewBinCutFlowTable sheet = ReadSheet(subBinCutSheetInfo);
                if (sheet.Rows.Count != 0)
                {
                    NewBinCutFlowTable.Check();
                    outputTables.Add(sheet);
                }
            }
            return outputTables;
        }

        private List<SubBinCutSheetInfo> ReadSubSheetInfo()
        {
            for (int i = 1; i < ConMaxSearchColumn; i++)
            {
                for (int j = 1; j < ConMaxSearchRow; j++)
                {
                    string value = ExcelWorksheet.GetCellValue(j, i);
                    if (CellDiff.IsLiked(value, ConHeaderBiningDomain))
                    {
                        StartCol = i;
                        StartRow = j;
                        break;
                    }
                }
            }

            var subBinCutSheetInfos = new List<SubBinCutSheetInfo>();
            var subBinCutSheetInfo = new SubBinCutSheetInfo
            {
                StartColumnNum = StartCol
            };
            bool firstFlag = true;
            int subFlowendColumn = 0;
            int endColumn = 0;
            for (int i = StartCol + 1; i <= EndCol; i++)
            {
                string value = ExcelWorksheet.GetCellValue(StartRow, i);
                string value2 = ExcelWorksheet.GetCellValue(StartRow + 1, i);

                endColumn = i;
                //Read Job Name
                if (_regex.IsMatch(value2))
                {
                    subFlowendColumn = i;
                    if (firstFlag)
                    {
                        subBinCutSheetInfo.JobName = value;
                        subBinCutSheetInfo.JobColumnNum = i;
                        firstFlag = false;
                    }
                }
                else if (CellDiff.IsLiked(value, ConHeaderBiningDomain))
                {
                    //Read end column
                    subBinCutSheetInfo.EndColNum = endColumn - 1;
                    subBinCutSheetInfo.SubFlowEndColumnNum = subFlowendColumn;
                    subBinCutSheetInfos.Add(subBinCutSheetInfo.Copy());
                    subBinCutSheetInfo = new SubBinCutSheetInfo
                    {
                        StartColumnNum = i
                    };
                    firstFlag = true;
                }
            }
            if (!subBinCutSheetInfos.Exists(x => x.JobName.EqualsIgnoreCase(subBinCutSheetInfo.JobName)))
            {
                subBinCutSheetInfo.EndColNum = endColumn;
                subBinCutSheetInfo.SubFlowEndColumnNum = subFlowendColumn;
                subBinCutSheetInfos.Add(subBinCutSheetInfo.Copy());
            }
            return subBinCutSheetInfos;
        }

        private void ReadHeader(SubBinCutSheetInfo subBinCutSheetInfo)
        {
            for (int i = subBinCutSheetInfo.StartColumnNum; i <= subBinCutSheetInfo.EndColNum; i++)
            {
                string firstHeader = ExcelWorksheet.GetCellValue(StartRow, i);
                string secondHeader = ExcelWorksheet.GetCellValue(StartRow + 1, i);

                if (CellDiff.IsLiked(firstHeader, ConHeaderBiningDomain) ||
                   CellDiff.IsLiked(secondHeader, ConHeaderBiningDomain))
                {
                    _conHeaderBiningDomainIndex = i;
                }
                else if (CellDiff.IsLiked(firstHeader, ConHeaderPerformanceMode) ||
                  CellDiff.IsLiked(secondHeader, ConHeaderPerformanceMode))
                {
                    _conHeaderPerformanceModeIndex = i;
                }
            }
        }

        private NewBinCutFlowTable ReadSheet(SubBinCutSheetInfo subBinCutSheetInfo)
        {
            var binCutFlowTable = new NewBinCutFlowTable
            {
                SheetName = ExcelWorksheet.Name,
                BinningDomainIndex = _conHeaderBiningDomainIndex,
                PerformanceModeIndex = _conHeaderPerformanceModeIndex
            };

            string mergeCell = ExcelWorksheet.MergedCells[StartRow, StartCol];
            if (mergeCell != null)
            {
                binCutFlowTable.StartRowIndex = new ExcelAddress(ExcelWorksheet.MergedCells[StartRow, StartCol]).End.Row;
            }
            else
            {
                binCutFlowTable.StartRowIndex = StartRow;
            }

            binCutFlowTable.JobName = subBinCutSheetInfo.JobName;
            binCutFlowTable.FinalJob = new Job(subBinCutSheetInfo.JobName.Split(',').First()).JobType.ToString();

            EnumBinCutTableType tableType = EnumBinCutTableType.Lv;
            EnumBinCutTableBinType tableBinType = EnumBinCutTableBinType.Bin1;
            for (int i = StartRow + 2; i <= EndRow; i++)
            {
                var row = new NewBinCutFlowSheetRow(ExcelWorksheet.Name, binCutFlowTable.FinalJob)
                {
                    RowNum = i,

                    BinningDomain = ExcelWorksheet.GetCellValue(i, _conHeaderBiningDomainIndex).Trim()
                };
                if (row.BinningDomain.StartsWithIgnoreCase("End"))
                {
                    continue;
                }

                if (row.BinningDomain.Contains("HVCC", StringComparison.OrdinalIgnoreCase) || row.BinningDomain.Contains("HBV", StringComparison.OrdinalIgnoreCase))
                {
                    tableType = EnumBinCutTableType.Hv;
                }

                if (row.BinningDomain.Contains("BIN1", StringComparison.OrdinalIgnoreCase))
                {
                    tableBinType = EnumBinCutTableBinType.Bin1;
                }
                else if (row.BinningDomain.Contains("BINX", StringComparison.OrdinalIgnoreCase))
                {
                    tableBinType = EnumBinCutTableBinType.BinX;
                }
                else if (row.BinningDomain.Contains("BINY", StringComparison.OrdinalIgnoreCase))
                {
                    tableBinType = EnumBinCutTableBinType.BinY;
                }

                if (binCutFlowTable.SheetName.Contains("POST", StringComparison.OrdinalIgnoreCase))
                {
                    tableType = EnumBinCutTableType.Post;
                }

                row.TableType = tableType;
                row.TableBinType = tableBinType;
                row.PerformanceMode = ExcelWorksheet.GetCellValue(i, _conHeaderPerformanceModeIndex).Trim();

                var subFlows = new List<string>();
                for (int j = subBinCutSheetInfo.StartColumnNum + 2; j <= subBinCutSheetInfo.SubFlowEndColumnNum; j++)
                {
                    string value = ExcelWorksheet.GetCellValue(i, j).Trim();
                    if (string.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    if (!_regex2.IsMatch(value) && !value.StartsWithIgnoreCase("Flow_TMPS"))
                    {
                        value = row.PerformanceMode + "#" + value;
                    }

                    subFlows.Add(value);
                }
                row.SubFlows = [.. subFlows.Where(x => !string.IsNullOrEmpty(x))];

                if (!string.IsNullOrEmpty(row.PerformanceMode) && !row.PerformanceMode.EqualsIgnoreCase("Performance Mode"))
                {
                    binCutFlowTable.Rows.Add(row);
                }
            }
            return binCutFlowTable;
        }
    }
}
