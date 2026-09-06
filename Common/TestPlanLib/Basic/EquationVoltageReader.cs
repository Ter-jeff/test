using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;
using CommonLib.Utility;

using CommonReaderLib;

using OfficeOpenXml;

using TestPlanLib.BinCut.Flow;
using TestPlanLib.Utility;

namespace TestPlanLib.Basic
{
    public partial class EquationVoltageReader(List<string>? binningTitleList = null, Dictionary<string, List<string>>? domainDic = null) : MySheetReader<BinCutFlowTables>
    {
        private const string ConHeaderPerformanceMode = "DC Spec|Performance Mode";
        private const string ConHeaderAllOther = "All Others";
        private const string ConHeaderVddCpu = "VDD_CPU";

        private const int ConMaxSearchColumn = 10;
        private const int ConMaxSearchRow = 10;

        [GeneratedRegex("^(?<str>.*)[@].*$", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        private readonly List<string>? _binningTitleList = binningTitleList;
        private readonly Dictionary<string, List<string>>? _domainDic = domainDic;

        public override BinCutFlowTables? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var outputSheets = new BinCutFlowTables();

            ExcelWorksheet = excelWorksheet;

            EndCol = ExcelWorksheet.Dimension.End.Column;
            EndRow = ExcelWorksheet.Dimension.End.Row;

            List<SubBinCutSheetInfo> subBinCutSheetInfos = ReadSubSheetInfo();

            foreach (SubBinCutSheetInfo subBinCutSheetInfo in subBinCutSheetInfos)
            {
                Dictionary<string, int> headers = ReadHeader(subBinCutSheetInfo);
                GetAllPinsInfo.AllPins allPins = GetAllPinsInfo.GetAllPins(subBinCutSheetInfo, StartRow, ExcelWorksheet);
                Dictionary<string, int> powerPins = allPins.PowerPins;
                BinCutFlowTable sheet = ReadSheet(subBinCutSheetInfo, headers, powerPins)!;
                sheet.PowerPins = powerPins;
                if (sheet.Rows.Count != 0)
                {
                    sheet.Check(_binningTitleList, _domainDic);
                    outputSheets.Add(sheet);
                }
            }
            return outputSheets;
        }

        private List<SubBinCutSheetInfo> ReadSubSheetInfo()
        {
            var subBinCutSheetInfos = new List<SubBinCutSheetInfo>();
            bool isFoundStartIdx = false;
            for (int i = 1; i < ConMaxSearchColumn; i++)
            {
                if (isFoundStartIdx)
                {
                    break;
                }

                for (int j = 1; j < ConMaxSearchRow; j++)
                {
                    string value = ExcelWorksheet.GetCellValue(j, i);
                    if (CellDiff.IsLiked(value, ConHeaderPerformanceMode))
                    {
                        StartCol = i;
                        StartRow = j;
                        isFoundStartIdx = true;
                        break;
                    }
                }
            }

            var subBinCutSheetInfo = new SubBinCutSheetInfo
            {
                StartColumnNum = StartCol,
                JobColumnNum = StartCol + 1,
                JobName = "CP1"
            };

            for (int i = StartCol + 1; i < EndCol; i++)
            {
                string value = ExcelWorksheet.GetCellValue(StartRow, i);
                //Read Job Name
                if (_regex.IsMatch(value))
                {
                    subBinCutSheetInfo.JobName = _regex.Match(value).Groups["str"].ToString().Trim();
                    subBinCutSheetInfo.JobColumnNum = i;
                }
                //Read end column
                //else if (IsLiked(value, ConHeaderBiningDomain))
                //{
                //    subBinCutSheetInfo.EndColNum = i - 1;
                //    subBinCutSheetInfos.Add(subBinCutSheetInfo);
                //    subBinCutSheetInfo = new SubBinCutSheetInfo();
                //    subBinCutSheetInfo.StartColumnNum = i;
                //    subBinCutSheetInfo.JobName = value;
                //}
                else if (CellDiff.IsLiked(value, ConHeaderPerformanceMode))
                {
                    //
                    subBinCutSheetInfo.EndColNum = i - 1;
                    subBinCutSheetInfos.Add(subBinCutSheetInfo);
                    subBinCutSheetInfo = new SubBinCutSheetInfo
                    {
                        StartColumnNum = i,
                        JobColumnNum = i + 1,
                        JobName = ExcelWorksheet.GetCellValue(StartRow, i + 1)
                    };
                }
            }
            subBinCutSheetInfo.EndColNum = EndCol;
            subBinCutSheetInfos.Add(subBinCutSheetInfo);
            return subBinCutSheetInfos;
        }

        private Dictionary<string, int> ReadHeader(SubBinCutSheetInfo subBinCutSheetInfo)
        {
            var headerColumnDictionary = new Dictionary<string, int>();
            for (int i = subBinCutSheetInfo.StartColumnNum; i <= subBinCutSheetInfo.EndColNum; i++)
            {
                string firstHeader = ExcelWorksheet.GetCellValue(StartRow, i);
                string secondHeader = ExcelWorksheet.GetCellValue(StartRow + 1, i);

                if (CellDiff.IsLiked(firstHeader, ConHeaderAllOther) ||
                    CellDiff.IsLiked(secondHeader, ConHeaderAllOther))
                {
                    headerColumnDictionary.Add(ConHeaderAllOther, i);
                    continue;
                }

                if (CellDiff.IsLiked(firstHeader, ConHeaderPerformanceMode) ||
                  CellDiff.IsLiked(secondHeader, ConHeaderPerformanceMode))
                {
                    headerColumnDictionary.Add(ConHeaderPerformanceMode, i);
                    continue;
                }

                if (CellDiff.IsLiked(firstHeader, ConHeaderVddCpu) ||
                 CellDiff.IsLiked(secondHeader, ConHeaderVddCpu))
                {
                    headerColumnDictionary.Add(ConHeaderVddCpu, i);
                }
            }
            return headerColumnDictionary;
        }

        private BinCutFlowTable? ReadSheet(SubBinCutSheetInfo subBinCutSheetInfo, Dictionary<string, int> headerColumnDictionary, Dictionary<string, int> powerPinDic)
        {
            var binCutFlowTable = new BinCutFlowTable
            {
                SheetName = ExcelWorksheet.Name
            };
            string mergeCell = ExcelWorksheet.MergedCells[StartRow, StartCol];
            var jobNames = new List<string>();
            if (mergeCell != null)
            {
                binCutFlowTable.Indices.StartRowIndex = new ExcelAddress(ExcelWorksheet.MergedCells[StartRow, StartCol]).End.Row;
            }
            else
            {
                binCutFlowTable.Indices.StartRowIndex = StartRow;
            }

            binCutFlowTable.JobName = subBinCutSheetInfo.JobName;
            //binCutFlowTable.FinalJob = new Job().GetJob(subBinCutSheetInfo.JobName).ToString();
            if (binCutFlowTable.JobName.Contains(','))
            {
                jobNames = [.. binCutFlowTable.JobName.Split(',').Select(x => x.Trim())];
            }
            else
            {
                jobNames.Add(binCutFlowTable.JobName);
            }

            foreach (string jobName in jobNames)
            {
                binCutFlowTable.FinalJob.Add(new Job(jobName).JobType.ToString());
            }
            if (!headerColumnDictionary.TryGetValue(ConHeaderPerformanceMode, out int performaceModeColumn) ||
                !headerColumnDictionary.TryGetValue(ConHeaderAllOther, out int allOtherColumn))
            //headerColumnDictionary.ContainsKey(ConHeaderAtpg) == false ||
            //headerColumnDictionary.ContainsKey(ConHeaderMbist) == false ||
            //headerColumnDictionary.ContainsKey(ConHeaderSpiRtos) == false)
            {
                return null;
            }

            //var atpgColumn = headerColumnDictionary.ContainsKey(ConHeaderAtpg) ? headerColumnDictionary[ConHeaderAtpg] : 0;
            //var mbist = headerColumnDictionary.ContainsKey(ConHeaderMbist) ? headerColumnDictionary[ConHeaderMbist] : 0;
            //var spiRtosColumn = headerColumnDictionary.ContainsKey(ConHeaderSpiRtos) ? headerColumnDictionary[ConHeaderSpiRtos] : 0;

            //binCutFlowTable.Indices.BiningDomainIndex = headerColumnDictionary[ConHeaderBiningDomain];
            binCutFlowTable.Indices.PerformanceModeIndex = performaceModeColumn;
            binCutFlowTable.Indices.AllOtherIndex = allOtherColumn;
            //binCutFlowTable.Indices.AtpgIndex = headerColumnDictionary.ContainsKey(ConHeaderAtpg) ? headerColumnDictionary[ConHeaderAtpg] : 0;
            //binCutFlowTable.Indices.MbistIndex = headerColumnDictionary.ContainsKey(ConHeaderMbist) ? headerColumnDictionary[ConHeaderMbist] : 0;
            //binCutFlowTable.Indices.SpiRtosIndex = headerColumnDictionary.ContainsKey(ConHeaderMbist) ? headerColumnDictionary[ConHeaderSpiRtos] : 0;

            const EnumBinCutTableType tableType = EnumBinCutTableType.Lv;
            EnumBinCutTableBinType tableBinType = EnumBinCutTableBinType.Bin1;
            for (int i = StartRow + 2; i <= EndRow; i++)
            {
                var row = new BinCutFlowSheetRow(ExcelWorksheet.Name, binCutFlowTable.FinalJob)
                {
                    RowNum = i
                };

                //row.BinningDomain = ExcelWorksheet.GetCellValue( i, biningDomainColumn).Trim();
                if (row.BinningDomain.StartsWithIgnoreCase("End"))
                {
                    continue;
                }

                //if (row.BinningDomain.ToUpper().Contains("HVCC") || row.BinningDomain.ToUpper().Contains("HBV"))
                //    tableType = EnumBinCutTableType.Hv;
                //if (row.BinningDomain.ToUpper().Contains("BIN1"))
                tableBinType = EnumBinCutTableBinType.Bin1;
                //else if (row.BinningDomain.ToUpper().Contains("BINX"))
                //    tableBinType = EnumBinCutTableBinType.BinX;
                //else if (row.BinningDomain.ToUpper().Contains("BINY"))
                //    tableBinType = EnumBinCutTableBinType.BinY;

                //if (binCutFlowTable.SheetName.ToUpper().Contains("POST"))
                //    tableType = EnumBinCutTableType.Post;

                row.TableType = tableType;
                row.TableBinType = tableBinType;
                row.PerformanceMode = ExcelWorksheet.GetCellValue(i, performaceModeColumn).Trim();
                row.AllOther = ExcelWorksheet.GetCellValue(i, allOtherColumn).Trim();
                //row.Atpg = atpgColumn == 0 ? "" : ExcelWorksheet.GetCellValue( i, atpgColumn).Trim();
                //row.Mbist = mbist == 0 ? "" : ExcelWorksheet.GetCellValue( i, mbist).Trim();
                //row.SpiRtos = spiRtosColumn == 0 ? "" : ExcelWorksheet.GetCellValue( i, spiRtosColumn).Trim();

                var binningValues = new List<PinInfo>();
                foreach (KeyValuePair<string, int> pin in powerPinDic)
                {
                    string value = ExcelWorksheet.GetCellValue(i, pin.Value).Trim();
                    binningValues.Add(new PinInfo { PinName = pin.Key, PinContext = value });
                }
                row.PinInfos = binningValues;

                if (!string.IsNullOrEmpty(row.PerformanceMode) && !row.PerformanceMode.EqualsIgnoreCase("DC Spec"))
                {
                    binCutFlowTable.Rows.Add(row);
                }
            }
            return binCutFlowTable;
        }
    }
}
