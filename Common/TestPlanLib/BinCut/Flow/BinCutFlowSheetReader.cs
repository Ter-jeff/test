using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;
using CommonLib.Utility;

using CommonReaderLib;

using OfficeOpenXml;

using TestPlanLib.Utility;

namespace TestPlanLib.BinCut.Flow
{
    public partial class BinCutFlowSheetReader(List<string>? binningTitleList = null, Dictionary<string, List<string>>? domainDic = null) : MySheetReader<BinCutFlowTables>
    {
        private const string ConHeaderBiningDomain = "Binning.*";
        private const string ConHeaderPerformanceMode = "Performance.*";
        private const string ConHeaderAllOther = "All Others";
        private const string ConHeaderAtpg = "ATPG|TD";
        private const string ConHeaderMbist = "MBIST|BIST";
        private const string ConHeaderSpiRtos = "SPI/RTOS|FUNC";
        private const string ConHeaderVddCpu = "VDD_CPU";
        private const string ConHeaderStatic = "Static";

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
            var subBincutFlowSheets = subBinCutSheetInfos.Where(x => x is not StaticSelsramSheetInfo).ToList();

            foreach (SubBinCutSheetInfo subBincutFlowSheet in subBincutFlowSheets)
            {
                Dictionary<string, int> headers = ReadHeader(subBincutFlowSheet);
                GetAllPinsInfo.AllPins allPins = GetAllPinsInfo.GetAllPins(subBincutFlowSheet, StartRow, ExcelWorksheet);
                Dictionary<string, int> affiliatedPins = allPins.AffiliatedPins;
                Dictionary<string, int> powerPins = allPins.PowerPins;
                var selsramPins = new Dictionary<string, int>();

                BinCutFlowTable sheet = ReadSheet(subBincutFlowSheet, headers, powerPins, selsramPins)!;
                sheet.PowerPins = powerPins;
                sheet.AffiliatedPin = affiliatedPins;
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
                    if (CellDiff.IsLiked(value, ConHeaderBiningDomain))
                    {
                        StartCol = i;
                        StartRow = j;
                        isFoundStartIdx = true;
                        break;
                    }
                }
            }

            var subBinCutSheetInfo = new SubBinCutSheetInfo();
            var staticSelsramSheetInfo = new StaticSelsramSheetInfo();
            bool isStaticSelsram = false;
            subBinCutSheetInfo.StartColumnNum = StartCol;
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
                else if (CellDiff.IsLiked(value, ConHeaderBiningDomain))
                {
                    if (!isStaticSelsram)
                    {
                        subBinCutSheetInfo.EndColNum = i - 1;
                        subBinCutSheetInfos.Add(subBinCutSheetInfo);
                    }
                    else
                    {
                        staticSelsramSheetInfo.EndColNum = i - 1;
                        subBinCutSheetInfos.Add(staticSelsramSheetInfo);
                        isStaticSelsram = false;
                    }
                    subBinCutSheetInfo = new SubBinCutSheetInfo
                    {
                        StartColumnNum = i,
                        JobName = value
                    };
                }
                else if (CellDiff.IsLiked(value, ConHeaderPerformanceMode))
                {
                    subBinCutSheetInfo.JobName = ExcelWorksheet.GetCellValue(StartRow, i + 1);
                    subBinCutSheetInfo.JobColumnNum = i + 1;
                }
                else if (CellDiff.IsLiked(value, ConHeaderStatic))
                {
                    subBinCutSheetInfo.EndColNum = i - 1;
                    subBinCutSheetInfos.Add(subBinCutSheetInfo);
                    staticSelsramSheetInfo = new StaticSelsramSheetInfo();
                    isStaticSelsram = true;
                    staticSelsramSheetInfo.StartColumnNum = i;
                }
            }
            if (!isStaticSelsram)
            {
                subBinCutSheetInfo.EndColNum = EndCol;
                subBinCutSheetInfos.Add(subBinCutSheetInfo);
            }
            else
            {
                staticSelsramSheetInfo.EndColNum = EndCol;
                subBinCutSheetInfos.Add(staticSelsramSheetInfo);
            }
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
                }
                else if (CellDiff.IsLiked(firstHeader, ConHeaderAtpg) ||
                       CellDiff.IsLiked(secondHeader, ConHeaderAtpg))
                {
                    headerColumnDictionary.Add(ConHeaderAtpg, i);
                }
                else if (CellDiff.IsLiked(firstHeader, ConHeaderBiningDomain) ||
                   CellDiff.IsLiked(secondHeader, ConHeaderBiningDomain))
                {
                    headerColumnDictionary.Add(ConHeaderBiningDomain, i);
                }
                else if (CellDiff.IsLiked(firstHeader, ConHeaderMbist) ||
                  CellDiff.IsLiked(secondHeader, ConHeaderMbist))
                {
                    headerColumnDictionary.Add(ConHeaderMbist, i);
                }
                else if (CellDiff.IsLiked(firstHeader, ConHeaderPerformanceMode) ||
                  CellDiff.IsLiked(secondHeader, ConHeaderPerformanceMode))
                {
                    headerColumnDictionary.Add(ConHeaderPerformanceMode, i);
                }
                else if (CellDiff.IsLiked(firstHeader, ConHeaderSpiRtos) ||
                    CellDiff.IsLiked(secondHeader, ConHeaderSpiRtos))
                {
                    headerColumnDictionary.Add(ConHeaderSpiRtos, i);
                }
                else if (CellDiff.IsLiked(firstHeader, ConHeaderVddCpu) ||
                    CellDiff.IsLiked(secondHeader, ConHeaderVddCpu))
                {
                    headerColumnDictionary.Add(ConHeaderVddCpu, i);
                }
            }
            return headerColumnDictionary;
        }

        private BinCutFlowTable? ReadSheet(SubBinCutSheetInfo subBinCutSheetInfo, Dictionary<string, int> headerColumnDictionary, Dictionary<string, int> powerPinDic, Dictionary<string, int> sramPinDic)
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
            if (!headerColumnDictionary.TryGetValue(ConHeaderBiningDomain, out int biningDomainColumn) ||
                !headerColumnDictionary.TryGetValue(ConHeaderPerformanceMode, out int performaceModeColumn) ||
                !headerColumnDictionary.TryGetValue(ConHeaderAllOther, out int allOtherColumn))
            {
                return null;
            }

            int atpgColumn = headerColumnDictionary.TryGetValue(ConHeaderAtpg, out int value1) ? value1 : 0;
            int mbist = headerColumnDictionary.TryGetValue(ConHeaderMbist, out int value2) ? value2 : 0;
            int spiRtosColumn = headerColumnDictionary.TryGetValue(ConHeaderSpiRtos, out int value3) ? value3 : 0;

            binCutFlowTable.Indices.BiningDomainIndex = biningDomainColumn;
            binCutFlowTable.Indices.PerformanceModeIndex = performaceModeColumn;
            binCutFlowTable.Indices.AllOtherIndex = allOtherColumn;
            binCutFlowTable.Indices.AtpgIndex = headerColumnDictionary.TryGetValue(ConHeaderAtpg, out int value4) ? value4 : 0;
            binCutFlowTable.Indices.MbistIndex = headerColumnDictionary.TryGetValue(ConHeaderMbist, out int value5) ? value5 : 0;
            binCutFlowTable.Indices.SpiRtosIndex = headerColumnDictionary.ContainsKey(ConHeaderMbist) ? headerColumnDictionary[ConHeaderSpiRtos] : 0;

            EnumBinCutTableType tableType = EnumBinCutTableType.Lv;
            EnumBinCutTableBinType tableBinType = EnumBinCutTableBinType.Bin1;

            ReadData(powerPinDic, sramPinDic, binCutFlowTable, biningDomainColumn, performaceModeColumn, allOtherColumn, atpgColumn, mbist, spiRtosColumn, ref tableType, ref tableBinType);
            return binCutFlowTable;
        }

        private void ReadData(Dictionary<string, int> powerPinDic, Dictionary<string, int> sramPinDic, BinCutFlowTable binCutFlowTable, int biningDomainColumn, int performaceModeColumn, int allOtherColumn, int atpgColumn, int mbist, int spiRtosColumn, ref EnumBinCutTableType enumBinCutTableType, ref EnumBinCutTableBinType enumBinCutTableBinType)
        {
            for (int i = StartRow + 2; i <= EndRow; i++)
            {
                var row = new BinCutFlowSheetRow(ExcelWorksheet.Name, binCutFlowTable.FinalJob)
                {
                    RowNum = i,

                    BinningDomain = ExcelWorksheet.GetCellValue(i, biningDomainColumn).Trim()
                };
                if (row.BinningDomain.StartsWithIgnoreCase("End"))
                {
                    continue;
                }

                if (row.BinningDomain.Contains("HVCC", StringComparison.OrdinalIgnoreCase) || row.BinningDomain.Contains("HBV", StringComparison.OrdinalIgnoreCase))
                {
                    enumBinCutTableType = EnumBinCutTableType.Hv;
                }

                if (row.BinningDomain.Contains("BIN1", StringComparison.OrdinalIgnoreCase))
                {
                    enumBinCutTableBinType = EnumBinCutTableBinType.Bin1;
                }
                else if (row.BinningDomain.Contains("BINX", StringComparison.OrdinalIgnoreCase))
                {
                    enumBinCutTableBinType = EnumBinCutTableBinType.BinX;
                }
                else if (row.BinningDomain.Contains("BINY", StringComparison.OrdinalIgnoreCase))
                {
                    enumBinCutTableBinType = EnumBinCutTableBinType.BinY;
                }

                if (binCutFlowTable.SheetName.Contains("POST", StringComparison.OrdinalIgnoreCase))
                {
                    enumBinCutTableType = EnumBinCutTableType.Post;
                }

                if (binCutFlowTable.SheetName.Contains("OUTSIDE", StringComparison.OrdinalIgnoreCase))
                {
                    enumBinCutTableType = EnumBinCutTableType.Post;
                }

                row.TableType = enumBinCutTableType;
                row.TableBinType = enumBinCutTableBinType;
                row.PerformanceMode = ExcelWorksheet.GetCellValue(i, performaceModeColumn).Trim();
                row.AllOther = ExcelWorksheet.GetCellValue(i, allOtherColumn).Trim();
                row.Atpg = atpgColumn == 0 ? "" : ExcelWorksheet.GetCellValue(i, atpgColumn).Trim();
                row.Mbist = mbist == 0 ? "" : ExcelWorksheet.GetCellValue(i, mbist).Trim();
                row.SpiRtos = spiRtosColumn == 0 ? "" : ExcelWorksheet.GetCellValue(i, spiRtosColumn).Trim();

                var binningValues = new List<PinInfo>();
                foreach (KeyValuePair<string, int> pin in powerPinDic)
                {
                    string value = ExcelWorksheet.GetCellValue(i, pin.Value).Trim();
                    binningValues.Add(new PinInfo { PinName = pin.Key, PinContext = value });
                }
                row.PinInfos = binningValues;

                if (sramPinDic != null)
                {
                    var sramBits = new List<SelsramInfo>();
                    foreach (KeyValuePair<string, int> pin in sramPinDic)
                    {
                        string value = ExcelWorksheet.GetCellValue(i, pin.Value).Trim();
                        sramBits.Add(new SelsramInfo { PinName = pin.Key, Bit = value });
                    }
                    row.SelsramInfos = sramBits;
                    row.Static = !row.SelsramInfos.Exists(x => x.Bit == "Cross");
                }

                if (!string.IsNullOrEmpty(row.PerformanceMode) && !row.PerformanceMode.EqualsIgnoreCase("Performance Mode"))
                {
                    binCutFlowTable.Rows.Add(row);
                }
            }
        }
    }
}
