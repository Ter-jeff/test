using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace ScghLib.Reader
{
    public partial class HardIpProdCharReader : MySheetReader<HardIpScghSheet>
    {
        private const string HeaderBlock = "Block";
        private const string HeaderMode = "Mode";
        private const string HeaderItem = "Item";
        private const string HeaderApplication = "Application";
        private const string HeaderUsage = "Usage";
        private const string HeaderForceCondition = "Force Condition";
        private const string HeaderDcUsed = "DC Spec";
        private const string HeaderEnableHln = "LV/NV/HV";
        private const string HeaderLabel = "Label";

        [GeneratedRegex(HeaderUsage, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();

        private readonly Dictionary<string, int> _headerIndexDic = [];
        private readonly List<int> _initIndexList = [];
        private readonly List<int> _payloadIndexList = [];

        public override HardIpScghSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;

            ReadHeader();

            return ReadContentFromOld() ?? new HardIpScghSheet();
        }

        private void ReadHeader()
        {
            string cellContent;
            for (int i = 1; i <= ExcelWorksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= ExcelWorksheet.Dimension.End.Column; j++)
                {
                    cellContent = EpplusExtensions.GetCellValue(ExcelWorksheet, i, j);
                    if (cellContent.EqualsIgnoreCase(HeaderBlock) ||
                        cellContent.EqualsIgnoreCase(HeaderLabel))
                    {
                        StartRow = i;
                        StartCol = j;
                        break;
                    }
                }
                if (StartRow != -1)
                {
                    break;
                }
            }

            if (StartRow == -1)
            {
                throw new Exception($"Can not find header \"Block\" in \"{ExcelWorksheet.Name}\" sheet. ");
            }

            for (int i = StartCol; i <= ExcelWorksheet.Dimension.End.Column; i++)
            {
                cellContent = EpplusExtensions.GetCellValue(ExcelWorksheet, StartRow, i).Trim().ToUpper();
                if (cellContent.ContainsIgnoreCase("INIT"))
                {
                    _initIndexList.Add(i);
                }
                else if (cellContent.ContainsIgnoreCase("PAYLOAD"))
                {
                    _payloadIndexList.Add(i);
                }
                else
                    if (MyRegex().IsMatch(cellContent))
                {
                    cellContent = HeaderUsage.ToUpper();
                }

                if (!string.IsNullOrEmpty(cellContent))
                {
                    _headerIndexDic.Add(cellContent, i);
                }
            }
        }

        private HardIpScghSheet? ReadContentFromOld()
        {
            int blockColumn = GetHeaderIndex(HeaderBlock);
            int modeColumn = GetHeaderIndex(HeaderMode);
            int itemColumn = GetHeaderIndex(HeaderItem);
            int applicationColumn = GetHeaderIndex(HeaderApplication);
            int usageColumn = GetHeaderIndex(HeaderUsage);
            int forceConditionColumn = GetHeaderIndex(HeaderForceCondition, true);
            int dcUsedColumn = GetHeaderIndex(HeaderDcUsed, true);
            int enableHlnColumn = GetHeaderIndex(HeaderEnableHln, true);
            if (blockColumn == -1 || modeColumn == -1)
            {
                return null;
            }

            var scghSheet = new HardIpScghSheet { SheetName = ExcelWorksheet.Name };
            for (int i = StartRow + 1; i <= ExcelWorksheet.Dimension.End.Row; i++)
            {
                var item = new ProdCharSheetRow(ExcelWorksheet.Name)
                {
                    Block = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, blockColumn),
                    Mode = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, modeColumn),
                    Item = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, itemColumn)
                };

                foreach (int initColumn in _initIndexList)
                {
                    string init = EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, initColumn);
                    if (!string.IsNullOrEmpty(init) && init != "NA")
                    {
                        item.InitList.Add(init);
                        item.InitAliasList.Add(init);
                    }
                }

                foreach (int payloadColumn in _payloadIndexList)
                {
                    string payload = EpplusExtensions.GetCellValue(ExcelWorksheet, i, payloadColumn);
                    if (!string.IsNullOrEmpty(payload) && payload != "NA")
                    {
                        item.PayloadList.Add(payload);
                        item.PayloadAliasList.Add(payload);
                    }
                }
                item.RowNum = i;
                item.Application = applicationColumn != -1 ? EpplusExtensions.GetMergedCellValue(ExcelWorksheet, i, applicationColumn) : GetApplication(item.PayloadValue);
                if (usageColumn != -1)
                {
                    item.Usage = EpplusExtensions.GetCellValue(ExcelWorksheet, i, usageColumn);
                }
                if (forceConditionColumn != -1)
                {
                    item.ForceCondition = EpplusExtensions.GetCellValue(ExcelWorksheet, i, forceConditionColumn);
                }

                if (dcUsedColumn != -1)
                {
                    item.DcUsed = EpplusExtensions.GetCellValue(ExcelWorksheet, i, dcUsedColumn);
                }

                if (enableHlnColumn != -1)
                {
                    item.EnableHlnv = EpplusExtensions.GetCellValue(ExcelWorksheet, i, enableHlnColumn);
                }

                if (!string.IsNullOrEmpty(item.PayloadValue))
                {
                    scghSheet.RowList.Add(item);
                }
            }

            foreach (ProdCharSheetRow row in scghSheet.RowList)
            {
                var all = scghSheet.RowList.SelectMany(x => x.GetInitAliasList()).ToList();
                all.AddRange([.. scghSheet.RowList.SelectMany(x => x.GetPayloadAliasList())]);
                if (all.Exists(x => x.EqualsIgnoreCase(row.Item)))
                {
                    row.IsGenFlow = false;
                }

                if (all.Exists(x => x.EqualsIgnoreCase(row.Mode)))
                {
                    row.IsGenFlow = false;
                }
            }

            return scghSheet;
        }

        private static string GetApplication(string payload)
        {
            if (payload.StartsWithIgnoreCase("pp_"))
            {
                return "Production";
            }
            else if (payload.StartsWithIgnoreCase("cz_"))
            {
                return "Characterization";
            }
            else if (payload.StartsWithIgnoreCase("dd_"))
            {
                return "Debug";
            }
            else if (payload.StartsWithIgnoreCase("ht_"))
            {
                return "HTOL";
            }

            return "";
        }

        private int GetHeaderIndex(string header, bool isOption = false)
        {
            if (_headerIndexDic.ContainsKey(header.ToUpper()))
            {
                return _headerIndexDic[header.ToUpper()];
            }

            if (!isOption)
            {
                throw new Exception(string.Join("Can not find header: \"{0}\" in \"{1}\" sheet.", header, ExcelWorksheet.Name));
            }

            return -1;
        }

    }

    public class HardIpScghSheet : ProdCharSheet
    {
    }
}
