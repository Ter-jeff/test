using System;
using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut
{
    public abstract class PowerBinningHarvSheetReaderBase<T1, T2> : MySheetReader<T1> where T1 : PowerBinningHavrSheet where T2 : PowerBinningHarvRow
    {
        protected readonly Dictionary<string, int> InputHeaderList = [];

        protected readonly Dictionary<string, int> PowerBinningSheetHeaderList = [];

        protected int IndexPowerbinning = -1;
        protected int IndexFuseName2Index = -1;
        protected int IndexFailCommand = -1;
        protected int IndexCommentIndex = -1;

        protected abstract bool GetFirstHeaderList();

        public override T1? ReadSheet(ExcelWorksheet excelWorksheet)
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

            if (!GetFirstHeaderList())
            {
                return null;
            }

            T1 powerBinningSheet = ReadSheet(sheetName);

            return powerBinningSheet;
        }

        protected virtual T1 ReadSheet(string sheetName)
        {
            var powerBinningSheet = (T1)Activator.CreateInstance(typeof(T1), sheetName)!;

            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = (T2)Activator.CreateInstance(typeof(T2), sheetName)!;
                row.RowNum = i;
                for (int j = StartCol; j <= EndCol; j++)
                {
                    string text = ExcelWorksheet.GetCellValue(i, j);
                    foreach (KeyValuePair<string, int> inputHeader in InputHeaderList)
                    {
                        if (j == inputHeader.Value)
                        {
                            var data = new Dictionary<string, string>
                            {
                                { inputHeader.Key, text }
                            };
                            row.Inputinfo.Add(data);
                        }
                    }
                    foreach (KeyValuePair<string, int> powerBinningSheetHeader in PowerBinningSheetHeaderList)
                    {
                        if (j == powerBinningSheetHeader.Value)
                        {
                            var data = new Dictionary<string, string>
                            {
                                { powerBinningSheetHeader.Key, text }
                            };
                            row.SheetInfo.Add(data);
                        }
                    }
                }

                if (IndexPowerbinning != -1)
                {
                    row.PowerBinning = ExcelWorksheet.GetCellValue(i, IndexPowerbinning).Trim();
                }

                if (IndexFuseName2Index != -1)
                {
                    row.FuseName2 = ExcelWorksheet.GetCellValue(i, IndexFuseName2Index).Trim();
                }

                if (IndexFailCommand != -1)
                {
                    row.FailCommand = ExcelWorksheet.GetCellValue(i, IndexFailCommand).Trim();
                }

                if (IndexCommentIndex != -1)
                {
                    row.Comment = ExcelWorksheet.GetCellValue(i, IndexCommentIndex).Trim();
                }

                if (string.IsNullOrEmpty(ExcelWorksheet.GetCellValue(i, 1)))
                {
                    break;
                }

                powerBinningSheet.Rows.Add(row);
            }
            foreach (KeyValuePair<string, int> inputheader in InputHeaderList)
            {
                if (inputheader.Key.EqualsIgnoreCase("product_identifier"))
                {
                    continue;
                }

                powerBinningSheet.InputHeader.Add(inputheader.Key, inputheader.Value);
            }
            powerBinningSheet.IndexPowerbinning = IndexPowerbinning = -1;
            powerBinningSheet.IndexFuseName2Index = IndexFuseName2Index = -1;
            powerBinningSheet.IndexFailCommand = IndexFailCommand = -1;
            powerBinningSheet.IndexCommentIndex = IndexCommentIndex = -1;

            return powerBinningSheet;
        }
    }
}
