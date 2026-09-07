using System;
using System.Collections.Generic;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.BinNumberLegacy
{
    public class SoftBinDetailReader
    {
        private const string ConCategory = "category";
        private const string ConNumber = "number";
        private const string ConSubdig = "Sub dig";
        private const string ConColumn = "Column";

        public static List<SoftBinDetail> ReadSheet(ExcelWorksheet excelWorksheet)
        {
            List<SoftBinDetail> detailList = [];
            try
            {
                for (int i = 3; i <= excelWorksheet.Dimension.End.Row; i++)
                {
                    string range = excelWorksheet.MergedCells[i, 1];
                    int start = i;
                    int end;
                    if (string.IsNullOrEmpty(range))
                    {
                        end = i;
                    }
                    else
                    {
                        end = new ExcelAddress(range).End.Row;
                        i = end;
                    }

                    SoftBinDetail oneCat = ReadOneCat(excelWorksheet, start, end);
                    oneCat.Category = ReadConfigCell(excelWorksheet, start, 1);
                    oneCat.SheetName = excelWorksheet.Name;
                    if (oneCat.Category?.Length != 0)
                    {
                        detailList.Add(oneCat);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return detailList;
        }

        private static SoftBinDetail ReadOneCat(ExcelWorksheet excelWorksheet, int startRow, int endRow)
        {
            SoftBinDetail categoryDetail = new SoftBinDetail();
            for (int i = 2; i <= excelWorksheet.Dimension.End.Column; i++)
            {
                string range = excelWorksheet.MergedCells[1, i];
                if (string.IsNullOrEmpty(range))
                {
                    break;
                }

                int start = i;
                int end = new ExcelAddress(range).End.Column;
                i = end;
                SoftDetailDigiDef digitalDef = ReadOneDigi(excelWorksheet, startRow, endRow, start, end);
                if (digitalDef != null)
                {
                    categoryDetail.DigiDefList.Add(digitalDef);
                }
            }
            return categoryDetail;
        }

        private static SoftDetailDigiDef ReadOneDigi(ExcelWorksheet excelWorksheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            SoftDetailDigiDef digiDef = new SoftDetailDigiDef();
            string columnName = "";
            string name;
            string condition;

            for (int i = startRow; i <= endRow; i++)
            {
                SoftDetailNumDef numDef = new SoftDetailNumDef();
                for (int j = startColumn; j <= endColumn; j++)
                {
                    columnName = EpplusExtensions.GetCellValue(excelWorksheet, 2, j);
                    if (columnName.EqualsIgnoreCase(ConCategory))
                    {
                        numDef.Category = ReadConfigCell(excelWorksheet, i, j);
                    }
                    else if (columnName.EqualsIgnoreCase(ConNumber))
                    {
                        numDef.Number = ReadConfigCell(excelWorksheet, i, j);
                    }
                    else if (columnName.EqualsIgnoreCase(ConColumn))
                    {
                        name = EpplusExtensions.GetCellValue(excelWorksheet, i, j);
                        condition = EpplusExtensions.GetCellValue(excelWorksheet, i, j + 1);
                        if (name.Length != 0)
                        {
                            numDef.AddCondition(name, condition);
                        }
                        j++;
                    }
                    else if (columnName.EqualsIgnoreCase(ConSubdig))
                    {
                        string subdig = EpplusExtensions.GetCellValue(excelWorksheet, i, j);
                        string[] digList = subdig.Split(',');
                        foreach (string s in digList)
                        {
                            numDef.Subdig.Add(s);
                        }
                    }
                }
                if (numDef.Category?.Length != 0 && !digiDef.NumDefList.Exists(p => p.Category == numDef.Category))
                {
                    digiDef.NumDefList.Add(numDef);
                }
            }

            return digiDef;
        }

        private static string ReadConfigCell(ExcelWorksheet excelWorksheet, int row, int column)
        {
            return EpplusExtensions.GetCellValue(excelWorksheet, row, column).Trim();
        }
    }
}
