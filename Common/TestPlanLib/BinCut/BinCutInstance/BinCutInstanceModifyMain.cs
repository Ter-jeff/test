using System.Collections.Generic;
using System.Drawing;
using System.IO;

using CommonLib.Extension;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using TestPlanLib.BinCut.Flow;
using TestPlanLib.PatternListCsvFile;
using TestPlanLib.Utility;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public class BinCutInstanceModifyMain
    {
        public static void WorkFlow(BinCutInstanceSheet binCutInstanceSheet, ExcelWorksheet excelWorksheet, string file, BinCutFlowSheets binCutFlowSheets, Dictionary<string, OriPatListItem> patternDictionary)
        {
            var binCutInstanceSheetReader = new BinCutInstanceSheetReader();
            var newBinCutInstanceSheets = new List<BinCutInstanceSheet> { binCutInstanceSheet };

            using var excelPackage = new ExcelPackage(new FileInfo(file));
            List<ExcelWorksheet> olds = GetInstacneSheets(excelPackage);
            foreach (ExcelWorksheet old in olds)
            {
                BinCutInstanceSheet oldBinCutInstanceSheet = binCutInstanceSheetReader.ReadSheet(old)!;
                BinCutInstanceSheetChecker.WorkFlow(oldBinCutInstanceSheet, binCutFlowSheets, patternDictionary);
                BinCutInstanceSheet? newBinCutInstanceSheet = newBinCutInstanceSheets.Find(x => x.SheetName.EqualsIgnoreCase(old.Name));
                if (newBinCutInstanceSheet == null)
                {
                    continue;
                }

                for (int i = 0; i < newBinCutInstanceSheet.Rows.Count; i++)
                {
                    BinCutInstanceRow newRow = newBinCutInstanceSheet.Rows[i];
                    bool found = false;
                    for (int index = 0; index < oldBinCutInstanceSheet.Rows.Count; index++)
                    {
                        BinCutInstanceRow oldRow = oldBinCutInstanceSheet.Rows[index];

                        bool isPatternSame = IsPatternSame(oldRow, newRow.PatternList);
                        if (!isPatternSame)
                        {
                            continue;
                        }

                        found = true;
                        bool isAllSame = BinCutInstanceRowUtility.IsConditionAllSame(newRow, oldRow);
                        if (!isAllSame)
                        {
                            if (string.IsNullOrEmpty(newRow.PatSetNameOrange))
                            {
                                excelWorksheet.Cells[newRow.RowNum, newBinCutInstanceSheet.PatSetNameOrangeColNumber].Value = oldRow.PatSetNameOrange;
                                excelWorksheet.Cells[newRow.RowNum, newBinCutInstanceSheet.PatSetNameOrangeColNumber].Style.Font.Color.SetColor(Color.Red);
                            }
                            excelWorksheet.Cells[newRow.RowNum, newBinCutInstanceSheet.PatSetNameOrangeColNumber].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            excelWorksheet.Cells[newRow.RowNum, newBinCutInstanceSheet.PatSetNameOrangeColNumber].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(newRow.PatSetNameOrange))
                            {
                                excelWorksheet.Cells[newRow.RowNum, newBinCutInstanceSheet.PatSetNameOrangeColNumber].Value = oldRow.PatSetNameOrange;
                                excelWorksheet.Cells[newRow.RowNum, newBinCutInstanceSheet.PatSetNameOrangeColNumber].Style.Font.Color.SetColor(Color.Red);
                            }
                            excelWorksheet.Cells[newRow.RowNum, newBinCutInstanceSheet.PatSetNameOrangeColNumber].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            excelWorksheet.Cells[newRow.RowNum, newBinCutInstanceSheet.PatSetNameOrangeColNumber].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                            break;
                        }
                    }

                    if (!found)
                    {
                        excelWorksheet.Cells[newRow.RowNum, newBinCutInstanceSheet.PatSetNameOrangeColNumber].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[newRow.RowNum, newBinCutInstanceSheet.PatSetNameOrangeColNumber].Style.Fill.BackgroundColor.SetColor(Color.Purple);
                    }
                }
            }
        }

        private static List<ExcelWorksheet> GetInstacneSheets(ExcelPackage excelPackage)
        {
            var olds = new List<ExcelWorksheet>();
            foreach (ExcelWorksheet sheet in excelPackage.Workbook.Worksheets)
            {
                if (sheet.Name.StartsWithIgnoreCase("Instance_") ||
                    sheet.Name.StartsWithIgnoreCase("BinCut_Instance"))
                {
                    olds.Add(sheet);
                }
            }
            return olds;
        }

        public static bool IsPatternSame(BinCutInstanceRow binCutInstanceRow, List<string> patternList)
        {
            if (patternList.Count != binCutInstanceRow.PatternList.Count)
            {
                return false;
            }

            for (int index = 0; index < patternList.Count; index++)
            {
                string pattern = patternList[index];
                if (!pattern.EqualsIgnoreCase(binCutInstanceRow.PatternList[index]))
                {
                    return false;
                }
            }
            return true;
        }
    }
}
