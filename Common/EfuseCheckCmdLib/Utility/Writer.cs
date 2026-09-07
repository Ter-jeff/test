using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.Static;

using CommonLib.Extension;

using EfuseCheckCmdLib.AlgorithmCheck;
using EfuseCheckCmdLib.EFuse.EFuseApp;
using EfuseCheckCmdLib.IgxlLogLib;

using LogLib.Utility;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using TestPlanLib.Efuse;

namespace EfuseCheckCmdLib.Utility
{
    public static class Writer
    {
        public static void SetHarvestSheetSummary(ExcelPackage excelPackage, ExcelWorksheet excelWorksheet)
        {
            if (HarvestFieldReader.HarvestFieldList.Count == 0)
            {
                return;
            }

            const string harvesFieldSheetName = "HarvestField";
            ExcelWorksheet harvestWorkSheet = excelPackage.Workbook.Worksheets.Add(SheetConst.Type5BitDefTable + "_" + harvesFieldSheetName);
            int startIdx = 12;
            try
            {
                excelWorksheet.Cells[1, 1, startIdx, excelWorksheet.Dimension.Columns].Copy(harvestWorkSheet.Cells[1, 1, startIdx, excelWorksheet.Dimension.Columns]);
                startIdx++;

                for (int i = 1; i <= excelWorksheet.Dimension.Rows; i++)
                {
                    object fieldName = excelWorksheet.Cells[i, 1].Value;
                    if (fieldName != null && HarvestFieldReader.HarvestFieldList.Exists(x => fieldName.ToString()!.EqualsIgnoreCase(x)))
                    {
                        excelWorksheet.Cells[i, 1, i, excelWorksheet.Dimension.Columns].Copy(harvestWorkSheet.Cells[startIdx, 1, startIdx, excelWorksheet.Dimension.Columns]);
                        startIdx++;
                    }
                }
            }
            catch (Exception)
            {
                EfuseStatic.Result = EfuseCheckResultType.Exception;
            }
            //HarvestFieldReader.HarvestFieldList.Clear();
        }

        public static void WriteFusePatternSummarySheet(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, XDiceInfo xDiceInfo)
        {
            if (excelWorksheet == null)
            {
                return;
            }

            try
            {
                int rowindex = 2;
                List<DataFormatDataRow> patternLines = xDiceInfo.FusePatternLines;
                excelWorksheet.Cells[1, 1].Value = "TestInstance Name";
                excelWorksheet.Cells[1, 2].Value = "Pattern";
                foreach (IGrouping<string, DataFormatDataRow> group in patternLines.GroupBy(p => p.TestName))
                {
                    if (string.IsNullOrEmpty(group.ToString()))
                    {
                        continue;
                    }

                    excelWorksheet.Cells[rowindex, 1].Value = group.Key;
                    foreach (DataFormatDataRow dataFormatDataRow in group)
                    {
                        excelWorksheet.Cells[rowindex, 2].Value = dataFormatDataRow.Pattern;
                        rowindex++;
                    }

                }
            }
            catch (Exception)
            {
                EfuseStatic.Result = EfuseCheckResultType.Exception;
                appendRichText("Generating the fuse pattern summary report failed, abort!!", "Red");
            }
        }

        public static void PrintCfgTableHeader(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, EfuseCfgTable efuseCfgTable, ref int colIndx)
        {
            int rowindx = 4;
            colIndx = 1;
            if (!efuseCfgTable.Scenario.TryGetValue(XParseDatalog.ScenarioInDatalog, out int cfgScenarioIndx))
            {
                if (EfuseStatic.IsCmd)
                {
                    appendRichText("Config didn't find! " + XParseDatalog.ScenarioInDatalog, "Red");
                }
                else
                {
                    ErrorMessageBox.Show("Config didn't find! " + XParseDatalog.ScenarioInDatalog);
                }

                return;
            }

            #region
            var indexList = new List<int> { efuseCfgTable.ConditionIdx, efuseCfgTable.MsbBitIdx, efuseCfgTable.LsbBitIdx, efuseCfgTable.BitWidthIdx, efuseCfgTable.PrgStageIdx, cfgScenarioIndx };
            #endregion
            //print Scenario(CFG and Datalog)
            excelWorksheet.Cells[1, colIndx].Value =
                $"Scenario in ConfigTable:{XParseDatalog.ScenarioInDatalog}, in Datalog:{XParseDatalog.ScenarioInDatalog}";
            excelWorksheet.Cells[2, colIndx].Value = $"Flag DisableChkLMT:{XParseDatalog.IsDisableChkLmt}";
            //print CFG Table Header
            foreach (int item in indexList)
            {
                excelWorksheet.Cells[rowindx, colIndx].Value = efuseCfgTable.TitleList[item];
                colIndx++;
            }
            rowindx++;

            //print CFG Table Content
            foreach (List<string> rows in efuseCfgTable.CfgRows)
            {
                colIndx = 1;
                if (rows.Count < cfgScenarioIndx)
                {
                    continue;
                }

                excelWorksheet.Cells[rowindx, colIndx].Value = rows[efuseCfgTable.ConditionIdx];
                colIndx++;
                excelWorksheet.Cells[rowindx, colIndx].Value = rows[efuseCfgTable.MsbBitIdx];
                colIndx++;
                excelWorksheet.Cells[rowindx, colIndx].Value = rows[efuseCfgTable.LsbBitIdx];
                colIndx++;
                excelWorksheet.Cells[rowindx, colIndx].Value = rows[efuseCfgTable.BitWidthIdx];
                colIndx++;
                excelWorksheet.Cells[rowindx, colIndx].Value = rows[efuseCfgTable.PrgStageIdx];
                colIndx++;
                excelWorksheet.Cells[rowindx, colIndx].Value = rows[cfgScenarioIndx];

                rowindx++;
            }
        }

        public static FileInfo CreateNewFile(string filename)
        {
            var newFile = new FileInfo(filename);
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(filename);
            }
            return newFile;
        }

        public static void HighLightCell(ExcelWorksheet excelWorksheet, int stRow, int stCol, int spRow, int spCol, Color color)
        {
            ExcelRange rng = excelWorksheet.SelectedRange[stRow, stCol, spRow, spCol];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(color);
        }

        public static void PrintBlockHeader(ExcelWorksheet excelWorksheet, EfuseBitDefTable efuseBitDefTable, List<string> oneRow, ref string curblock, ref int lastAdRow)
        {
            string latestblock = oneRow[efuseBitDefTable.BlockIdx];
            if (!Regex.IsMatch(latestblock, curblock, RegexOptions.IgnoreCase))
            {
                efuseBitDefTable.Titles[efuseBitDefTable.NameIdx] = latestblock;
                // EfuseBitDefTable.Titles[EfuseBitDefTable.NameIdx].Replace(curblock, latestblock);
            }
            for (int colIdx = 0; colIdx < efuseBitDefTable.Titles.Count; colIdx++)
            {
                excelWorksheet.Cells[lastAdRow, colIdx + 1].Value = efuseBitDefTable.Titles[colIdx];
            }

            curblock = latestblock;
            lastAdRow++;
        }

        public static bool IsNeedHighlightRow(string stage, EfuseScriptConfig efuseScriptConfig)
        {
            if (efuseScriptConfig.AliasProgStageDic.TryGetValue(stage, out string? value))
            {
                stage = value;
            }

            List<string> highLighJobs = [.. EfuseStatic.HighLightJobs.ToUpper().Split([","], StringSplitOptions.RemoveEmptyEntries)];
            if (highLighJobs.Contains(stage.ToUpper()))
            {
                return true;
            }

            return false;
        }
    }
}
