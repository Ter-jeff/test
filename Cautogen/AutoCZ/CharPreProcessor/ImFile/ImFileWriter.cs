using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase;
using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader;
using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;
using Cautogen.common.ReaderWriter.Reader.InputReader;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Utility;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPreProcessor.ImFile
{
    public class ImFileWriter
    {
        public string WriteIntermediateFile(Dictionary<string, List<Characterization>> charPlanSheetDict, string imFilePath)
        {
            MessageWriter.WriteMessage("Writing result file...", EnumMessageLevel.Info);
            var dfcList = new Dictionary<string, List<string>>();
            using (var excel = new ExcelPackage(new FileInfo(imFilePath)))
            {
                foreach (KeyValuePair<string, List<Characterization>> planSheet in charPlanSheetDict.Where(planSheet => planSheet.Value.Count != 0))
                {
                    ExcelWorksheet sh = excel.Workbook.Worksheets.Add(planSheet.Key);
                    foreach (Characterization charItem in planSheet.Value)
                    {
                        if (!string.IsNullOrEmpty(charItem.FailInfo) && charItem.FailInfo.ToUpper().Contains("DFC"))
                        {
                            if (!dfcList.ContainsKey(charItem.TpName))
                            {
                                dfcList.Add(charItem.TpName, new List<string>());
                            }

                            dfcList[charItem.TpName].Add(charItem.FailInfo);
                        }
                        else if (!string.IsNullOrEmpty(charItem.Dfc))
                        {
                            string name =
                                $"{charItem.TpName}{charItem.Dfc}_CZ_{charItem.OtherSupplies.Split(' ').Last()}";
                            if (!dfcList.ContainsKey(name))
                            {
                                dfcList.Add(name, new List<string>());
                            }

                            if (!charItem.Dfc.Trim().Equals("0", StringComparison.OrdinalIgnoreCase))
                            {
                                dfcList[name].Add(charItem.Dfc);
                            }
                        }
                    }
                    ImFileRow.WriteHeader(sh);
                    _WriteContent(planSheet.Value, sh);
                }


                if (dfcList.Any())
                {
                    ExcelWorksheet dmc = excel.Workbook.Worksheets.Add("DFC_List");
                    _WriteDFCContent(dfcList, dmc);
                }

                if (CharPlan.OptionalTimesetting != null)
                {
                    ExcelWorksheet sh = excel.Workbook.Worksheets.Add("timesettings");
                    CharPlan.OptionalTimesetting.Write(sh);
                }

                // generate power merge result
                _WriteMergeResult(excel);
                _WriteEmaMappingTable(excel);
                _WriteAdaptiveCoolingTable(excel);
                if (CharPlan.IsContainsDashBoard)
                {
                    _WriteDashBoard(excel);
                }

                excel.Save();
            }

            LogHelper.Info("Finished!! You can check result in file " + imFilePath);
            MessageWriter.WriteMessage("\r\nFinished!! You can check result in file " + imFilePath, EnumMessageLevel.Result);
            return imFilePath;
        }

        private static void _WriteDashBoard(ExcelPackage ep)
        {
            ExcelWorksheet sh = ep.Workbook.Worksheets.Add("PatternDashBoard");
            try
            {

                int rowindex = 1;
                foreach (KeyValuePair<string, PatternData> pattern in PatternListInputReader.PatternList)
                {
                    sh.Cells[rowindex, 1].Value = pattern.Key;
                    sh.Cells[rowindex, 2].Value = pattern.Value.Use;
                    sh.Cells[rowindex, 3].Value = pattern.Value.FileVersion;
                    sh.Cells[rowindex, 4].Value = pattern.Value.TimesetVersion;
                    rowindex++;
                }

                MessageWriter.WriteMessage("Generate sheet " + sh.Name + " successfully!", EnumMessageLevel.Info);
            }
            catch (Exception e)
            {
                MessageWriter.WriteMessage("Generate sheet " + sh.Name + " failed! " + e.Message, EnumMessageLevel.Error);
            }
        }

        private static void _WriteDFCContent(Dictionary<string, List<string>> dfcRowList, ExcelWorksheet sh)
        {
            try
            {
                int rowindex = 2;
                sh.Cells[1, 1].Value = "Test Instance";
                foreach (KeyValuePair<string, List<string>> dfc in dfcRowList)
                {
                    if (dfc.Value.Count == 0)
                    {
                        continue;
                    }

                    sh.Cells[rowindex, 1].Value = dfc.Key;
                    rowindex++;
                }
                MessageWriter.WriteMessage("Generate sheet " + sh.Name + " successfully!", EnumMessageLevel.Info);
            }
            catch (Exception e)
            {
                MessageWriter.WriteMessage("Generate sheet " + sh.Name + " failed! " + e.Message, EnumMessageLevel.Error);
            }
        }


        private static void _WriteContent(IEnumerable<Characterization> charPlanSheet, ExcelWorksheet sh)
        {
            try
            {
                var imFileRowList = CharRowGroup.GetCharRowGroups(charPlanSheet).Select(charRow => new ImFileRow(charRow, sh.Name)).ToList();

                sh.Cells[3, 1].LoadFromCollection(imFileRowList);

                // auto fit
                for (int i = 1; i <= sh.Dimension.End.Column; i++)
                {
                    sh.Column(i).TryAutoFit();
                }

                MessageWriter.WriteMessage("Generate sheet " + sh.Name + " successfully!", EnumMessageLevel.Info);
            }
            catch (Exception e)
            {
                MessageWriter.WriteMessage("Generate sheet " + sh.Name + " failed! " + e.Message, EnumMessageLevel.Error);
            }
        }

        private static void _WriteMergeResult(ExcelPackage excel)
        {
            if (UtilityMain.UtilityData.PowerMergeResult.Rows.Count < 2)
            {
                return;
            }

            try
            {
                ExcelWorksheet wSheet = excel.Workbook.Worksheets.Add("PowerMergeResult");
                wSheet.Cells.LoadFromDataTable(UtilityMain.UtilityData.PowerMergeResult, true);
                wSheet.Column(1).TryAutoFit();
                wSheet.Column(2).TryAutoFit();
                wSheet.Column(3).TryAutoFit();
                MessageWriter.WriteMessage("Generate PowerMergeResult sheet successfully!", EnumMessageLevel.Info);
            }
            catch (Exception e)
            {
                MessageWriter.WriteMessage("Write Power Merge Result failed! " + e.Message, EnumMessageLevel.Error);
            }
        }

        private static void _WriteEmaMappingTable(ExcelPackage excel)
        {
            if (UtilityMain.UtilityData.EmaMappingItems.Count == 0)
            {
                return;
            }

            ExcelWorksheet wSheet = excel.Workbook.Worksheets.Add("EmaMapping");
            int rowIndex = 1;

            foreach (Cautogen.Utility.EmaMappingItem item in UtilityMain.UtilityData.EmaMappingItems)
            {
                wSheet.Cells[rowIndex, 1].Value = item.Pattern;
                wSheet.Cells[rowIndex, 2].Value = "segment";
                #region print header
                int colindex = 3;
                foreach (string @case in item.CasesList)
                {
                    wSheet.Cells[rowIndex, colindex].Value = @case;
                    colindex++;
                }
                #endregion
                rowIndex++;
                #region print data by different cases
                foreach (KeyValuePair<string, List<Cautogen.Utility.EmaSubset>> subSet in item.ReferenceSets)
                {
                    wSheet.Cells[rowIndex, 2].Value = subSet.Key;
                    int colIndex = 3;
                    foreach (string subSetCase in subSet.Value.First().Data.Values)
                    {
                        wSheet.Cells[rowIndex, colIndex].Value = subSetCase;
                        colIndex++;
                    }
                    rowIndex++;
                }
                #endregion
                //wSheet.Cells[rowIndex, 1].Value = item.Label;
                //wSheet.Cells[rowIndex, 2].Value = item.Value;

            }

        }

        private static void _WriteAdaptiveCoolingTable(ExcelPackage excel)
        {

            if (CharPlan.AdaptiveCooling == null)
            {
                return;
            }

            if (!CharPlan.AdaptiveCooling.Any())
            {
                return;
            }

            ExcelWorksheet wSheet = excel.Workbook.Worksheets.Add("AdaptiveCooling");
            //Writing header
            wSheet.Cells[1, 1].Value = "Insertion";
            wSheet.Cells[1, 2].Value = "TemperatureC";
            wSheet.Cells[1, 3].Value = "Enable";
            wSheet.Cells[1, 4].Value = "minDeltaC";
            wSheet.Cells[1, 5].Value = "maxDeltaC";
            wSheet.Cells[1, 6].Value = "TimeoutSec";

            int rowIndex = 2;

            foreach (KeyValuePair<string, AdaptiveCoolingData> item in CharPlan.AdaptiveCooling)
            {
                wSheet.Cells[rowIndex, 1].Value = item.Key;
                wSheet.Cells[rowIndex, 2].Value = item.Value.TemperatureC;
                wSheet.Cells[rowIndex, 3].Value = item.Value.Enable;
                wSheet.Cells[rowIndex, 4].Value = item.Value.MinDeltaC;
                wSheet.Cells[rowIndex, 5].Value = item.Value.MaxDeltaC;
                wSheet.Cells[rowIndex, 6].Value = item.Value.TimeoutSec;

                rowIndex++;
            }
        }

    }
}
