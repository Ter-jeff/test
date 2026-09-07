using System;
using System.Collections.Generic;
using System.Linq;

using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using LogLib.Utility;

using OfficeOpenXml;


namespace Automation.GenerateIgxl.HardIp.InputReader
{
    public class TestPlanReader
    {
        protected const string RegRunPattern = @"Run\s*the\s*pattern.*";

        protected const string ErrorMsgMissPat = "Missing pattern in \"Pattern\" Column";
        protected const string ErrorMsgMissPatWithMeasurement = "Missing pattern for measuments";
        protected const string ErrorMsgPatWithMeasurement = "pattern row exist measuments";
        protected int TtrIndex;
        protected int EnableIndex;
        protected int SiteflagIndex;
        protected int FailflagIndex;
        protected int PartIndex;
        protected int NoBinoutIndex;
        protected int DescriptionIndex;
        protected int PatternIndex;
        protected int ForceIndex;
        //protected int _forceCharIndex;
        protected int RegisterIndex;
        protected int MiscInfoIndex;
        protected int MeasIndex;
        protected int TestNameIndex;
        protected List<MeasLimit> MeasLimits;
        protected Dictionary<string, int> HeaderOrder;

        public ExcelWorksheet ExcelWorksheet;
        protected int StartRow;
        protected int EndRow;

        public virtual TestPlanSheet ReadSheet(ExcelWorksheet sheet, bool ignoreDescription = true)
        {
            var planSheet = new TestPlanSheet { SheetName = sheet.Name };
            ExcelWorksheet = sheet;
            HeaderOrder = EpplusExtensions.GetHeaderOrder(sheet);
            StartRow = 2;

            if (sheet.Dimension == null)
            {
                return planSheet;
            }

            EndRow = sheet.Dimension.End.Row;
            PatternRow patRow = null;
            //WirelessPatternRow patRow = null;
            //CheckUnrecognizeHeader
            CheckUnrecognizeHeader(HardIpPattern.KnownHeaders);

            //Initial herder index
            GetHeaderIndexs();
            SetHardipSheetsHeaderIdx(planSheet);

            //find the forceCondition used by all the patterns in the sheet
            planSheet.ForceIndex = ForceIndex;
            planSheet.ForceStr = GetSheetForceConditions();
            if (string.IsNullOrEmpty(planSheet.ForceStr))
            {
                StartRow = 2;
            }

            #region read testPlan patterns
            for (int j = StartRow; j <= EndRow; j++)
            {
                string patternName = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, j, PatternIndex)).Trim(',').Trim(';');
                if (patternName != "")
                {
                    patRow = new PatternRow();
                    //patRow = new WirelessPatternRow();
                    ReadPatternRow(sheet, j, patRow, patternName);
                    //PatternExistMeasument

                    planSheet.PatternRows.Add(patRow);
                    string measStr = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, j, MeasIndex));
                    if (!string.IsNullOrEmpty(measStr))
                    {
                        ErrorReportManager.AddError(HardIpErrorType.E_PatternExistMeasument_01, sheet.Name, j, 0, []);
                    }
                }
                else
                {

                    CheckRunPattern(sheet, j);
                    string measStr = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, j, MeasIndex));
                    string testName = EpplusExtensions.GetCellValue(sheet, j, TestNameIndex);
                    //If meas column not blank, but not belong to any pattern, flag error
                    if (patRow == null && !string.IsNullOrEmpty(measStr))
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_MisPatternForMeasument_01,
                            sheet.Name,
                            j,
                            0,
                            [])
                        ;
                        continue;
                    }
                    //Will ignore the rows that Pattern column, Force Condition column, RegisterAssignment and Meas column are blank
                    if (string.IsNullOrEmpty(measStr) && string.IsNullOrEmpty(testName))
                    {
                        if (patRow != null)
                        {
                            string registerAssignment = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, j, RegisterIndex));
                            //read RegisterAssignment
                            patRow.RegisterAssignment = UpdatePatternItem(patRow.RegisterAssignment, registerAssignment);

                            string forceStr = EpplusExtensions.GetCellValue(sheet, j, ForceIndex);
                            //Read post pattern force condition
                            patRow.PostPatForceCondition = UpdatePatternItem(patRow.PostPatForceCondition, forceStr);

                            string miscInfoStr = EpplusExtensions.GetCellValue(sheet, j, MiscInfoIndex);
                            //Read MiscInfo

                            bool isRtos = SearchInfo.IsHardipRtosSheet(sheet.Name);
                            patRow.MiscInfo = UpdatePatternItem(patRow.MiscInfo, miscInfoStr, isRtos);
                        }
                        continue;
                    }
                    ReadNonPatternRowNew(sheet, ref j, patRow);
                }
            }
            #endregion
            //DividePatternRow(planSheet.PatternRows);
            return planSheet;
        }

        protected void CheckUnrecognizeHeader(List<string> knownHeaders)
        {
            foreach (string header in HeaderOrder.Keys)
            {
                bool isKnown = false;
                foreach (string knowHeader in knownHeaders)
                {
                    if (Regex.IsMatch(header, knowHeader, RegexOptions.IgnoreCase))
                    {
                        isKnown = true;
                        break;
                    }

                }
                if (!isKnown)
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.W_UnrecognisedHeader_01,
                        ExcelWorksheet.Name,
                        1,
                        0,
                        [header]
                    );
                }
            }
        }

        protected void GetHeaderIndexs()
        {
            MeasLimits = new List<MeasLimit>();
            List<string> jobList = LocalSpecs.AllJobsHardIp ?? new List<string> { "CP1", "CP2", "FT1", "FT2", "FT3" };
            var jobColNotFoundList = new List<string>();
            foreach (string job in jobList)
            {
                var limit = new MeasLimit(job)
                {
                    LoHeaderIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, "^" + job + HardIpPattern.LoHeader,
                            false),
                    HiHeaderIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, "^" + job + HardIpPattern.HiHeader, false)
                };

                if (limit.LoHeaderIndex != -1 && limit.HiHeaderIndex != -1)
                {
                    limit.JobName = job;
                }

                if (limit.LoHeaderIndex == -1)
                {
                    jobColNotFoundList.Add($"{job} Lo Limit");
                }

                if (limit.HiHeaderIndex == -1)
                {
                    jobColNotFoundList.Add($"{job} Hi Limit");
                }

                MeasLimits.Add(limit);
            }
            if (jobColNotFoundList.Any())
            {
                string headers = string.Join(", ", jobColNotFoundList);
                ErrorReportManager.AddError(
                    HardIpErrorType.E_MissingHeader_01,
                    ExcelWorksheet.Name,
                    1,
                    0,
                    [headers, ExcelWorksheet.Name]
                );
            }
            TtrIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.TtrHeader);
            int headerIndex = HeaderOrder.FirstOrDefault(a => Regex.IsMatch(a.Key, HardIpPattern.NoBinOutHeader, RegexOptions.IgnoreCase) &&
                !Regex.IsMatch(a.Key, "Pattern Release Status")).Value;
            if (headerIndex > 0)
            {
                NoBinoutIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.NoBinOutHeader);
            }

            PartIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.PartHeader, false);
            DescriptionIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.DescriptionHeader);
            PatternIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.PatternHeader);
            ForceIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.ForceConditionHeader);
            //_forceCharIndex = GetHeaderIndex(_sheet.Name, _headerOrder, HardIpPattern.ForceConditionCharHeader, false);
            TestNameIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.TestNameHeader, false);
            RegisterIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.RegisterAssignmentHeader);
            MiscInfoIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.MiscInfoHeader);
            MeasIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.MeasHeader);
            EnableIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.EnableHeader, false);
            SiteflagIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.SiteFlagHeader, false);
            FailflagIndex = GetHeaderIndex(ExcelWorksheet.Name, HeaderOrder, HardIpPattern.FailFlagHeader, false);

        }

        protected string GetSheetForceConditions()
        {
            string forceStr = "";
            for (int j = StartRow; j <= EndRow; j++)
            {
                string patternName = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, j, PatternIndex));
                if (patternName != "")
                {
                    break;
                }

                string forceCondition = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, j, ForceIndex));
                if (forceCondition != "")
                {
                    forceStr += forceCondition + ";";
                }

                StartRow++;
            }
            return forceStr.Trim(';');
        }

        protected virtual TestPlanRow ReadTestPlanRow(int rowNum)
        {
            var testPlanRow = new TestPlanRow
            {
                RowNum = rowNum,
                Description = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, DescriptionIndex)),
                ForceCondition = SearchInfo.TrimSpace(EpplusExtensions.GetMergedCellValue(ExcelWorksheet, rowNum, ForceIndex)),
                MiscInfo = SearchInfo.TrimSpace(EpplusExtensions.GetMergedCellValue(ExcelWorksheet, rowNum, MiscInfoIndex)),
                RegisterAssignment = RegisterIndex != -1 ? SearchInfo.TrimSpace(EpplusExtensions.GetMergedCellValue(ExcelWorksheet, rowNum, RegisterIndex)) : ""
            };
            if (TestNameIndex != 1)
            {
                testPlanRow.TestName = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, TestNameIndex));
            }

            if (ExcelWorksheet.MergedCells[rowNum, MeasIndex] != null)
            {
                testPlanRow.MergeRowNumForMeas = new ExcelAddress(ExcelWorksheet.MergedCells[rowNum, MeasIndex]).Start.Row;
                testPlanRow.Meas = EpplusExtensions.GetCellValue(ExcelWorksheet, testPlanRow.MergeRowNumForMeas, MeasIndex).Replace("\n", "").Replace("\t", "").Trim(';');
            }
            else
            {
                testPlanRow.MergeRowNumForMeas = 0;
                testPlanRow.Meas = EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, MeasIndex).Replace("\n", "").Replace("\t", "").Trim(';');
            }
            foreach (MeasLimit limit in MeasLimits)
            {
                var newLimit = new MeasLimit(limit.JobName)
                {
                    LoLimit = EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, limit.LoHeaderIndex).Replace(" ", ""),
                    HiLimit = EpplusExtensions.GetCellValue(ExcelWorksheet, rowNum, limit.HiHeaderIndex).Replace(" ", ""),
                    LoHeaderIndex = limit.LoHeaderIndex,
                    HiHeaderIndex = limit.HiHeaderIndex
                };
                testPlanRow.Limits.Add(newLimit);
            }

            return testPlanRow;
        }

        public void SetHardipSheetsHeaderIdx(TestPlanSheet testPlanSheet)
        {
            Dictionary<string, int> idxDic = testPlanSheet.PlanHeaderIdx;
            idxDic.Add("ttrIndex", TtrIndex);
            idxDic.Add("noBinoutIndex", TtrIndex);
            idxDic.Add("enableIndex", EnableIndex);
            idxDic.Add("failflagIndex", FailflagIndex);
            idxDic.Add("siteflagIndex", SiteflagIndex);
            idxDic.Add("partIndex", PartIndex);
            idxDic.Add("descriptionIndex", DescriptionIndex);
            idxDic.Add("patternIndex", PatternIndex);
            idxDic.Add("forceIndex", ForceIndex);
            idxDic.Add("testnameIndex", TestNameIndex);
            idxDic.Add("registerIndex", RegisterIndex);
            idxDic.Add("miscInfoIndex", MiscInfoIndex);
            idxDic.Add("measIndex", MeasIndex);
            testPlanSheet.PlanHeaderIdx = idxDic;
        }

        protected void ReadPatternRow(ExcelWorksheet sheet, int rowIndex, PatternRow patRow, string patternName)
        {
            patRow.PatternColumnNum = PatternIndex;
            patRow.SheetName = sheet.Name;
            patRow.RowNum = rowIndex;

            if (rowIndex == -1 || TtrIndex == -1)
            {
                return;
            }

            patRow.TtrStr = SearchInfo.TrimSpace(EpplusExtensions.GetMergedCellValue(sheet, rowIndex, TtrIndex));
            if (PartIndex != -1)
            {
                patRow.PartStr = SearchInfo.TrimSpace(EpplusExtensions.GetMergedCellValue(sheet, rowIndex, PartIndex));
            }

            patRow.NoBinOutStr = SearchInfo.TrimSpace(EpplusExtensions.GetMergedCellValue(sheet, rowIndex, NoBinoutIndex));
            patRow.FailFlag = EpplusExtensions.GetCellValue(sheet, rowIndex, FailflagIndex);
            patRow.Enable = EpplusExtensions.GetCellValue(sheet, rowIndex, EnableIndex);
            patRow.SiteFlag = EpplusExtensions.GetCellValue(sheet, rowIndex, SiteflagIndex);
            patRow.Description = EpplusExtensions.GetCellValue(sheet, rowIndex, DescriptionIndex);
            patRow.Pattern = new PatternClass(patternName);

            patRow.ForceCondition.ForceCondition = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, rowIndex, ForceIndex));
            if (TestNameIndex != 1)
            {
                patRow.SpecifyTestName = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, rowIndex, TestNameIndex));
            }

            patRow.RegisterAssignment = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, rowIndex, RegisterIndex));

            if (SearchInfo.IsHardipRtosSheet(sheet.Name))
            {
                patRow.MiscInfo = EpplusExtensions.GetCellValue(sheet, rowIndex, MiscInfoIndex);
                patRow.MiscInfo = patRow.MiscInfo.Replace("\n", "").Replace("\r", "");
            }
            else
            {
                patRow.MiscInfo = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, rowIndex, MiscInfoIndex));
            }
        }

        protected virtual void ReadNonPatternRowNew(ExcelWorksheet sheet, ref int rowindex, PatternRow patRow)
        {
            try
            {
                //find whether the force condition column is merged or not
                bool isMergedCell = sheet.MergedCells[rowindex, ForceIndex] != null;
                bool isMergeKey = !isMergedCell && HardIpConstData.RegMergeIndex.IsMatch(sheet.Cells[rowindex, ForceIndex].Text ?? "");

                var patChildRow = new PatSubChildRow { IsMerged = isMergedCell || isMergeKey };
                if (!(isMergedCell || isMergeKey))
                {
                    patChildRow.TpRows.Add(ReadTestPlanRow(rowindex));
                }
                else
                {
                    int endrowindex = rowindex;
                    if (isMergedCell)
                    {
                        endrowindex = new ExcelAddress(sheet.MergedCells[rowindex, ForceIndex]).End.Row;
                    }
                    else
                    {
                        string mergeKey = HardIpConstData.RegMergeIndex.Match(sheet.Cells[endrowindex, ForceIndex].Text ?? "").Value;

                        if (!string.IsNullOrWhiteSpace(mergeKey))
                        {
                            while ((sheet.Cells[endrowindex + 1, ForceIndex].Text ?? "").Contains(mergeKey))
                            {
                                endrowindex++;
                            }
                        }
                    }

                    for (int mergedRow = rowindex; mergedRow <= endrowindex; mergedRow++)
                    {
                        patChildRow.TpRows.Add(ReadTestPlanRow(mergedRow));
                    }

                    rowindex = endrowindex;
                }
                patRow.PatChildRows.Add(patChildRow);
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        protected void CheckRunPattern(ExcelWorksheet sheet, int rowIndex)
        {
            string description = SearchInfo.TrimSpace(EpplusExtensions.GetCellValue(sheet, rowIndex, DescriptionIndex));
            //If "Run the pattern" exist Description column but pattern column is blank, flag error
            if (Regex.IsMatch(description, RegRunPattern, RegexOptions.IgnoreCase))
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_MissingPatternInTestPlan_01,
                    sheet.Name,
                    rowIndex,
                    PatternIndex,
                    []
                );
            }
        }

        protected string UpdatePatternItem(string origin, string value, bool isRtos = false)
        {
            string result = origin;
            if (!string.IsNullOrEmpty(value))
            {
                result += value;
            }

            return isRtos ? result : SearchInfo.TrimSpace(result);
        }

        public List<PatternRow> FilterTestItemsByPatterns(List<PatternRow> items, List<string> usedPatterns)
        {
            return items.Where(item => usedPatterns.Exists(p => item.Pattern.PatternSetList.SelectMany(x => x).Any(x => x.Equals(p, StringComparison.OrdinalIgnoreCase)))).ToList();
        }

        public int GetHeaderIndex(string sheetName, Dictionary<string, int> headerOrder, string header, bool optionalFlag = true)
        {
            int headerIndex = headerOrder.FirstOrDefault(a => Regex.IsMatch(a.Key, header, RegexOptions.IgnoreCase) && !a.Key.ContainsIgnoreCase("Pattern Release Status")).Value;
            if (headerIndex > 0)
            {
                return headerIndex;
            }

            if (optionalFlag)
            {
                header = header.Replace(@"\s*", " ").Replace(@"\s", " ").Replace(".*", "");
                ErrorReportManager.AddError(
                    HardIpErrorType.E_MissingHeader_01,
                    sheetName,
                    1,
                    0,
                    [header, sheetName]
                );
            }
            return -1;
        }
    }
}
