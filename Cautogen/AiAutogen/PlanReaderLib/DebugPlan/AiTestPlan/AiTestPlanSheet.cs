using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using CommonReaderLib;
using CommonReaderLib.PatternListCsv;

namespace DebugPlanReaderLib.DebugPlan
{
    public class AiTestPlanSheet : MySheet
    {
        public int IndexAiType = -1;
        public int IndexComment = -1;
        public int IndexDataLoggingSetting = -1;
        public int IndexOrder = -1;
        public int IndexPatternStart = -1;
        public int IndexSearch = -1;
        public int IndexSelsramDssc = -1;

        public int IndexStartRow = -1;
        public int IndexTempCondition = -1;
        public int IndexFailCyclePoint = -1;
        public int IndexTestInstanceName = -1;
        public int IndexTimeset = -1;
        public int IndexUseNotUse = -1;
        public int IndexVoltageCategory = -1;
        public int IndexPowerRunScenario = -1;
        public int IndexAcCategory = -1;
        public int IndexRetention = -1;
        public int IndexDigSrc = -1;
        public int IndexUSL = -1;
        public int IndexLSL = -1;

        #region Constructor

        public AiTestPlanSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<AiTestPlanRow>();
        }

        #endregion

        public List<AiTestPlanRow> Rows { set; get; }

        internal void Check()
        {
            CheckHeader();
            CheckByColumn();
        }

        private void CheckHeader()
        {
            if (IndexUseNotUse == -1)
            {
                AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["Use/Not Use"]);
                ErrorReportManager.AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["Use/Not Use"]);
            }
            if (IndexComment == -1)
            {
                AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["Comment"]);
                ErrorReportManager.AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["Comment"]);
            }
            if (IndexTestInstanceName == -1)
            {
                AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["Test instance name"]);
                ErrorReportManager.AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["Test instance name"]);
            }
            if (IndexAiType == -1)
            {
                AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["AI type"]);
                ErrorReportManager.AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["AI type"]);
            }
            if (IndexDataLoggingSetting == -1)
            {
                AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["Data logging setting"]);
                ErrorReportManager.AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["Data logging setting"]);
            }
            if (IndexTimeset == -1)
            {
                AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["Timeset"]);
                ErrorReportManager.AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["Timeset"]);
            }
            if (IndexVoltageCategory == -1)
            {
                AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["Voltage Category"]);
                ErrorReportManager.AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["Voltage Category"]);
            }
            if (IndexOrder == -1)
            {
                AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["Order"]);
                ErrorReportManager.AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["Order"]);
            }
            if (IndexSearch == -1)
            {
                AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["Search"]);
                ErrorReportManager.AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["Search"]);
            }
            if (IndexTempCondition == -1)
            {
                AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["Temp. Condition"]);
                ErrorReportManager.AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["Temp. Condition"]);
            }
            if (IndexSelsramDssc == -1)
            {
                AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["SELSRAM_DSSC"]);
                ErrorReportManager.AddError(AutoAiErrorType.E_MissingHeader_01, SheetName, IndexStartRow, 0, ["SELSRAM_DSSC"]);
            }
            if (IndexUSL == -1)
            {
                AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["USL"]);
                ErrorReportManager.AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["USL"]);
            }
            if (IndexLSL == -1)
            {
                AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["LSL"]);
                ErrorReportManager.AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["LSL"]);
            }
            if (IndexAcCategory == -1)
            {
                AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["Ac Category"]);
                ErrorReportManager.AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["Ac Category"]);
            }
            if (IndexPowerRunScenario == -1)
            {
                AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["Power Run Scenario"]);
                ErrorReportManager.AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["Power Run Scenario"]);
            }
            //if (IndexRetention == -1)
            //{
            //    Errors.Add(new Error()
            //    {
            //        EnumErrorType = EnumErrorType.MissingHeader,
            //        ErrorLevel = ErrorLevel.Warning,
            //        SheetName = Name,
            //        RowNum = IndexStartRow,
            //        ColNum = 0,
            //        Message = string.Format("[Header] Missing the header: {0}", "Retention")
            //    });
            //}
            if (IndexDigSrc == -1)
            {
                AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["DigSrc"]);
                ErrorReportManager.AddError(AutoAiErrorType.W_MissingHeader_02, SheetName, IndexStartRow, 0, ["DigSrc"]);
            }
        }
        private void CheckByColumn()
        {
            foreach (var row in Rows)
            {
                if (!row.UseNotUse.Equals("Use", StringComparison.CurrentCultureIgnoreCase) &&
                    !row.UseNotUse.Equals("Not Use", StringComparison.CurrentCultureIgnoreCase))
                {
                    AddError(AutoAiErrorType.W_FormatError_01, SheetName, row.RowNum, IndexUseNotUse, [row.UseNotUse]);
                    ErrorReportManager.AddError(AutoAiErrorType.W_FormatError_01, SheetName, row.RowNum, IndexUseNotUse, [row.UseNotUse]);
                }

                if (!row.AiType.Equals("Data log", StringComparison.CurrentCultureIgnoreCase) &&
                    !row.AiType.Equals("1D", StringComparison.CurrentCultureIgnoreCase) &&
                    !row.AiType.Equals("2D", StringComparison.CurrentCultureIgnoreCase))
                {
                    AddError(AutoAiErrorType.E_FormatError_02, SheetName, row.RowNum, IndexAiType, [row.AiType]);
                    ErrorReportManager.AddError(AutoAiErrorType.E_FormatError_02, SheetName, row.RowNum, IndexAiType, [row.AiType]);
                }

                if (!row.DataLoggingSetting.Equals("NA", StringComparison.CurrentCultureIgnoreCase) &&
                    !row.DataLoggingSetting.StartsWith("DFCList", StringComparison.CurrentCultureIgnoreCase) &&
                    !row.DataLoggingSetting.StartsWith("DFCStep", StringComparison.CurrentCultureIgnoreCase))
                //!Regex.IsMatch(row.DataLoggingSetting, @"\d+\s?FC", RegexOptions.IgnoreCase))
                {
                    AddError(AutoAiErrorType.E_FormatError_03, SheetName, row.RowNum, IndexDataLoggingSetting, [row.DataLoggingSetting]);
                    ErrorReportManager.AddError(AutoAiErrorType.E_FormatError_03, SheetName, row.RowNum, IndexDataLoggingSetting, [row.DataLoggingSetting]);
                }
                else
                {
                    CheckDataLoggingSetting(row);
                }

                if (!row.SelsramDssc.StartsWith("SELSRM", StringComparison.CurrentCultureIgnoreCase) &&
                    !row.SelsramDssc.StartsWith("DSELSRM", StringComparison.CurrentCultureIgnoreCase))
                {
                    AddError(AutoAiErrorType.E_FormatError_08, SheetName, row.RowNum, IndexSelsramDssc, [row.SelsramDssc]);
                    ErrorReportManager.AddError(AutoAiErrorType.E_FormatError_08, SheetName, row.RowNum, IndexSelsramDssc, [row.SelsramDssc]);
                }

                #region check order

                var arr = row.Order.Split(';').ToList();
                if (!string.IsNullOrEmpty(row.Order))
                {
                    if (row.EnumAiType == EnumAiType.Datalog)
                    {
                        //Errors.Add(new Error
                        //{
                        //    EnumErrorType = EnumErrorType.FormatError,
                        //    ErrorLevel = ErrorLevel.Error,
                        //    SheetName = Name,
                        //    RowNum = row.RowNum,
                        //    ColNum = IndexOrder,
                        //    Message = string.Format("[Order] The AI type is datalog, so \"{0}\" will be ignored !!!",
                        //        row.Order)
                        //});
                    }
                    else if (row.EnumAiType == EnumAiType.Shmoo1D && arr.Count != 1)
                    {
                        //Errors.Add(new Error
                        //{
                        //    EnumErrorType = EnumErrorType.FormatError,
                        //    ErrorLevel = ErrorLevel.Error,
                        //    SheetName = Name,
                        //    RowNum = row.RowNum,
                        //    ColNum = IndexOrder,
                        //    Message = string.Format(
                        //        "[Order] The AI type is 1D shmoo, so the count of \"{0}\" should be 1 !!!", row.Order)
                        //});
                    }
                    else if (row.EnumAiType == EnumAiType.Shmoo2D && arr.Count != 2)
                    {
                        //Errors.Add(new Error
                        //{
                        //    EnumErrorType = EnumErrorType.FormatError,
                        //    ErrorLevel = ErrorLevel.Error,
                        //    SheetName = Name,
                        //    RowNum = row.RowNum,
                        //    ColNum = IndexOrder,
                        //    Message = string.Format(
                        //        "[Order] The AI type is 2D shmoo, so the count of \"{0}\" should be 2 !!!", row.Order)
                        //});
                    }

                    foreach (var item in arr)
                    {
                        var pins = row.Pins.Select(x => x.Name).ToList();
                        if (!pins.Contains(item, StringComparer.OrdinalIgnoreCase))
                        {
                            //Errors.Add(new Error
                            //{
                            //    EnumErrorType = EnumErrorType.FormatError,
                            //    ErrorLevel = ErrorLevel.Error,
                            //    SheetName = Name,
                            //    RowNum = row.RowNum,
                            //    ColNum = IndexOrder,
                            //    Message = string.Format("[Order] The pin \"{0}\" is not existed !!!", item)
                            //});
                        }
                    }
                }
                #endregion
                CheckShmooPinCount(row);
                CheckShmooJump(row);
            }
        }
        private void CheckDataLoggingSetting(AiTestPlanRow row)
        {
            if (row.DataLoggingSetting.StartsWith("DFCList", StringComparison.CurrentCultureIgnoreCase))
            {
                const string pattern = @"\((?<value>[\w|,]+)\)";
                if (!Regex.IsMatch(row.DataLoggingSetting, pattern))
                {
                    AddError(AutoAiErrorType.E_FormatError_04, SheetName, row.RowNum, IndexDataLoggingSetting, [row.DataLoggingSetting]);
                    ErrorReportManager.AddError(AutoAiErrorType.E_FormatError_04, SheetName, row.RowNum, IndexDataLoggingSetting, [row.DataLoggingSetting]);
                }
                else
                {
                    var dfc = Regex.Match(row.DataLoggingSetting, pattern).Groups["value"].ToString();
                    var array = dfc.Split(',').Select(x => x.Trim()).Select(x => Regex.Replace(x, "mV$", "", RegexOptions.IgnoreCase)).ToList();
                    int value;
                    if (!array.All(x => int.TryParse(x, out value)))
                    {
                        AddError(AutoAiErrorType.E_FormatError_05, SheetName, row.RowNum, IndexDataLoggingSetting, [row.DataLoggingSetting]);
                        ErrorReportManager.AddError(AutoAiErrorType.E_FormatError_05, SheetName, row.RowNum, IndexDataLoggingSetting, [row.DataLoggingSetting]);
                    }
                }
            }
            else if (row.DataLoggingSetting.StartsWith("DFCStep", StringComparison.CurrentCultureIgnoreCase))
            {
                const string pattern = @"\((?<value>[\w|,]+)\)";
                if (!Regex.IsMatch(row.DataLoggingSetting, pattern))
                {
                    AddError(AutoAiErrorType.E_FormatError_06, SheetName, row.RowNum, IndexDataLoggingSetting, [row.DataLoggingSetting]);
                    ErrorReportManager.AddError(AutoAiErrorType.E_FormatError_06, SheetName, row.RowNum, IndexDataLoggingSetting, [row.DataLoggingSetting]);
                }
                else
                {
                    var dfc = Regex.Match(row.DataLoggingSetting, pattern).Groups["value"].ToString();
                    var array = dfc.Split(',').Select(x => x.Trim()).Select(x => Regex.Replace(x, "mV$", "", RegexOptions.IgnoreCase)).ToList();
                    int value;
                    if (!array.All(x => int.TryParse(x, out value)) || (array.Count() != 2))
                    {
                        AddError(AutoAiErrorType.E_FormatError_07, SheetName, row.RowNum, IndexDataLoggingSetting, [row.DataLoggingSetting]);
                        ErrorReportManager.AddError(AutoAiErrorType.E_FormatError_07, SheetName, row.RowNum, IndexDataLoggingSetting, [row.DataLoggingSetting]);
                    }
                }
            }
        }
        private void CheckShmooPinCount(AiTestPlanRow row)
        {
            if (row.EnumAiType == EnumAiType.Shmoo1D)
            {
                if (row.Pins.Where(x => x.IsSearch).Count() < 1)
                {
                    //Errors.Add(new Error
                    //{
                    //    EnumErrorType = EnumErrorType.FormatError,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexAiType,
                    //    Message = string.Format("[AIType] AIType is 1D, but the count of shmoo pin < 1 !!!")
                    //});
                }
            }
            if (row.EnumAiType == EnumAiType.Shmoo2D)
            {
                if (row.Pins.Where(x => x.IsSearch).Count() < 2)
                {
                    //Errors.Add(new Error
                    //{
                    //    EnumErrorType = EnumErrorType.FormatError,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexAiType,
                    //    Message = string.Format("[AIType] AIType is 2D, but the count of shmoo pin < 2 !!!")
                    //});
                }
            }
        }
        private void CheckShmooJump(AiTestPlanRow row)
        {
            var count = row.Pins.Count(x => x.IsSearch);
            if (row.EnumAiType == EnumAiType.Shmoo1D)
            {
                if (row.IsJump())
                {
                    var testMethods = row.GetTestMethods();
                    if (testMethods.Count != 1)
                    {
                        //Errors.Add(new Error
                        //{
                        //    EnumErrorType = EnumErrorType.FormatError,
                        //    ErrorLevel = ErrorLevel.Error,
                        //    SheetName = Name,
                        //    RowNum = row.RowNum,
                        //    ColNum = IndexSearch,
                        //    Message = string.Format("[Search] The \"{0}\" don't match AI type Shmoo_1D !!!", row.Search)
                        //});
                    }
                    else
                    {
                        var pin = row.GetPinElementAt(0);
                        var stepCnt = pin.Steps;
                        if (testMethods.First().Name.Equals("Jump", StringComparison.CurrentCultureIgnoreCase) &&
                            stepCnt < int.Parse(testMethods.First().Arguments))
                        {
                            //Errors.Add(new Error
                            //{
                            //    EnumErrorType = EnumErrorType.FormatError,
                            //    ErrorLevel = ErrorLevel.Error,
                            //    SheetName = Name,
                            //    RowNum = row.RowNum,
                            //    ColNum = IndexSearch,
                            //    Message = string.Format(
                            //        "[Search] The jump steps of \"{0}\" should be smaller than pin \"{1}\" \"{2}\" !!!",
                            //        row.Search, pin.Name, stepCnt.ToString("#.####"))
                            //});
                        }
                    }
                }
            }
            else if (row.EnumAiType == EnumAiType.Shmoo2D)
            {
                if (row.IsJump())
                {
                    var testMethods = row.GetTestMethods();
                    if (testMethods.Count != 2)
                    {
                        //Errors.Add(new Error
                        //{
                        //    EnumErrorType = EnumErrorType.FormatError,
                        //    ErrorLevel = ErrorLevel.Error,
                        //    SheetName = Name,
                        //    RowNum = row.RowNum,
                        //    ColNum = IndexSearch,
                        //    Message = string.Format("[Search] The \"{0}\" don't match AI type Shmoo_2D !!!",
                        //        row.Search)
                        //});
                    }
                    else
                    {
                        var pin = row.GetPinElementAt(0);
                        var stepCnt = pin.Steps;
                        if (testMethods.First().Name.Equals("Jump", StringComparison.CurrentCultureIgnoreCase) &&
                            stepCnt < int.Parse(testMethods.First().Arguments))
                        {
                            //Errors.Add(new Error
                            //{
                            //    EnumErrorType = EnumErrorType.FormatError,
                            //    ErrorLevel = ErrorLevel.Error,
                            //    SheetName = Name,
                            //    RowNum = row.RowNum,
                            //    ColNum = IndexSearch,
                            //    Message = string.Format(
                            //        "[Search] The jump steps of \"{0}\" should be smaller than pin \"{1}\" \"{2}\" !!!",
                            //        row.Search, pin.Name, stepCnt.ToString("#.####"))
                            //});
                        }

                        var testMethod = testMethods.ElementAt(1);
                        if (testMethod.Name.Equals("Jump", StringComparison.CurrentCultureIgnoreCase))
                        {
                            //Errors.Add(new Error
                            //{
                            //    EnumErrorType = EnumErrorType.FormatError,
                            //    ErrorLevel = ErrorLevel.Error,
                            //    SheetName = Name,
                            //    RowNum = row.RowNum,
                            //    ColNum = IndexSearch,
                            //    Message = string.Format("[Search] The Y Shmoo of \"{0}\" can not be jump !!!",
                            //        row.Search)
                            //});
                        }
                    }
                }
            }
        }

        internal void CheckTimeSet(List<string> timeSets)
        {
            foreach (var row in Rows)
            {
                var mappingErrs = CheckMappingTimeset(row);
                //var planErr = new Error
                //{
                //    EnumErrorType = EnumErrorType.MissingTimesetFile,
                //    ErrorLevel = ErrorLevel.Error,
                //    SheetName = Name,
                //    RowNum = row.RowNum,
                //    ColNum = IndexTimeset,
                //    Message = string.Format(
                //                "[Timeset] The {0} is not existed in timeSet folder or test program !!!", row.Timeset)
                //};

                if (!timeSets.Any(x => x.ToUpper().Equals(row.Timeset.ToUpper())))
                {
                    if (!mappingErrs.Any())
                    {
                        //Errors.Add(new Error()
                        //{
                        //    EnumErrorType = EnumErrorType.MissingTimesetFile,
                        //    ErrorLevel = ErrorLevel.Warning,
                        //    SheetName = Name,
                        //    RowNum = row.RowNum,
                        //    ColNum = IndexTimeset,
                        //    Message = string.Format(
                        //        "[Timeset] The {0} is not existed in timeSet folder or test program !!!, it will use mapping: {1}", row.Timeset, row.TimesetMapping.First())
                        //});
                    }
                    else
                    {
                        //Errors.AddRange(mappingErrs);
                        //Errors.Add(planErr);
                    }
                }
                else
                {
                    row.TimesetMapping.Insert(0, row.Timeset);
                }
            }
        }

        internal List<Error> CheckMappingTimeset(AiTestPlanRow row)
        {
            var subError = new List<Error>();
            if (row.TimesetMapping.Count > 1)
            {
                AddError(AutoAiErrorType.W_CanNotDetermineWhichSpecToUse_01, SheetName, row.RowNum, row.IndexMappingPattern, [row.MappingPattern, string.Join(",", row.TimesetMapping)]);
                ErrorReportManager.AddError(AutoAiErrorType.W_CanNotDetermineWhichSpecToUse_01, SheetName, row.RowNum, row.IndexMappingPattern, [row.MappingPattern, string.Join(",", row.TimesetMapping)]);
            }
            if (row.TimesetMapping.Count < 1)
            {
                AddError(AutoAiErrorType.E_CanNotDetermineWhichSpecToUse_02, SheetName, row.RowNum, row.IndexMappingPattern, [row.MappingPattern]);
                ErrorReportManager.AddError(AutoAiErrorType.E_CanNotDetermineWhichSpecToUse_02, SheetName, row.RowNum, row.IndexMappingPattern, [row.MappingPattern]);
            }
            return subError;
        }

        internal void CheckMappingDcLevels(AiTestPlanRow row)
        {
            if (row.DcLevelsMapping.Count > 1)
            {
                AddError(AutoAiErrorType.W_CanNotDetermineWhichSpecToUse_03, SheetName, row.RowNum, row.IndexMappingPattern, [row.MappingPattern, string.Join(",", row.DcLevelsMapping)]);
                ErrorReportManager.AddError(AutoAiErrorType.W_CanNotDetermineWhichSpecToUse_03, SheetName, row.RowNum, row.IndexMappingPattern, [row.MappingPattern, string.Join(",", row.DcLevelsMapping)]);
            }
            if (row.DcLevelsMapping.Count < 1)
            {
                AddError(AutoAiErrorType.E_CanNotDetermineWhichSpecToUse_04, SheetName, row.RowNum, row.IndexMappingPattern, [row.MappingPattern]);
                ErrorReportManager.AddError(AutoAiErrorType.E_CanNotDetermineWhichSpecToUse_04, SheetName, row.RowNum, row.IndexMappingPattern, [row.MappingPattern]);
            }
        }

        internal void CheckDcSpec(List<string> dcSpecs)
        {
            foreach (var row in Rows)
            {
                if (!dcSpecs.Contains(row.DcCategory, StringComparer.OrdinalIgnoreCase))
                {
                    //Errors.Add(new Error
                    //{
                    //    EnumErrorType = EnumErrorType.Missing,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexVoltageCategory,
                    //    Message = string.Format("[VoltageCategory] The {0} is not existed in program !!!",
                    //        row.VoltageCategory)
                    //});
                }
                else
                {
                    var mappingDc = row.DcLevelsMapping.Where(x => string.Equals(row.DcCategory, x.Split(';')[0], StringComparison.OrdinalIgnoreCase)).ToList();
                    if (mappingDc == null || mappingDc.Count() == 0)
                    {
                        //Errors.Add(new Error
                        //{
                        //    EnumErrorType = EnumErrorType.MissingDcInMappingDcAndLevels,
                        //    ErrorLevel = ErrorLevel.Warning,
                        //    SheetName = Name,
                        //    RowNum = row.RowNum,
                        //    ColNum = IndexVoltageCategory,
                        //    Message = string.Format("[VoltageCategory] {0} is not used for payload ({1}) in program !!!",
                        //    row.VoltageCategory, row.MappingPattern)
                        //});
                        row.DcLevelsMapping = new List<string>();
                    }
                    else
                    {
                        row.DcLevelsMapping = mappingDc;
                        CheckMappingDcLevels(row);
                    }
                }
            }
        }

        internal void CheckAcSpec(List<string> acSpecs)
        {
            foreach (var row in Rows)
            {
                var mappingErrs = CheckMappingAc(row);
                //var planErr = new Error
                //{
                //    EnumErrorType = EnumErrorType.Missing,
                //    ErrorLevel = ErrorLevel.Error,
                //    SheetName = Name,
                //    RowNum = row.RowNum,
                //    ColNum = IndexAcCategory,
                //    Message = string.Format(
                //                "[AcCategory] {0} in test plan can’t be found from base program, please double check the AC category naming.", row.AcCategory)
                //};

                if (!acSpecs.Any(x => x.ToUpper().Equals(row.AcCategory.ToUpper())))
                {
                    var planAcMessage = string.IsNullOrEmpty(row.AcCategory) ?
                        "[AcCategory] AC is blank in AI test plan" :
                        "[AcCategory] " + row.AcCategory + " in test plan can’t be found from base program";

                    if (!mappingErrs.Any())
                    {
                        //Errors.Add(new Error()
                        //{
                        //    EnumErrorType = EnumErrorType.Missing,
                        //    ErrorLevel = ErrorLevel.Warning,
                        //    SheetName = Name,
                        //    RowNum = row.RowNum,
                        //    ColNum = IndexAcCategory,
                        //    Message = string.Format(
                        //        "{0}, it will use mapping: {1} .", planAcMessage, row.AcCategoryMapping.First())
                        //});
                    }
                    else
                    {
                        //Errors.AddRange(mappingErrs);
                        //Errors.Add(planErr);
                    }
                }
                else
                {
                    row.AcCategoryMapping.Insert(0, row.AcCategory);
                }
            }
        }

        internal List<Error> CheckMappingAc(AiTestPlanRow row)
        {
            var subError = new List<Error>();
            if (row.AcCategoryMapping.Count > 1)
            {
                //subError.Add(new Error
                //{
                //    EnumErrorType = EnumErrorType.CanNotDetermineWhichSpecToUse,
                //    ErrorLevel = ErrorLevel.Warning,
                //    SheetName = Name,
                //    RowNum = row.RowNum,
                //    ColNum = row.IndexMappingPattern,
                //    Message = string.Format("[Mapping] Multi ac category for payload({0}) in base program: {1} .", row.MappingPattern, string.Join(",", row.AcCategoryMapping)),
                //    Comments = new List<string> { String.Join(",", row.AcCategoryMapping) }
                //});
            }
            if (row.AcCategoryMapping.Count < 1)
            {
                //subError.Add(new Error
                //{
                //    EnumErrorType = EnumErrorType.CanNotDetermineWhichSpecToUse,
                //    ErrorLevel = ErrorLevel.Error,
                //    SheetName = Name,
                //    RowNum = row.RowNum,
                //    ColNum = row.IndexMappingPattern,
                //    Message = string.Format("[Mapping] None of ac category for payload({0}) in base program.", row.MappingPattern)
                //});
            }
            return subError;
        }

        internal void CheckPattern(HashSet<string> patterns, PatternListSheet patternListSheet)
        {
            foreach (var row in Rows)
            {
                if (row.Payloads == null || row.Payloads.Count == 0)
                {
                    //Errors.Add(new Error
                    //{
                    //    EnumErrorType = EnumErrorType.MissingPatternInDashboard,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = 1,
                    //    Message = string.Format(
                    //    "[Pattern] The Instance \"{0}\" not existed any payload pattern !!!", row.Parameter)
                    //});
                }

                for (var i = 0; i < row.Patterns.Count; i++)
                {
                    var pattern = row.Patterns[i];
                    var patternInDashboad = patternListSheet.Rows.FirstOrDefault(x => x.Pattern.Equals(pattern.OriName, StringComparison.CurrentCultureIgnoreCase));
                    if (patternInDashboad == null)
                    {
                        //Errors.Add(new Error
                        //{
                        //    EnumErrorType = EnumErrorType.MissingPatternInDashboard,
                        //    ErrorLevel = ErrorLevel.Error,
                        //    SheetName = Name,
                        //    RowNum = row.RowNum,
                        //    ColNum = IndexPatternStart + i,
                        //    Message = string.Format(
                        //         "[Pattern] The pattern \"{0}\" not existed in pattern dashboard !!!", pattern.OriName)
                        //});
                    }
                    else
                    {
                        if (!patterns.Contains(patternInDashboad.PatternVersion, StringComparer.OrdinalIgnoreCase))
                        {
                            //Errors.Add(new Error
                            //{
                            //    EnumErrorType = EnumErrorType.MissingPatternFile,
                            //    ErrorLevel = ErrorLevel.Error,
                            //    SheetName = Name,
                            //    RowNum = row.RowNum,
                            //    ColNum = IndexPatternStart + i,
                            //    Message = string.Format(
                            //        "[Pattern] The pattern \"{0}\" not existed in pattern folder !!!", pattern.OriName)
                            //});
                        }
                    }
                }
            }
        }

        internal void CheckPins(List<string> pins, List<string> acSymbols)
        {
            var pinsWithoutUnderline = new Dictionary<string, string>();
            pins.ForEach(x => pinsWithoutUnderline[x.Replace("_", "").ToUpper()] = x.ToUpper());
            acSymbols.ForEach(x => pinsWithoutUnderline[x.Replace("_", "").ToUpper()] = x.ToUpper());
            foreach (var row in Rows)
            {
                foreach (var pin in row.Pins)
                {
                    pin.Name = pin.Name.ToUpper();
                    if (pinsWithoutUnderline.ContainsKey(pin.Name))
                    {
                        pin.Name = pinsWithoutUnderline[pin.Name];
                    }
                }
            }
            var firstRow = Rows.First();
            for (var i = 0; i < firstRow.Pins.Count; i++)
            {
                var pin = firstRow.Pins[i];
                if (!pins.Contains(pin.Name, StringComparer.OrdinalIgnoreCase) &&
                    !acSymbols.Contains(pin.Name, StringComparer.OrdinalIgnoreCase))
                {
                    //Errors.Add(new Error
                    //{
                    //    EnumErrorType = EnumErrorType.Missing,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = IndexStartRow - 1,
                    //    ColNum = pin.IndexStart,
                    //    Message = string.Format("[Pin] The pin \"{0}\" not existed in test program !!!", pin.Name)
                    //});
                }
            }

            foreach (var row in Rows)
            {
                for (var i = 0; i < row.Pins.Count; i++)
                {
                    var pin = row.Pins[i];

                    if (pins.Contains(pin.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        ValidateVoltagePin(row, pin);
                    }
                    else if (acSymbols.Contains(pin.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        ValidateAcSymbolPin(row, pin);
                    }
                }
            }
        }

        private void ValidateVoltagePin(AiTestPlanRow row, Pin pin)
        {
            string value;
            if (!string.IsNullOrEmpty(pin.Start) && !pin.Start.TryConvertToVolt(out value))
            {
                //Errors.Add(new Error
                //{
                //    EnumErrorType = EnumErrorType.FormatError,
                //    ErrorLevel = ErrorLevel.Error,
                //    SheetName = Name,
                //    RowNum = row.RowNum,
                //    ColNum = pin.IndexStart,
                //    Message = string.Format("[Pin] The syntax of value \"{0}\" is not correct !!!",
                //        pin.Start)
                //});
            }

            if (!string.IsNullOrEmpty(pin.Stop) && !pin.Stop.TryConvertToVolt(out value))
            {
                //Errors.Add(new Error
                //{
                //    EnumErrorType = EnumErrorType.FormatError,
                //    ErrorLevel = ErrorLevel.Error,
                //    SheetName = Name,
                //    RowNum = row.RowNum,
                //    ColNum = pin.IndexStop,
                //    Message = string.Format("[Pin] The syntax of value \"{0}\" is not correct !!!",
                //        pin.Stop)
                //});
            }

            if (!string.IsNullOrEmpty(pin.Step) && !pin.Step.TryConvertToVolt(out value))
            {
                //Errors.Add(new Error
                //{
                //    EnumErrorType = EnumErrorType.FormatError,
                //    ErrorLevel = ErrorLevel.Error,
                //    SheetName = Name,
                //    RowNum = row.RowNum,
                //    ColNum = pin.IndexStep,
                //    Message = string.Format("[Pin] The syntax of value \"{0}\" is not correct !!!",
                //        pin.Step)
                //});
            }

            decimal startVar;
            decimal stopVar;
            if (decimal.TryParse(pin.Start, out startVar))
            {
                if (startVar >= 1.3m)
                {
                    //Errors.Add(new Error
                    //{
                    //    EnumErrorType = EnumErrorType.VoltageHigherThan1p3,
                    //    ErrorLevel = ErrorLevel.Warning,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = pin.IndexStart,
                    //    Message = string.Format("[Pin] Shmoo start value {0} >= 1.3 !!!",
                    //    pin.Start)
                    //});
                }
            }
            if (decimal.TryParse(pin.Stop, out stopVar))
            {
                if (stopVar >= 1.3m)
                {
                    //Errors.Add(new Error
                    //{
                    //    EnumErrorType = EnumErrorType.VoltageHigherThan1p3,
                    //    ErrorLevel = ErrorLevel.Warning,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = pin.IndexStop,
                    //    Message = string.Format("[Pin] Shmoo stop value {0} >= 1.3 !!!",
                    //    pin.Stop)
                    //});
                }
            }
            if (decimal.TryParse(pin.Start, out startVar) && decimal.TryParse(pin.Stop, out stopVar))
            {
                if (startVar < stopVar)
                {
                    //Errors.Add(new Error
                    //{
                    //    EnumErrorType = EnumErrorType.VddShmooLowToHigh,
                    //    ErrorLevel = ErrorLevel.Warning,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = pin.IndexStart,
                    //    Message = string.Format("[Pin] Shmoo range is from low to high, {0} to {1} !!!",
                    //    pin.Start, pin.Stop)
                    //});
                }
            }
            pin.Type = "Pin";
        }

        private void ValidateAcSymbolPin(AiTestPlanRow row, Pin pin)
        {
            string value;
            if (!string.IsNullOrEmpty(pin.Start) && !pin.Start.TryConvertToFreq(out value))
            {
                //Errors.Add(new Error
                //{
                //    EnumErrorType = EnumErrorType.FormatError,
                //    ErrorLevel = ErrorLevel.Error,
                //    SheetName = Name,
                //    RowNum = row.RowNum,
                //    ColNum = pin.IndexStart,
                //    Message = string.Format("[Pin] The syntax of value \"{0}\" is not correct !!!",
                //        pin.Start)
                //});
            }

            if (!string.IsNullOrEmpty(pin.Stop) && !pin.Stop.TryConvertToFreq(out value))
            {
                //Errors.Add(new Error
                //{
                //    EnumErrorType = EnumErrorType.FormatError,
                //    ErrorLevel = ErrorLevel.Error,
                //    SheetName = Name,
                //    RowNum = row.RowNum,
                //    ColNum = pin.IndexStop,
                //    Message = string.Format("[Pin] The syntax of value \"{0}\" is not correct !!!",
                //        pin.Stop)
                //});
            }

            if (!string.IsNullOrEmpty(pin.Step) && !pin.Step.TryConvertToFreq(out value))
            {
                //Errors.Add(new Error
                //{
                //    EnumErrorType = EnumErrorType.FormatError,
                //    ErrorLevel = ErrorLevel.Error,
                //    SheetName = Name,
                //    RowNum = row.RowNum,
                //    ColNum = pin.IndexStep,
                //    Message = string.Format("[Pin] The syntax of value \"{0}\" is not correct !!!",
                //        pin.Step)
                //});
            }
            pin.Type = "AcSymbol";
        }

        internal void CheckRetention()
        {
            //TODO
        }

        internal void CheckDigSrc(string patternFolder, List<HardIpReference> patInfoData)
        {
            if (!patInfoData.Any())
            {
                //var error = new Error
                //{
                //    EnumErrorType = EnumErrorType.MissingPatternInfoFile,
                //    ErrorLevel = ErrorLevel.Warning,
                //    Message = string.Format("Can't get any pattern info from the path: ", patternFolder)
                //};
                //Errors.Add(error);
            }

            foreach (var row in Rows)
            {
                CheckSelsramForDigSrc(row, patInfoData);
                if (IndexDigSrc != -1)
                    CheckDigSrcByRow(row, patInfoData);
            }
        }
        private void CheckSelsramForDigSrc(AiTestPlanRow row, List<HardIpReference> patInfoData)
        {
            foreach (var pattern in row.Patterns)
            {
                if (pattern.Name.ToUpper().Contains("_SRMDSSC"))
                {
                    HardIpReference targetPatInfo = null;
                    if (pattern.Version == "")
                        targetPatInfo = patInfoData.LastOrDefault(x => x.Payload.ToUpper() == pattern.Name.ToUpper());
                    else
                        targetPatInfo = patInfoData.FirstOrDefault(x => x.Payload.ToUpper() == pattern.Name.ToUpper() && x.Version.ToUpper() == pattern.Version.ToUpper());
                    if (targetPatInfo == null)
                    {
                        //var error = new Error
                        //{
                        //    EnumErrorType = EnumErrorType.MissingSelsramSendBitInPatternInfo,
                        //    ErrorLevel = ErrorLevel.Warning,
                        //    SheetName = Name,
                        //    RowNum = row.RowNum,
                        //    ColNum = IndexSelsramDssc,
                        //    Message = string.Format($"Can't get selsram pattern: {pattern.Name}, version: {pattern.Version} in pattern info."),
                        //};
                        //Errors.Add(error);
                        continue;
                    }
                    if (string.IsNullOrEmpty(targetPatInfo.SendBitStr))
                    {
                        //var error = new Error
                        //{
                        //    EnumErrorType = EnumErrorType.MissingSelsramSendBitInPatternInfo,
                        //    ErrorLevel = ErrorLevel.Warning,
                        //    SheetName = Name,
                        //    RowNum = row.RowNum,
                        //    ColNum = IndexSelsramDssc,
                        //    Message = string.Format($"Can't get send bit infomation in pattern info for selsram pattern: {pattern.Name}."),
                        //};
                        //Errors.Add(error);
                        continue;
                    }


                    var sgmtList = targetPatInfo.SendBitStr.Split('+').ToList();
                    foreach (var sgmt in sgmtList)
                    {
                        var sgmtRegex = Regex.Match(sgmt, @"(?<header>sgmt\d+)_(?<length>\d+)", RegexOptions.IgnoreCase);
                        if (!sgmtRegex.Success)
                            continue;
                    }
                }
            }
        }
        private void CheckDigSrcByRow(AiTestPlanRow row, List<HardIpReference> patInfoData)
        {
            if (string.IsNullOrEmpty(row.DigSrc.Trim()))
                return;
            var splitByPat = row.DigSrc.Split(',');
            foreach (var split in splitByPat)
            {
                var settings = split.Split(':');
                if (settings.Length < 2)
                {
                    //var error = new Error
                    //{
                    //    EnumErrorType = EnumErrorType.WrongDigSrcFormat,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexDigSrc,
                    //    Message = string.Format($"Can't get pattern index and bits setting that split by \":\", content: {split}."),
                    //};
                    //Errors.Add(error);
                    continue;
                }

                var patSetting = Regex.Match(settings[0].Trim(), @"^(Pat|Pattern)(?<index>\d)$+", RegexOptions.IgnoreCase);
                var bitSetting = settings[1].Trim();
                if (!patSetting.Success)
                {
                    //var error = new Error
                    //{
                    //    EnumErrorType = EnumErrorType.WrongDigSrcFormat,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexDigSrc,
                    //    Message = string.Format($"Wrong pattern header format: {settings[0]}."),
                    //};
                    //Errors.Add(error);
                    continue;
                }

                if (string.IsNullOrEmpty(bitSetting))
                {
                    //var error = new Error
                    //{
                    //    EnumErrorType = EnumErrorType.WrongDigSrcFormat,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexDigSrc,
                    //    Message = string.Format($"Bits setting is empty: {split}."),
                    //};
                    //Errors.Add(error);
                    continue;
                }

                var patIndex = patSetting.Groups["index"].Value;
                var targetPat = row.Patterns.FirstOrDefault(x => x.Index == patIndex);
                if (targetPat == null)
                {
                    //var error = new Error
                    //{
                    //    EnumErrorType = EnumErrorType.WrongDigSrcFormat,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexDigSrc,
                    //    Message = string.Format($"Can't find the pattern field: {settings[0].Trim()}."),
                    //};
                    //Errors.Add(error);
                    continue;
                }

                HardIpReference targetPatInfo = null;
                if (targetPat.Version == "")
                    targetPatInfo = patInfoData.LastOrDefault(x => x.Payload.ToUpper() == targetPat.Name.ToUpper());
                else
                    targetPatInfo = patInfoData.FirstOrDefault(x => x.Payload.ToUpper() == targetPat.Name.ToUpper() && x.Version.ToUpper() == targetPat.Version.ToUpper());
                if (targetPatInfo == null)
                {
                    //var error = new Error
                    //{
                    //    EnumErrorType = EnumErrorType.WrongDigSrcFormat,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexDigSrc,
                    //    Message = string.Format($"Can't find the pattern info of {settings[0].Trim()} : {targetPat.Name.ToUpper()}, version: {targetPat.Version.ToUpper()}."),
                    //};
                    //Errors.Add(error);
                    continue;
                }
                if (targetPat.SelsramDigSrc)
                {
                    //var error = new Error
                    //{
                    //    EnumErrorType = EnumErrorType.OverwriteSelsramByDigSrcDefinition,
                    //    ErrorLevel = ErrorLevel.Warning,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexDigSrc,
                    //    Message = string.Format($"Overwrite SELSRAM by {split} for {targetPat.Name.ToUpper()}."),
                    //};
                    //Errors.Add(error);
                }
                CheckDigSrcForPatCell(row, bitSetting, targetPat, targetPatInfo);
            }
        }

        private void CheckDigSrcForPatCell(AiTestPlanRow row, string bitSetting, PatternDate pat, HardIpReference patternInfo)
        {
            if (Regex.IsMatch(bitSetting, @"^[0|1]+$"))
                CheckDigSrcSegByBits(row, pat, bitSetting, patternInfo);
            else
                CheckDigSrcSegBySgmts(row, pat, bitSetting, patternInfo);
        }

        private void CheckDigSrcSegByBits(AiTestPlanRow row, PatternDate pat, string bits, HardIpReference patternInfo)
        {
            var sgmtList = patternInfo.SendBitStr.Split('+').ToList();
            var unsetBit = bits;
            var resultList = new List<string>();
            foreach (var sgmt in sgmtList)
            {
                var sgmtRegex = Regex.Match(sgmt, @"(?<header>sgmt\d+)_(?<length>\d+)", RegexOptions.IgnoreCase);
                if (!sgmtRegex.Success)
                {
                    //var error = new Error
                    //{
                    //    EnumErrorType = EnumErrorType.WrongDigSrcFormat,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexDigSrc,
                    //    Message = string.Format($"Wrong Send Bit Str in pattern info, pattern: {pat.Name}."),
                    //};
                    //Errors.Add(error);
                    continue;
                }

                var header = sgmtRegex.Groups["header"].Value;
                var length = int.Parse(sgmtRegex.Groups["length"].Value);
                if (unsetBit.Length < length)
                {
                    if (unsetBit.Length == 0)
                        break;
                    resultList.Add(string.Format($"{header}=0b{unsetBit.PadLeft(length, '0')}."));
                    break;
                }
                else
                {
                    resultList.Add(string.Format($"{header}=0b{unsetBit.Substring(0, length)}."));
                    unsetBit = unsetBit.Substring(length);
                    continue;
                }
            }
        }

        private void CheckDigSrcSegBySgmts(AiTestPlanRow row, PatternDate pat, string sgmtStr, HardIpReference patternInfo)
        {
            var sgmtList = patternInfo.SendBitStr.Split('+').ToList();
            var sgmtSettingList = sgmtStr.Split(';').ToList();
            var resultList = new List<string>();
            var bitsDict = sgmtList.ToDictionary(x => x, x => "");
            foreach (var sgmtSetting in sgmtSettingList)
            {
                var settingRegex = Regex.Match(sgmtStr, @"^(?<header>sgmt\d+)=0b(?<bits>[0|1]+)$");
                if (!settingRegex.Success)
                {
                    //var error = new Error
                    //{
                    //    EnumErrorType = EnumErrorType.WrongDigSrcFormat,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexDigSrc,
                    //    Message = string.Format($"Wrong sgmt format definition on {sgmtStr}, it must be like \"sgmt\\d+=0b[0|1]+\"."),
                    //};
                    //Errors.Add(error);
                    continue;
                }
                var header = settingRegex.Groups["header"].Value;
                var bits = settingRegex.Groups["bits"].Value;

                var patInfoSgmt = sgmtList.FirstOrDefault(x => x.StartsWith(header + "_", StringComparison.OrdinalIgnoreCase));
                if (string.IsNullOrEmpty(patInfoSgmt))
                {
                    //var error = new Error
                    //{
                    //    EnumErrorType = EnumErrorType.WrongDigSrcFormat,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexDigSrc,
                    //    Message = string.Format($"Can't find {header} in Send bit Str in pattern info, pattern: {pat.Name}."),
                    //};
                    //Errors.Add(error);
                    continue;
                }

                var sgmtRegex = Regex.Match(patInfoSgmt, @"(?<header>sgmt\d+)_(?<length>\d+)", RegexOptions.IgnoreCase);
                if (!sgmtRegex.Success)
                {
                    //var error = new Error
                    //{
                    //    EnumErrorType = EnumErrorType.WrongDigSrcFormat,
                    //    ErrorLevel = ErrorLevel.Error,
                    //    SheetName = Name,
                    //    RowNum = row.RowNum,
                    //    ColNum = IndexDigSrc,
                    //    Message = string.Format($"{header} lacks length infomation in Send bit Str in pattern info, pattern: {pat.Name}."),
                    //};
                    //Errors.Add(error);
                    continue;
                }
            }
        }
    }
}
