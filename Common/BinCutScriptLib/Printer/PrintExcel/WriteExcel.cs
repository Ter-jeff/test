using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

using CommonLib.Datalog;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;

using TestPlanLib;
using TestPlanLib.BinCut;
using TestPlanLib.BinCut.BinCutConfig;
using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.Flow;

namespace BinCutScriptLib.Printer.PrintExcel
{
    internal class WriteExcel
    {
        public class PatSteps
        {
            public string PatName = string.Empty;
            public List<SiteInfo> SiteInfos = [];
            public int[] StepCount = [];
            public List<CpInfo> CpCount = [];
            public List<PatCp> PatCpCount = [];
            public int AllPwrCount;
            public int PwrIdx;
            public void InitStepCount()
            {
                StepCount = new int[AllPwrCount];
                for (int i = 0; i < AllPwrCount; i++)
                {
                    StepCount[i] = 0;
                }
            }
        }

        public class CpInfo
        {
            public double CpValue;
            public int Count;
        }

        public class PatCp
        {
            public string ItemName = string.Empty;
            public int Count;
        }

        public class SiteInfo
        {
            public int Idx;
            public int Step;
            public double Cp;
        }

        private const string PatternSummary = "PatternSummary";

        public object MisValue = Type.Missing;
        public ExcelPackage XlPackage;
        public ExcelWorkbook XlWorkBook;
        public string DestExcel;
        public BinningTables BinningTables;
        private readonly List<Base.SiteInfo> _allDiceInfos;
        private readonly List<string> _powerNames;
        public List<string> InstNameList;

        //CP1
        public WriteExcel(string outPath, List<List<Base.SiteInfo>> allDiceInfos, BinningTables binningTables)
        {
            string dirName = Path.GetDirectoryName(outPath)!;
            string fileName = Path.GetFileNameWithoutExtension(outPath);
            DestExcel = Path.Combine(dirName, fileName + ".xlsx");
            BinningTables = binningTables;
            _allDiceInfos = [.. allDiceInfos.SelectMany(x => x)];
            InstNameList = [.. allDiceInfos.SelectMany(x => x).SelectMany(x => x.InstanceList).Select(x => x.InstanceName).Distinct()];
            XlPackage = new ExcelPackage();
            XlWorkBook = XlPackage.Workbook;
            _powerNames = [.. _allDiceInfos.SelectMany(x => x.AllPowers.Select(y => y.PinMode)).Distinct()];
        }

        public void WriteMissExtraCheck(Job job, BinCutPatternReport binCutPatternReport, List<BinCutFlowTable> binCutFlowTables, string tempFolder)
        {
            string binCutPatternSummary = Path.Combine(tempFolder, PatternSummary + ".txt");
            binCutPatternReport.WriteTxt(binCutPatternSummary, []);
            XlWorkBook.AddTxt(binCutPatternSummary);

            List<BinCutPatternRow> binCutPatternReportRows = GetEnv(job, binCutPatternReport, tempFolder);

            PatternCheckMiss(job, binCutFlowTables, binCutPatternReportRows, out List<string> modes);

            PatternCheckExtra(modes);
        }

        private static List<BinCutPatternRow> GetEnv(Job job, BinCutPatternReport binCutPatternReport, string tempFolder)
        {
            #region Get Enable and env from T/P
            var binCutPatternReportRows = binCutPatternReport.Rows.Where(y =>
            {
                string? jobTestStage = y.BinCutInstanceRow?.JobTestStage;
                return y.Jobs.Exists(x => x.EqualsIgnoreCase(job.JobName)) &&
                    (string.IsNullOrEmpty(jobTestStage) || jobTestStage.EqualsIgnoreCase("All") || jobTestStage.Split([',', ' '], StringSplitOptions.RemoveEmptyEntries).Any(x => x.EqualsIgnoreCase(job.JobName)));
            }).ToList();
            IEnumerable<IGrouping<string, BinCutPatternRow>> groups = binCutPatternReportRows.GroupBy(x => x.FlowName);
            foreach (IGrouping<string, BinCutPatternRow> group in groups)
            {
                var readFlowSheet = new ReadFlowSheet();
                List<string> fileList = [.. Directory.GetFiles(tempFolder, group.Key + ".txt")];
                if (fileList.Count != 0)
                {
                    SubFlowSheet flowSheet = readFlowSheet.GetSheet(fileList.First());
                    foreach (BinCutPatternRow row in group)
                    {
                        if (flowSheet.Rows.Exists(x => x.Parameter.EqualsIgnoreCase(row.InstanceName) &&
                                     x.Opcode.EqualsIgnoreCase("Test")))
                        {
                            FlowRow flowRow = flowSheet.Rows.Find(x => x.Parameter.EqualsIgnoreCase(row.InstanceName) &&
                                                                       x.Opcode.EqualsIgnoreCase("Test"))!;
                            row.Env = flowRow.Env;
                            row.Enbable = flowRow.Enable;
                        }
                    }
                }
            }
            #endregion
            return binCutPatternReportRows;
        }

        private void PatternCheckMiss(Job job, List<BinCutFlowTable> binCutFlowTables, List<BinCutPatternRow> binCutPatternRows, out List<string> modes)
        {
            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("PatternCheck_DataLogMiss");
            BinCutFlowTable flow = binCutFlowTables.Find(x => new Job(x.JobName).JobType.Equals(job.JobType))!;
            modes = [.. flow.Rows.Select(x => x.PerformanceMode).OrderByDescending(x => x.Length)];
            var testPlanRowGroup = binCutPatternRows.ChunkBy(x => x.InstanceName).ToList();

            PrintHeader(out int currRow, out string title, xlWorkSheet, out int titleCount, out int patRow);

            SetKeyInLog(modes);

            PrintData(modes, ref currRow, xlWorkSheet, testPlanRowGroup, titleCount, ref patRow);

            xlWorkSheet.Cells[1, 11].TryAutoFitColumns();
            xlWorkSheet.Cells[1, 1, 1, titleCount].AutoFilter = true;
            xlWorkSheet.View.FreezePanes(2, 1);
            IExcelConditionalFormattingEqual cond1 = xlWorkSheet.ConditionalFormatting.AddEqual(new ExcelAddress("A:A"));
            cond1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cond1.Style.Fill.BackgroundColor.Color = Color.Red;
            cond1.Style.Font.Color.Color = Color.White;
            cond1.Formula = "\"No\"";
        }

        private void PrintData(List<string> modes, ref int currRow, ExcelWorksheet excelWorksheet, List<IGrouping<string, BinCutPatternRow>> testPlanRowGroup, int titleCount, ref int patRow)
        {
            foreach (Base.SiteInfo allDiceInfo in _allDiceInfos)
            {
                if (!allDiceInfo.IsBinCutConfig)
                {
                    continue;
                }

                if (allDiceInfo.AllPowers.Count == 0)
                {
                    continue;
                }

                var datalogSite = allDiceInfo.InstanceList.Select(x => x.InDatalogKey).ToList();
                foreach (IGrouping<string, BinCutPatternRow> rows in testPlanRowGroup)
                {
                    BinCutPatternRow firstRow = rows.First();
                    InstanceData? instacne = allDiceInfo.InstanceList.Find(x => x.InstanceName.EqualsIgnoreCase(firstRow.InstanceName));
                    if (instacne != null)
                    {
                        if (!BinCutConfig.IsDoAll && instacne.IsSearch && instacne.FinalStep == -1)
                        {
                            break;
                        }
                    }

                    if (firstRow.IsInterpolateSkip)
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(firstRow.DeviceCondition))
                    {
                        string deviceCondition = firstRow.DeviceCondition;
                        ReplaceFlagNameByStatus(ref deviceCondition, allDiceInfo.HarvesFlags);
                        bool isRunInstance = ParserExpression(deviceCondition);
                        if (!isRunInstance)
                        {
                            continue;
                        }
                    }

                    string mode = GetMode(rows.First().InstanceName, modes);
                    string type = rows.First().InstanceName.Split('_').Last();
                    List<string> patterns = BinCutData.HasPatList ? [.. rows.Select(x => x.PatternVer.ToUpper())] : [.. rows.Select(x => x.Pattern)];
                    string mergePat = GetMergePat(patterns, rows.First().InstanceName);
                    string str = mode + "-" + mergePat + "-" + type;
                    string str2 = rows.First().Condition + "-" + mergePat;
                    allDiceInfo.InFlowKey.Add(str);
                    if (!datalogSite.Exists(x => x.EqualsIgnoreCase(str)))
                    {
                        List<string> patternList = BinCutData.HasPatList ? [.. rows.Select(x => x.PatternVer.ToUpper())] : [.. rows.Select(x => x.Pattern)];
                        string patlist = rows.First().Condition + "-";
                        mergePat = string.Join("-", patternList);
                        if (string.IsNullOrEmpty(mergePat))
                        {
                            mergePat = rows.First().InstanceName;
                        }

                        if (mergePat.Contains('+'))
                        {
                            mergePat = mergePat.Replace('+', '-');
                        }

                        patlist += mergePat;
                        if (!patlist.EqualsIgnoreCase(str2))
                        {
                            continue;
                        }

                        var flowPatList = rows.Select(x => x.Pattern).ToList();
                        string isOrderError = GetIsOrderError(datalogSite, mode, type, flowPatList);
                        string enable = "";
                        string jobTestStage = "";
                        string sheetName = "";
                        int rowNum = 0;
                        if (firstRow.BinCutInstanceRow != null)
                        {
                            jobTestStage = firstRow.BinCutInstanceRow.JobTestStage;
                            enable = !string.IsNullOrEmpty(firstRow.BinCutInstanceRow.EnableFlow) ? firstRow.BinCutInstanceRow.EnableFlow : firstRow.BinCutInstanceRow.EnableAndDevice;
                            sheetName = firstRow.BinCutInstanceRow.SheetName;
                            rowNum = firstRow.BinCutInstanceRow.RowNum;
                        }
                        int sortbin = allDiceInfo.SortBin;
                        currRow = excelWorksheet.Cells[currRow, 1].PrintExcelRow(new object[] { isOrderError, sheetName, rowNum, allDiceInfo.TuchNum, allDiceInfo.Site, sortbin, allDiceInfo.SortLineNo, type.ToUpper(), mode, firstRow.FlowName, firstRow.Instance, rows.First().InstanceName, enable, rows.First().Env, rows.First().DeviceCondition, jobTestStage });

                        if (patternList.Count != 0)
                        {
                            for (int idx = 0; idx < patternList.Count; idx++)
                            {
                                excelWorksheet.Cells[patRow, titleCount + idx].PrintExcelRow([patternList[idx]]);
                            }
                        }
                        patRow = currRow;
                    }
                }
            }
        }

        private static string GetIsOrderError(List<string> datalogSite, string mode, string type, List<string> flowPatList)
        {
            string isOrderError = "No";
            foreach (string log in datalogSite)
            {
                List<string> logarr = [.. log.Split('-')];
                if (logarr.First().EqualsIgnoreCase(mode) && logarr.Last().EqualsIgnoreCase(type))
                {
                    if (flowPatList.Count == logarr.Count - 2)
                    {
                        isOrderError = "No";
                        var flags = new List<bool>();
                        foreach (string pat in flowPatList)
                        {
                            flags.Add(logarr.Exists(x => x.EqualsIgnoreCase(pat)));
                        }

                        if (flags.TrueForAll(x => x))
                        {
                            isOrderError = "Yes";
                            break;
                        }
                    }
                }
            }

            return isOrderError;
        }

        private void SetKeyInLog(List<string> modes)
        {
            #region set key in log
            foreach (Base.SiteInfo allDiceInfo in _allDiceInfos)
            {
                foreach (InstanceData instance in allDiceInfo.InstanceList)
                {
                    string mode = GetMode(instance.InstanceName, modes);
                    string type = instance.InstanceName.Split('_').Last();
                    List<Tuple<string, int>>? dicPats = BinCutData.HasPatList ? GetPatternVerList(instance.PatternRows) : GetGenericPatternList(instance.PatternRows);
                    if (dicPats != null && dicPats.Count != 0)
                    {
                        var patterns = dicPats.Select(x => x.Item1).ToList();
                        string mergePat = GetMergePat(patterns, instance.InstanceName);
                        string logStr = mode + "-" + mergePat + "-" + type;
                        instance.InDatalogKey = logStr;
                    }
                }
            }
            #endregion
        }

        private static void PrintHeader(out int currRow, out string title, ExcelWorksheet excelWorksheet, out int titleCount, out int patRow)
        {
            //print title
            currRow = 1;
            title = "OrderError,Sheet Name,RowNum,TouchDown,Site,Bin,LineNo,Type,Performance mode,Flow name,Instance,Instance Name,Enable,Env,SiteFlag(per site),Testing Stage,Pattern";
            string[] arr = title.Split(',');
            titleCount = arr.Length;
            currRow = excelWorksheet.Cells[currRow, 1].PrintExcelRow(arr);
            patRow = currRow;
        }

        private void PatternCheckExtra(List<string> modes)
        {
            int currRow;
            string title;
            #region PatternCheck_Extra
            ExcelWorksheet xlWorkSheetextra = XlWorkBook.AddSheet("PatternCheck_DataLogExtra");
            //ptring title
            currRow = 1;
            title = "TouchDown,Site,SortBin,Mode,Instance Name,Pattern Name,Pattern Line";
            currRow = xlWorkSheetextra.Cells[currRow, 1].PrintExcelRow(title.Split(','));

            var extradata = new List<List<object>>();
            foreach (Base.SiteInfo allDiceInfo in _allDiceInfos)
            {
                if (!allDiceInfo.IsBinCutConfig)
                {
                    continue;
                }

                foreach (InstanceData instance in allDiceInfo.InstanceList)
                {
                    if (instance.FinalStep == -1)
                    {
                        break;
                    }

                    string mode = GetMode(instance.InstanceName, modes);
                    string type = instance.InstanceName.Split('_').Last();
                    List<Tuple<string, int>>? dicPats = BinCutData.HasPatList ? GetPatternVerList(instance.PatternRows) : GetGenericPatternList(instance.PatternRows);
                    if (dicPats != null && dicPats.Count != 0)
                    {
                        string mergePat = string.Join("-", dicPats.Select(x => x.Item1));
                        if (string.IsNullOrEmpty(mergePat))
                        {
                            mergePat = instance.InstanceName;
                        }

                        string str = mode + "-" + mergePat + "-" + type;
                        if (!allDiceInfo.InFlowKey.Exists(x => x.EqualsIgnoreCase(str)))
                        {
                            int sortbin = allDiceInfo.SortBin;
                            if (dicPats.Count == 0)
                            {
                                extradata.Add([allDiceInfo.TuchNum, allDiceInfo.Site, sortbin, mode, instance.InstanceName, "", ""]);
                            }
                            else
                            {
                                extradata.AddRange(dicPats.Select(pat => new List<object> { allDiceInfo.TuchNum, allDiceInfo.Site, sortbin, mode, instance.InstanceName, pat.Item1, pat.Item2 }));
                            }
                        }
                    }
                }
            }

            xlWorkSheetextra.Cells[currRow, 1].PrintExcelRowByList(extradata);
            xlWorkSheetextra.MergeColumn(4);
            xlWorkSheetextra.MergeColumn(5);
            xlWorkSheetextra.Cells.TryAutoFitColumns();

            XlWorkBook.Worksheets.MoveAfter("PatternCheck_DataLogExtra", PatternSummary);
            XlWorkBook.Worksheets.MoveAfter("PatternCheck_DataLogMiss", "PatternCheck_DataLogExtra");

            #endregion
        }

        public void WriteErrorReport()
        {
            if (ErrorReportManager.GetErrorList().Count > 0)
            {
                ExcelWorksheet erroExcelWorksheet = XlWorkBook.AddSheet("Error");
                //print title
                int currRow = 1;
                string[] title = ["ErrorType", "ErrorMessage", "SheetName", "RowNum", "ColNum", "Comment"];
                currRow = erroExcelWorksheet.Cells[currRow, 1].PrintExcelRow(title);
                IOrderedEnumerable<Error> errors = ErrorReportManager.GetErrorList().OrderBy(x => x.SheetName).ThenBy(x => x.RowNum).ThenBy(x => x.ColNum);
                //print data
                var data = new List<List<string>>
                {
                    errors.Select(x => x.ErrorLevel.ToString()).ToList(),
                    errors.Select(x => x.Message).ToList(),
                    errors.Select(x => x.SheetName).ToList(),
                    errors.Select(x => x.RowNum.ToString(CultureInfo.InvariantCulture)).ToList(),
                    errors.Select(x => x.ColNum.ToString(CultureInfo.InvariantCulture)).ToList()
                };
                erroExcelWorksheet.Cells[currRow, 1].PrintExcelColByList(data);

                erroExcelWorksheet.Cells.TryAutoFitColumns();

                //Hyperlink for error
                int cnt = 2;
                int colRow = 2;
                foreach (Error item in ErrorReportManager.GetErrorList())
                {
                    if (IsContainSheet(item.SheetName) && item.RowNum != 0 && item.ColNum != 0)
                    {
                        ExcelWorksheet targetSheet = XlWorkBook.Worksheets[item.SheetName];
                        ExcelRange targetRange = targetSheet.Cells[item.RowNum, item.ColNum];
                        targetRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        if (item.ErrorLevel == EnumErrorLevel.Error)
                        {
                            targetRange.Style.Fill.BackgroundColor.SetColor(Color.Red);
                        }
                        else
                        {
                            targetRange.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        }

                        ExcelRange linkRange = erroExcelWorksheet.Cells[cnt, colRow];
                        string cellPos = ExcelCellBase.GetAddress(item.RowNum, item.ColNum);
                        linkRange.Hyperlink = new ExcelHyperLink(targetSheet.Name + "!" + cellPos, linkRange.Text);
                        linkRange.StyleName = "HyperLink";
                        linkRange.Style.Font.Color.SetColor(Color.Blue);
                        linkRange.Style.Font.UnderLine = true;
                    }
                    cnt++;
                }
            }
        }

        public void WriteResult()
        {
            if (InstNameList.Count == 0)
            {
                return;
            }

            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("Result");

            //print title
            xlWorkSheet.Cells[1, 1].Value = "Chip Number";
            xlWorkSheet.Cells[2, 1].Value = "Bin Number";
            int currRow = 1;
            string[] title = [.. _allDiceInfos.Select(x => $"X:{x.XCoor} Y:{x.YCoor}")];
            currRow = xlWorkSheet.Cells[currRow, 2].PrintExcelRow(title);
            string[] subtitle = [.. _allDiceInfos.Select(x => $"Sort:{x.Sort} Bin:{x.SortBin}")];
            currRow = xlWorkSheet.Cells[currRow, 2].PrintExcelRow(subtitle);

            int dataRow = currRow;
            //print data
            int instCnt = InstNameList.Count;
            int diceCnt = _allDiceInfos.Count;
            string[,] array = new string[instCnt, diceCnt];
            for (int i = 0; i < array.GetLength(0); i++)
            {
                for (int j = 0; j < array.GetLength(1); j++)
                {
                    array[i, j] = "N";
                }
            }

            for (int instIdx = 0; instIdx < InstNameList.Count; instIdx++) //<GpuTd_Soc_PP_INLP_MG111_PLLP_SC_PT41_TDF_COM_AUT_ALLFV_DM_PP_BV>
            {
                for (int diceIdx = 0; diceIdx < _allDiceInfos.Count; diceIdx++)
                {
                    Base.SiteInfo oneDice = _allDiceInfos[diceIdx];
                    InstanceData? instanceData = oneDice.InstanceList.Find(x => x.InstanceName == InstNameList[instIdx]);

                    if (instanceData != null)
                    {
                        if (instanceData.IsCheckPassByInstance)
                        {
                            array[instIdx, diceIdx] = "P";
                        }
                        else
                        {
                            array[instIdx, diceIdx] = "F";
                        }
                    }
                }
            }
            currRow = xlWorkSheet.Cells[currRow, 2].PrintExcelRange(array);
            IExcelConditionalFormattingEqual cond1 = xlWorkSheet.ConditionalFormatting.AddEqual(new ExcelAddress(dataRow, 2, dataRow + InstNameList.Count, _allDiceInfos.Count + 2));
            cond1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cond1.Style.Fill.BackgroundColor.Color = Color.Red;
            cond1.Style.Font.Color.Color = Color.White;
            cond1.Formula = "\"F\"";

            //print item
            xlWorkSheet.Cells[dataRow, 1].PrintExcelCol(InstNameList.ToArray());

            //print check pass dice
            foreach (Base.SiteInfo diceinfo in _allDiceInfos)
            {
                diceinfo.CheckResult.CheckPass = diceinfo.InstanceList.All(x => x.IsCheckPassByInstance) ? "P" : "F";
                if (diceinfo.InstanceList.Count == 0)
                {
                    diceinfo.CheckResult.CheckPass = "N";
                }
            }
            xlWorkSheet.Cells[currRow, 1].Value = "Check Result";
            string[] result = [.. _allDiceInfos.Select(x => x.CheckResult.CheckPass)];
            currRow = xlWorkSheet.Cells[currRow, 2].PrintExcelRow(result);

            currRow += 2;
            xlWorkSheet.Cells[currRow, 1].Value = "Interpolation Check ";
            string[] interpolation = [.. _allDiceInfos.Select(x => x.CheckResult.IsInterpolationPass ? "P" : "F")];
            currRow = xlWorkSheet.Cells[currRow, 2].PrintExcelRow(interpolation);

            //print summary
            int passdice = _allDiceInfos.Count(x => x.AllPowers.Count != 0);
            string[,] summary =
            {
                {"Total Chip Number =", diceCnt.ToString(CultureInfo.InvariantCulture)},
                {"Total Check Pass Number = ",_allDiceInfos.Count(x => x.CheckResult.CheckPass == "P").ToString(CultureInfo.InvariantCulture)},
                {"Total Check Fail Number = ",_allDiceInfos.Count(x => x.CheckResult.CheckPass == "F").ToString(CultureInfo.InvariantCulture)},
                {"Total Non Test Number =  ",_allDiceInfos.Count(x => x.CheckResult.CheckPass == "N").ToString(CultureInfo.InvariantCulture)},
                {"",""},
                {"Total Pass Number =",passdice.ToString(CultureInfo.InvariantCulture)},
                {"Total Fail Number =",(diceCnt-passdice).ToString(CultureInfo.InvariantCulture)}
            };
            currRow += 2;
            xlWorkSheet.Cells[currRow, 1].PrintExcelRange(summary);
            xlWorkSheet.Column(1).Width = 60;
            for (int i = 2; i <= _allDiceInfos.Count + 2; i++)
            {
                xlWorkSheet.Column(i).Width = 15;
            }
        }

        public void WriteHistogram()
        {
            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("Histogram");
            int maxSite = _allDiceInfos.Max(x => x.Site) + 1;

            int currRow = 1;
            currRow = xlWorkSheet.Cells[currRow, 1].PrintExcelRow(["EQ NO."]);

            //print data
            for (int pwrIdx = 0; pwrIdx < _powerNames.Count; pwrIdx++)  //Loop all power, ex: VDD_CPU_MC601->VDD_CPU_MC602->...
            {
                EnumPowerType pwrNameType = BinCutAlgorithmService.GetTypeByPowerName(_powerNames[pwrIdx]);
                if (pwrNameType == EnumPowerType.Others)
                {
                    continue;
                }

                var powerTitle = new List<object>() { _powerNames[pwrIdx], "Total" };

                for (int site = 0; site < maxSite; site++)
                {
                    powerTitle.Add("Site " + site);
                }

                xlWorkSheet.Cells[currRow++, 1].PrintExcelRow(powerTitle.ToArray());

                #region one mode
                for (int binIdx = 0; binIdx < BinningTables.Count; binIdx++)
                {
                    BinningTable binningTable = BinningTables[binIdx];
                    foreach (BinningRow row in binningTable.Rows)
                    {
                        if (_powerNames[pwrIdx].Contains(row.RowData[binningTable.ModeIdx]))  //ex: MC601
                        {
                            string eqName = row.RowData[binningTable.EqnIdx];
                            int eqNo = int.Parse(eqName[1..]);
                            int[] eqCntAry = new int[maxSite];
                            foreach (Base.SiteInfo oneDice in _allDiceInfos)
                            {
                                if (oneDice.AllPowers.Count == 0)
                                {
                                    continue;
                                }

                                PowerZone pwrRef = oneDice.AllPowers[pwrIdx];
                                if (pwrRef.SearchStatus != EnumSearchStatus.Search)
                                {
                                    continue;
                                }

                                if (pwrRef.PossibleSteps[pwrRef.FinalStep].EqName == eqNo && oneDice.Bin == binIdx + 1)
                                {
                                    eqCntAry[oneDice.Site]++;
                                }
                            }

                            var array = new List<object>
                            {
                                $"BinCut{binIdx + 1} EQ{eqNo}",
                                eqCntAry.Sum()
                            };
                            for (int site = 0; site < maxSite; site++)
                            {
                                array.Add(eqCntAry[site]);
                            }

                            xlWorkSheet.Cells[currRow++, 1].PrintExcelRow(array.ToArray());
                        }
                    }
                }
                #endregion
            }
            xlWorkSheet.Cells.TryAutoFitColumns();
        }

        public void WritePassList()
        {
            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("PassList");

            //print title
            int currRow = 1;
            currRow = xlWorkSheet.Cells[currRow, 1].PrintExcelRow(["Soft", "Bin", "ECID", "Site", "PM", "IDS", "EQ", "Lvcc", "Product"]);

            var powerNames = new List<string>();
            for (int idx = 0; idx < InstNameList.Count; idx++)
            {
                if (!InstNameList[idx].EndsWith("_BV"))
                {
                    continue;
                }

                if (_allDiceInfos.SelectMany(x => x.InstanceList).ToList().Exists(y => y.InstanceName == InstNameList[idx]))
                {
                    int itemPwrIdx = _allDiceInfos.SelectMany(x => x.InstanceList).First(y => y.InstanceName == InstNameList[idx]).PowersIdx;
                    powerNames.Add(_powerNames[itemPwrIdx]);
                }
            }
            powerNames = [.. powerNames.Distinct()];

            //print data
            foreach (Base.SiteInfo oneDice in _allDiceInfos)
            {
                if (oneDice.AllPowers.Count == 0)
                {
                    continue;
                }

                foreach (string powerName in powerNames)
                {

                    PowerZone? pwrRef = oneDice.AllPowers.Find(x => x.PinMode.EqualsIgnoreCase(powerName));
                    if (pwrRef == null)
                    {
                        continue;
                    }

                    EnumPowerType pwrNameType = BinCutAlgorithmService.GetTypeByPowerName(pwrRef.PinMode);
                    if (pwrNameType == EnumPowerType.Others)
                    {
                        continue;
                    }

                    int finalStep = pwrRef.FinalStep;
                    if (pwrRef.IsBinOut)
                    {
                        var array = new List<object>
                        {
                            oneDice.Sort,
                            oneDice.Bin,
                            $"X:{oneDice.XCoor} Y:{oneDice.YCoor}",
                            oneDice.Site,
                            pwrRef.PinMode,
                            -1,
                            -1,
                            -1,
                            -1,
                        };
                        xlWorkSheet.Cells[currRow++, 1].PrintExcelRow(array.ToArray());
                    }
                    else if (finalStep == -1)
                    {
                        var array = new List<object>
                        {
                            oneDice.Sort,
                            oneDice.Bin,
                            $"X:{oneDice.XCoor} Y:{oneDice.YCoor}",
                            oneDice.Site,
                            pwrRef.PinMode,
                            Math.Round(pwrRef.IdsValue,4,MidpointRounding.AwayFromZero),
                            0,
                            0,
                            0,
                        };
                        xlWorkSheet.Cells[currRow++, 1].PrintExcelRow(array.ToArray());
                    }
                    else
                    {
                        var array = new List<object>
                        {
                            oneDice.Sort,
                            oneDice.Bin,
                            $"X:{oneDice.XCoor} Y:{oneDice.YCoor}",
                            oneDice.Site,
                            pwrRef.PinMode,
                            Math.Round(pwrRef.IdsValue,4,MidpointRounding.AwayFromZero),
                            pwrRef.PossibleSteps[finalStep].EqName,
                            Math.Round(pwrRef.PossibleSteps[finalStep].Lvcc,4,MidpointRounding.AwayFromZero),
                            Math.Round(pwrRef.PossibleSteps[finalStep].ProductValue,4,MidpointRounding.AwayFromZero),
                        };
                        xlWorkSheet.Cells[currRow++, 1].PrintExcelRow(array.ToArray());
                    }
                }
            }
            xlWorkSheet.Cells.TryAutoFitColumns();
        }

        public void WriteAllSitePatternFailList()
        {
            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("AllSitePatternFailList");
            int currRow = 1;
            var titles = new List<string> { "TestIns Name", "Test Name", "Bin", "EQN", "Total Fail Site Count" };
            currRow = xlWorkSheet.Cells[currRow, 1].PrintExcelRow(titles.ToArray());
            for (int idx = 0; idx < InstNameList.Count; idx++)
            {
                IEnumerable<string> testNames = _allDiceInfos.SelectMany(x => x.InstanceList).Where(x => x.InstanceName.EqualsIgnoreCase(InstNameList[idx]))
                    .SelectMany(x => x.FailPatternData).Select(x => x.PatternName).Distinct();
                IEnumerable<int> eqnTotal = _allDiceInfos.SelectMany(x => x.InstanceList).Where(x => x.InstanceName.EqualsIgnoreCase(InstNameList[idx]))
                    .SelectMany(x => x.FailPatternData).Select(x => x.EqName);
                IEnumerable<int> binTotal = _allDiceInfos.SelectMany(x => x.InstanceList).Where(x => x.InstanceName.EqualsIgnoreCase(InstNameList[idx]))
                    .SelectMany(x => x.FailPatternData).Select(x => x.Bin);
                foreach (string testName in testNames)
                {
                    for (int i = binTotal.Min(); i <= binTotal.Max(); i++)
                    {
                        for (int j = eqnTotal.Max(); j >= eqnTotal.Min(); j--)
                        {
                            IEnumerable<PatternInfo> failPat = _allDiceInfos.SelectMany(x => x.InstanceList).Where(x => x.InstanceName.EqualsIgnoreCase(InstNameList[idx])).SelectMany(x => x.FailPatternData).Where(x => x.PatternName == testName && x.Bin == i && x.EqName == j);
                            int diceCount = _allDiceInfos.SelectMany(x => x.InstanceList).Count(x => x.InstanceName.EqualsIgnoreCase(InstNameList[idx]));
                            if (failPat.Count() == diceCount)
                            {
                                var array = new List<object> { InstNameList[idx], testName, i, j, diceCount };
                                xlWorkSheet.Cells[currRow++, 1].PrintExcelRow(array.ToArray());
                            }
                        }
                    }
                }
            }
            xlWorkSheet.MergeColumn(1);
            xlWorkSheet.Cells.TryAutoFitColumns();
            xlWorkSheet.Column(1).Width = 20;
        }

        public void WritePatternFailList()
        {
            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("PatternFailList");
            int currRow = 1;
            var titles = new List<string> { "TestIns Name", "Test Name", "Site", "Bin", "EQN" };
            currRow = xlWorkSheet.Cells[currRow, 1].PrintExcelRow(titles.ToArray());
            for (int idx = 0; idx < InstNameList.Count; idx++)
            {
                IEnumerable<string> testNames = _allDiceInfos.SelectMany(x => x.InstanceList).Where(x => x.InstanceName.EqualsIgnoreCase(InstNameList[idx]))
                    .SelectMany(x => x.FailPatternData).Select(x => x.PatternName).Distinct();
                foreach (string testName in testNames)
                {
                    foreach (Base.SiteInfo oneDice in _allDiceInfos)
                    {
                        int diceIndex = _allDiceInfos.IndexOf(oneDice);
                        InstanceData? instanceData = oneDice.InstanceList.Find(x => x.InstanceName == InstNameList[idx]);
                        if (instanceData == null || instanceData.FailPatternData == null)
                        {
                            continue;
                        }

                        List<PatternInfo> patternData = instanceData.FailPatternData.FindAll(x => x.PatternName == testName);
                        if (patternData != null)
                        {
                            foreach (PatternInfo oneData in patternData)
                            {
                                var array = new List<object> { InstNameList[idx], testName, diceIndex, oneData.Bin, oneData.EqName };
                                xlWorkSheet.Cells[currRow++, 1].PrintExcelRow(array.ToArray());
                            }
                        }
                    }
                }
            }
            xlWorkSheet.MergeColumn(1);
            xlWorkSheet.Cells.TryAutoFitColumns();
            xlWorkSheet.Column(1).Width = 20;
        }

        public void WriteCofSummary()
        {
            ExcelWorksheet xlWorkSheetChart = XlWorkBook.AddSheet("COFSummaryChart");
            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("COFSummary");
            //print title
            int currRow = 1;
            int chartCurrRow = 1;
            var titles = new List<string> { "TestIns Name", "Test Name" };

            foreach (Base.SiteInfo oneDice in _allDiceInfos)
            {
                titles.Add($"X:{oneDice.XCoor} Y:{oneDice.YCoor}");
            }

            currRow = xlWorkSheet.Cells[currRow, 1].PrintExcelRow(titles.ToArray());
            var patSteps = new List<PatSteps>();
            int patIdx = -1;

            PrintData(xlWorkSheetChart, xlWorkSheet, ref currRow, ref chartCurrRow, patSteps, ref patIdx);

            xlWorkSheet.MergeColumn(1);
            xlWorkSheet.Cells.TryAutoFitColumns();
            xlWorkSheet.Column(1).Width = 20;
        }

        private void PrintData(ExcelWorksheet xlWorkSheetChart, ExcelWorksheet xlWorkSheet, ref int currRow, ref int chartCurrRow, List<PatSteps> patStepss, ref int patIdx)
        {
            for (int idx = 0; idx < InstNameList.Count; idx++)
            {
                int pwrIdx = -1;
                IEnumerable<string> testNames = _allDiceInfos.SelectMany(x => x.InstanceList).Where(x => x.InstanceName.EqualsIgnoreCase(InstNameList[idx]))
                    .SelectMany(x => x.PatternResultRows).Select(x => x.TestName).Distinct();
                HandleByTestNames(xlWorkSheet, ref currRow, patStepss, idx, ref pwrIdx, testNames);

                if (testNames.Any() && patIdx != idx)
                {
                    patIdx = ProcessByTestNames(xlWorkSheetChart, ref chartCurrRow, patStepss, idx, pwrIdx);
                }
            }
        }

        private int ProcessByTestNames(ExcelWorksheet excelWorksheet, ref int chartCurrRow, List<PatSteps> patStepss, int idx, int pwrIdx)
        {
            int patIdx;
            var chartTitles = new List<string> { "TestIns Name", InstNameList[idx] };
            chartCurrRow = excelWorksheet.Cells[chartCurrRow, 1].PrintExcelRow(chartTitles.ToArray());
            var arrayEq = new List<object> { "TestName EQ" };
            List<PowerStep> pwrList = _allDiceInfos.Find(x => x.AllPowers[pwrIdx] != null)!.AllPowers[pwrIdx].AllSteps;
            int itemGap = 10;
            //CP counter by each pattern
            foreach (PatSteps onePat in patStepss)
            {
                foreach (SiteInfo siteInfo in onePat.SiteInfos)
                {
                    if (onePat.CpCount.Find(x => x.CpValue.Equals(siteInfo.Cp)) != null)
                    {
                        onePat.CpCount.Find(x => x.CpValue.Equals(siteInfo.Cp))!.Count++;
                    }
                    else
                    {
                        onePat.CpCount.Add(new CpInfo() { CpValue = siteInfo.Cp, Count = 1 });
                    }
                }
            }

            //calculate X coor range for EQ line chart
            for (int i = 0; i < pwrList.Count; i++)
            {
                arrayEq.Add($"Bin{pwrList[i].Bin} EQ{pwrList[i].EqName}");
            }
            arrayEq.Add("Fail");
            double maxCp;
            double minCp;
            //calculate X coor range for CP bar chart
            if (patStepss.SelectMany(x => x.CpCount).Sum(y => y.CpValue) > 0)
            {
                maxCp = patStepss.SelectMany(x => x.CpCount).Max(y => y.CpValue);
                minCp = patStepss.SelectMany(x => x.CpCount).Where(y => y.CpValue > 0).Min(y => y.CpValue);
            }
            else
            {
                maxCp = 0;
                minCp = 0;
            }
            var arrayCp = new List<object> { "TestName CP" };

            //Using Max/Min CP information to gen  X items for CP bar chart
            for (double i = minCp; i <= maxCp; i += itemGap)
            {
                string itemName = i + "~" + (i + itemGap - 1);
                arrayCp.Add(itemName);

                foreach (PatSteps onePat in patStepss)
                {
                    int count = onePat.CpCount.FindAll(x => x.CpValue >= i && x.CpValue < i + itemGap).Sum(y => y.Count);
                    onePat.PatCpCount.Add(new PatCp() { ItemName = itemName, Count = count });
                }
            }
            int xeqCount = arrayEq.Count;
            int xcpStart = xeqCount + 2;
            int xcpCount = arrayCp.Count;
            excelWorksheet.Cells[chartCurrRow, 1].PrintExcelRow(arrayEq.ToArray());
            excelWorksheet.Cells[chartCurrRow++, xcpStart].PrintExcelRow(arrayCp.ToArray());

            arrayEq.Clear();
            arrayCp.Clear();
            //calculate steps count for each pattern
            foreach (PatSteps onePat in patStepss)
            {
                onePat.InitStepCount();
                foreach (SiteInfo siteInfo in onePat.SiteInfos)
                {
                    onePat.StepCount[siteInfo.Step]++;
                }
            }

            foreach (PatSteps onePat in patStepss)
            {
                arrayEq.Add($"{onePat.PatName} ");
                arrayCp.Add($"{onePat.PatName} ");
                for (int i = 0; i < onePat.AllPwrCount; i++)
                {
                    arrayEq.Add(onePat.StepCount[i]);
                }
                for (int j = 0; j < onePat.PatCpCount.Count; j++)
                {
                    arrayCp.Add(onePat.PatCpCount[j].Count);
                }
                excelWorksheet.Cells[chartCurrRow, 1].PrintExcelRow(arrayEq.ToArray());
                excelWorksheet.Cells[chartCurrRow++, xcpStart].PrintExcelRow(arrayCp.ToArray());
                arrayEq.Clear();
                arrayCp.Clear();
            }
            //Drawing EQ line chart and CP bar chart
            chartCurrRow = DrawChart(excelWorksheet, InstNameList[idx], chartCurrRow, xeqCount, xcpCount, patStepss);
            chartCurrRow++;
            patStepss.Clear();
            patIdx = idx;
            return patIdx;
        }

        private void HandleByTestNames(ExcelWorksheet excelWorksheet, ref int currRow, List<PatSteps> patStepss, int idx, ref int pwrIdx, IEnumerable<string> testNames)
        {
            foreach (string testName in testNames)
            {
                var array = new List<object> { InstNameList[idx], testName };
                string testPat = testName[..(testName.Length - testName.Split('_').Last().Length - 1)];

                foreach (Base.SiteInfo oneDice in _allDiceInfos)
                {
                    int diceIndex = _allDiceInfos.IndexOf(oneDice);
                    InstanceData? instanceData = oneDice.InstanceList.Find(x => x.InstanceName == InstNameList[idx]);
                    if (instanceData == null || instanceData.PatternResultRows == null)
                    {
                        array.Add("N/A");
                        continue;
                    }
                    LimitRow? row = instanceData.PatternResultRows.Find(x => x.TestName.EqualsIgnoreCase(testName));
                    if (row == null)
                    {
                        array.Add("N/A");
                        continue;
                    }

                    array.Add(row.Measured);

                    pwrIdx = oneDice.InstanceList.Find(x => x.InstanceName == InstNameList[idx])!.PowersIdx;
                    if (pwrIdx == -1)
                    {
                        continue;
                    }

                    int pwrCount = oneDice.AllPowers[pwrIdx].AllSteps.Count;
                    int binXincrease = oneDice.AllPowers[pwrIdx].AllSteps.FindIndex(x => x.Bin.Equals(2));
                    int binYincrease = oneDice.AllPowers[pwrIdx].AllSteps.FindIndex(x => x.Bin.Equals(3));

                    if (patStepss.Find(x => x.PatName == testPat) != null)
                    {
                        if (patStepss.Find(x => x.PatName == testPat)!.SiteInfos.Find(y => y.Idx.Equals(diceIndex)) == null)
                        {
                            var siteInfos = new SiteInfo { Idx = diceIndex };
                            patStepss.Find(x => x.PatName == testPat)!.SiteInfos.Add(siteInfos);
                        }
                    }
                    else
                    {
                        var siteInfos = new SiteInfo { Idx = diceIndex };
                        patStepss.Add(new PatSteps { PatName = testPat, AllPwrCount = pwrCount + 1, PwrIdx = pwrIdx });
                        patStepss.Find(x => x.PatName == testPat)!.SiteInfos.Add(siteInfos);
                    }
                    SiteInfo testPatInfo = patStepss.Find(x => x.PatName == testPat)!.SiteInfos.Find(y => y.Idx.Equals(diceIndex))!;
                    string testItem = row.TestName.Split('_').Last();
                    switch (testItem)
                    {
                        case "CP":
                            {
                                testPatInfo.Cp = row.Measured;
                                break;
                            }
                        case "EQN":
                            {
                                if (testPatInfo.Step < binXincrease && testPatInfo.Step == 0) // for bin 1
                                {
                                    testPatInfo.Step += binXincrease - (int)row.Measured;
                                }
                                else if (row.Measured != 0)
                                {
                                    testPatInfo.Step += (int)row.Measured - 1;
                                }
                                break;
                            }
                        case "PASSBIN":
                            {
                                switch (row.Measured.ToString())
                                {
                                    case "0":
                                        {
                                            testPatInfo.Step += pwrCount;
                                            break;
                                        }
                                    case "2":
                                        {
                                            testPatInfo.Step += binXincrease;
                                            break;
                                        }
                                    case "3":
                                        {
                                            testPatInfo.Step += binYincrease;
                                            break;
                                        }
                                }
                                break;
                            }
                    }
                }
                excelWorksheet.Cells[currRow++, 1].PrintExcelRow(array.ToArray());
            }
        }

        private static int DrawChart(ExcelWorksheet excelWorksheet, string nameOfChart, int startRow, int eqCount, int cpCount, List<PatSteps> patStepss)
        {
            int chartWidth = 10;
            int chartHeigh = 15;
            //EQ line chart part
            var lineChart = excelWorksheet.Drawings.AddChart(nameOfChart + "Line", eChartType.LineMarkers) as ExcelLineChart;
            lineChart!.From.Row = startRow;
            int startCol = 1;
            int patCount = patStepss.Count;
            int seriesXRow = startRow - 1 - patCount;
            ExcelRange eqseriesX = excelWorksheet.Cells[seriesXRow, 2, seriesXRow, eqCount];
            lineChart.From.Column = startCol;
            lineChart.To.Column = startCol + chartWidth;
            lineChart.To.Row = startRow + chartHeigh;
            for (int i = patCount; i > 0; i--)
            {
                int seriesYRow = startRow - i;
                ExcelRange seriesY = excelWorksheet.Cells[seriesYRow, 2, seriesYRow, eqCount];
                lineChart.Series.Add(seriesY, eqseriesX).Header = excelWorksheet.Cells[seriesYRow, 1].Value.ToString();
            }
            lineChart.Legend.Font.SetFromFont(new Font("Arial", 6));
            lineChart.XAxis.MajorTickMark = eAxisTickMark.None;
            lineChart.XAxis.MinorTickMark = eAxisTickMark.Cross;
            lineChart.YAxis.MajorTickMark = eAxisTickMark.Out;
            lineChart.YAxis.MinorTickMark = eAxisTickMark.None;
            lineChart.Title.Font.Color = Color.Black;
            lineChart.Title.Font.SetFromFont(new Font("Times New Roman", 8));
            lineChart.Title.Text = nameOfChart;
            /*********************************************************************************************/
            //CP bar chart part
            ExcelChart cpBarChart = excelWorksheet.Drawings.AddChart(nameOfChart + "Bar", eChartType.ColumnClustered);
            int cpStartCol = eqCount + 2;
            cpBarChart.From.Column = cpStartCol;
            cpBarChart.To.Column = cpStartCol + chartWidth;
            cpBarChart.From.Row = startRow;
            cpBarChart.To.Row = startRow + chartHeigh;
            cpBarChart.XAxis.MajorTickMark = eAxisTickMark.In;
            cpBarChart.XAxis.MinorTickMark = eAxisTickMark.None;
            cpBarChart.YAxis.MajorTickMark = eAxisTickMark.Out;
            cpBarChart.YAxis.MinorTickMark = eAxisTickMark.None;
            cpBarChart.Title.Font.Color = Color.Black;
            cpBarChart.Title.Text = nameOfChart;
            cpBarChart.Legend.Font.SetFromFont(new Font("Arial", 6));
            cpBarChart.Title.Font.SetFromFont(new Font("Times New Roman", 8));
            ExcelRange seriesXCp = excelWorksheet.Cells[seriesXRow, cpStartCol + 1, seriesXRow, cpStartCol + cpCount - 1];
            for (int i = patCount; i > 0; i--)
            {
                int seriesYRow = startRow - i;
                ExcelRange seriesY = excelWorksheet.Cells[seriesYRow, cpStartCol + 1, seriesYRow, cpStartCol + cpCount - 1];
                cpBarChart.Series.Add(seriesY, seriesXCp).Header = excelWorksheet.Cells[seriesYRow, cpStartCol].Value.ToString();
            }
            return startRow + chartHeigh;
        }

        public void WritePerformanceList()
        {
            GeneratePerformanceSheet("PerformanceList", isDetailedMode: false);
        }

        public void WritePerformanceList1()
        {
            GeneratePerformanceSheet("PerformanceList_1", isDetailedMode: true);
        }

        private void GeneratePerformanceSheet(string sheetName, bool isDetailedMode)
        {
            if (InstNameList.Count == 0)
            {
                return;
            }

            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet(sheetName);
            int currRow = 1;

            // 1. Generate Headers
            List<string> titles = isDetailedMode
                ? ["Type", "Perf Mode", "TestIns Name", "Parameter", "Total", "Avg"]
                : ["Type", "Perf Mode", "TestIns Name", "Parameter"];

            foreach (Base.SiteInfo oneDice in _allDiceInfos)
            {
                titles.Add($"X:{oneDice.XCoor} Y:{oneDice.YCoor}");
            }
            currRow = xlWorkSheet.Cells[currRow, 1].PrintExcelRow(titles.ToArray());

            // 2. Cache instance references to fix SelectMany nested execution performance issues
            var allInstances = _allDiceInfos.SelectMany(x => x.InstanceList).ToList();

            // 3. Process Rows
            currRow = PrintData(isDetailedMode, xlWorkSheet, currRow, allInstances);

            // Format Sheet Formatting Constraints
            xlWorkSheet.MergeColumn(3);
            xlWorkSheet.Cells.TryAutoFitColumns();
            xlWorkSheet.Column(3).Width = 20;
        }

        private int PrintData(bool isDetailedMode, ExcelWorksheet excelWorksheet, int currRow, List<InstanceData> instanceDatas)
        {
            for (int idx = 0; idx < InstNameList.Count; idx++)
            {
                string currentInst = InstNameList[idx];

                // Unique skip rules preserved per variant
                if (!isDetailedMode && currentInst.EndsWithIgnoreCase("HBV"))
                {
                    break;
                }

                if (isDetailedMode && currentInst.Contains("HBV>"))
                {
                    break;
                }

                // Extract Type Segment
                string[] instSpt = currentInst.Split('_');
                string performanceType = GetPerformanceType(instSpt);

                // Extract Mode Segment
                string powerName = "";
                InstanceData? matchedInstance = instanceDatas.FirstOrDefault(y => y.InstanceName == currentInst);
                if (matchedInstance != null)
                {
                    powerName = _powerNames[matchedInstance.PowersIdx];
                }

                // Initialize Data Rows
                var arrayBin = new List<object> { performanceType, powerName, currentInst, "Bin" };
                var arrayEqn = new List<object> { performanceType, powerName, currentInst, "Equation" };
                var arrayLvcc = new List<object> { performanceType, powerName, currentInst, "LVCC" };
                var arrayIds = new List<object> { performanceType, powerName, currentInst, "IDS" };
                var arraySearchStep = new List<object> { performanceType, powerName, currentInst, "STEP" };
                var arrayActualStep = new List<object> { performanceType, powerName, currentInst, "StepDiff" };

                var binNoList = new List<double>();
                var eqNoList = new List<double>();
                var lvccList = new List<double>();
                var idsList = new List<double>();
                var stepList = new List<double>();
                var actualstepList = new List<double>();

                // Collect Core Data Metrics
                foreach (Base.SiteInfo oneDice in _allDiceInfos)
                {
                    InstanceData? instanceData = oneDice.InstanceList.Find(x => x.InstanceName == currentInst);
                    double binNo = instanceData?.Bin ?? -1;
                    double eqNo = instanceData?.Eqns ?? -1;
                    double lvcc = instanceData?.Lvcc ?? 0;
                    double ids = instanceData?.Ids ?? 0;
                    double step = instanceData?.UsedSteps ?? 0;
                    double actualstep = instanceData?.ActualSteps ?? 0;

                    if (isDetailedMode)
                    {
                        binNoList.Add(binNo);
                        eqNoList.Add(eqNo);
                        lvccList.Add(lvcc);
                        idsList.Add(ids);
                        stepList.Add(step);
                        actualstepList.Add(actualstep);
                    }
                    else
                    {
                        arrayBin.Add((int)binNo);
                        arrayEqn.Add((int)eqNo);
                        arrayLvcc.Add(Math.Round(lvcc, 4, MidpointRounding.AwayFromZero));
                        arrayIds.Add(Math.Round(ids, 4, MidpointRounding.AwayFromZero));
                        arraySearchStep.Add(Math.Round(step, 4, MidpointRounding.AwayFromZero));
                        arrayActualStep.Add(Math.Round(actualstep, 4, MidpointRounding.AwayFromZero));
                    }
                }

                // Process Averages and Totals via SetArrayValue for detailed mode
                if (isDetailedMode)
                {
                    SetArrayValue(arrayBin, binNoList);
                    SetArrayValue(arrayEqn, eqNoList);
                    SetArrayValue(arrayLvcc, [.. lvccList.Select(x => Math.Round(x, 4, MidpointRounding.AwayFromZero))]);
                    SetArrayValue(arrayIds, [.. idsList.Select(x => Math.Round(x, 4, MidpointRounding.AwayFromZero))]);
                    SetArrayValue(arraySearchStep, [.. stepList.Select(x => Math.Round(x, 4, MidpointRounding.AwayFromZero))]);
                    SetArrayValue(arrayActualStep, [.. actualstepList.Select(x => Math.Round(x, 4, MidpointRounding.AwayFromZero))]);
                }

                // Print Accumulated Outputs sequentially
                excelWorksheet.Cells[currRow++, 1].PrintExcelRow(arrayBin.ToArray());
                excelWorksheet.Cells[currRow++, 1].PrintExcelRow(arrayEqn.ToArray());
                excelWorksheet.Cells[currRow++, 1].PrintExcelRow(arrayLvcc.ToArray());
                excelWorksheet.Cells[currRow++, 1].PrintExcelRow(arrayIds.ToArray());
                excelWorksheet.Cells[currRow++, 1].PrintExcelRow(arraySearchStep.ToArray());
                excelWorksheet.Cells[currRow++, 1].PrintExcelRow(arrayActualStep.ToArray());
            }

            return currRow;
        }

        private static string GetPerformanceType(string[] instSpt)
        {
            string performanceType = instSpt[0];
            for (int instSptIdx = 1; instSptIdx < instSpt.Length; instSptIdx++)
            {
                string strTmp = instSpt[instSptIdx].ToUpper();
                if (strTmp.Contains("CPU") || strTmp.Contains("GPU") || strTmp.Contains("SOC"))
                {
                    performanceType += "_" + instSpt[instSptIdx];
                }
            }

            return performanceType;
        }

        public void WritePerTouchdownStepList()
        {
            if (InstNameList.Count == 0)
            {
                return;
            }

            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("PerTouchdownStepList");

            //pring title
            PrintHeader(xlWorkSheet, out int currRow, out int tchnumtmp);

            var arrayMaxExtraStep = new List<double>();
            var arrayActualExtraStep = new Dictionary<double, string>();

            //print data
            PrintData(xlWorkSheet, ref currRow, ref tchnumtmp, arrayMaxExtraStep, arrayActualExtraStep);

            tchnumtmp = -1;
            var arrayTotalMaxExtraStep = new List<object> { "", "", "", "TotalMaxExtraStep" };
            foreach (Base.SiteInfo oneDice in _allDiceInfos)
            {
                if (oneDice.TuchNum != tchnumtmp && tchnumtmp != -1)
                {
                    arrayTotalMaxExtraStep.Add("");
                }
                if (oneDice.TuchNum != tchnumtmp)
                {
                    arrayTotalMaxExtraStep.Add(Math.Round(arrayMaxExtraStep[oneDice.TuchNum], 4, MidpointRounding.AwayFromZero));
                }
                else
                {
                    arrayTotalMaxExtraStep.Add("");
                }

                tchnumtmp = oneDice.TuchNum;
            }

            xlWorkSheet.MergeColumn(3);
            xlWorkSheet.Cells.TryAutoFitColumns();
            xlWorkSheet.Column(3).Width = 20;

            xlWorkSheet.Cells[currRow++, 1].PrintExcelRow(arrayTotalMaxExtraStep.ToArray());
        }

        private void PrintData(ExcelWorksheet excelWorksheet, ref int currRow, ref int tchnumtmp, List<double> arrayMaxExtraStep, Dictionary<double, string> arrayActualExtraStep)
        {
            for (int idx = 0; idx < InstNameList.Count; idx++)
            {
                if (InstNameList[idx].EndsWithIgnoreCase("HBV")) //HV不印
                {
                    continue;
                }

                List<string> instSpt = [.. InstNameList[idx].Split('_')];
                string performanceType = instSpt[0];
                for (int instSptIdx = 1; instSptIdx < instSpt.Count; instSptIdx++)
                {
                    string strTmp = instSpt[instSptIdx].ToUpper();
                    if (strTmp.Contains("CPU") || strTmp.Contains("GPU") || strTmp.Contains("SOC"))
                    {
                        performanceType += "_" + instSpt[instSptIdx];
                    }
                }

                string powerName = "";
                if (_allDiceInfos.SelectMany(x => x.InstanceList).ToList().Exists(y => y.InstanceName == InstNameList[idx]))
                {
                    int itemPwrIdx = _allDiceInfos.SelectMany(x => x.InstanceList).First(y => y.InstanceName == InstNameList[idx]).PowersIdx;
                    powerName = _powerNames[itemPwrIdx];
                }

                var arraySearchStep = new List<object> { performanceType, powerName, InstNameList[idx], "STEP" };
                var arrayActualStep = new List<object> { performanceType, powerName, InstNameList[idx], "StepDiff" };
                var arrayExtraStep = new List<object> { performanceType, powerName, InstNameList[idx], "ExtraStep" };

                arrayActualExtraStep.Clear();
                tchnumtmp = -1;
                double maxextrastep = 0;
                double maxactualstep = 0;
                foreach (Base.SiteInfo oneDice in _allDiceInfos)
                {
                    InstanceData? instanceData = oneDice.InstanceList.Find(x => x.InstanceName == InstNameList[idx]);

                    double step = instanceData?.UsedSteps ?? 0;
                    double actualstep = instanceData?.ActualSteps ?? 0;
                    double extrastep = step - actualstep;

                    if (tchnumtmp == -1)
                    {
                        maxextrastep = extrastep;
                        maxactualstep = actualstep;
                    }
                    if (oneDice.TuchNum != tchnumtmp && tchnumtmp != -1)
                    {
                        arraySearchStep.Add("");
                        arrayExtraStep.Add("");
                        arrayActualStep.Add(arrayActualExtraStep[maxactualstep]);

                        arrayActualExtraStep.Clear();

                        if (arrayMaxExtraStep.Count - 1 < tchnumtmp)
                        {
                            arrayMaxExtraStep.Add(maxextrastep);
                        }
                        else
                        {
                            arrayMaxExtraStep[tchnumtmp] += maxextrastep;
                        }

                        maxextrastep = extrastep;
                        maxactualstep = actualstep;
                    }
                    else if (oneDice.TuchNum == tchnumtmp && tchnumtmp != -1)
                    {
                        if (maxextrastep < extrastep)
                        {
                            maxextrastep = extrastep;
                        }

                        if (maxactualstep < actualstep)
                        {
                            maxactualstep = actualstep;
                        }
                    }

                    if (arrayActualExtraStep.TryGetValue(actualstep, out string? value))
                    {
                        arrayActualExtraStep[actualstep] = value + ", " +
                                                           $"X:{oneDice.XCoor} Y:{oneDice.YCoor}";
                    }
                    else
                    {
                        arrayActualExtraStep.Add(actualstep, $"X:{oneDice.XCoor} Y:{oneDice.YCoor}");
                    }

                    arraySearchStep.Add(Math.Round(step, 4, MidpointRounding.AwayFromZero));
                    arrayActualStep.Add(Math.Round(actualstep, 4, MidpointRounding.AwayFromZero));
                    arrayExtraStep.Add(Math.Round(extrastep, 4, MidpointRounding.AwayFromZero));

                    tchnumtmp = oneDice.TuchNum;
                }

                if (tchnumtmp >= 0)
                {
                    arraySearchStep.Add("");
                    arrayExtraStep.Add("");
                    arrayActualStep.Add(arrayActualExtraStep[maxactualstep]);

                    if (arrayMaxExtraStep.Count - 1 < tchnumtmp)
                    {
                        arrayMaxExtraStep.Add(maxextrastep);
                    }
                    else
                    {
                        arrayMaxExtraStep[tchnumtmp] += maxextrastep;
                    }
                }

                excelWorksheet.Cells[currRow++, 1].PrintExcelRow(arraySearchStep.ToArray());
                excelWorksheet.Cells[currRow++, 1].PrintExcelRow(arrayActualStep.ToArray());
                excelWorksheet.Cells[currRow++, 1].PrintExcelRow(arrayExtraStep.ToArray());
            }
        }

        private void PrintHeader(ExcelWorksheet excelWorksheet, out int currRow, out int tchnumtmp)
        {
            currRow = 1;
            var titles = new List<string> { "Type", "Perf Mode", "TestIns Name", "Parameter" };

            tchnumtmp = -1;
            foreach (Base.SiteInfo oneDice in _allDiceInfos)
            {

                if (oneDice.TuchNum != tchnumtmp && tchnumtmp != -1)
                {
                    titles.Add("");
                }
                titles.Add($"X:{oneDice.XCoor} Y:{oneDice.YCoor}");

                tchnumtmp = oneDice.TuchNum;
            }

            currRow = excelWorksheet.Cells[currRow, 1].PrintExcelRow(titles.ToArray());
        }

        private static void SetArrayValue(List<object> arrayBin, List<double> binNoList)
        {
            arrayBin.Add(binNoList.Count);
            double average = binNoList.Count == 0 ? 0 : binNoList.Average();
            arrayBin.Add(average);
            foreach (double eqNo in binNoList)
            {
                arrayBin.Add(eqNo);
            }
        }

        public void WritePowerBinningList()
        {
            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("PowerBinning");

            //ptring title
            int currRow = 1;
            var titles = new List<string> { "Type", "Parameter" };
            foreach (Base.SiteInfo oneDice in _allDiceInfos)
            {
                titles.Add($"X:{oneDice.XCoor} Y:{oneDice.YCoor}");
            }

            currRow = xlWorkSheet.Cells[currRow, 1].PrintExcelRow(titles.ToArray());

            //print data
            currRow = PrintData(xlWorkSheet, currRow);

            currRow += 3;
            xlWorkSheet.Cells[currRow, 1].Value = "PowerBinning Check";
            string[] list = [.. _allDiceInfos.Select(x => x.CheckResult.IsPowerBinningPass ? "P" : "F")];
            currRow = xlWorkSheet.Cells[currRow, 3].PrintExcelRow(list);

            xlWorkSheet.Cells[currRow, 1].Value = "IsPowerBinningFail";
            string[] powerBinningFail = [.. _allDiceInfos.Select(x => x.PowerBinningFail)];
            currRow = xlWorkSheet.Cells[currRow, 3].PrintExcelRow(powerBinningFail);

            xlWorkSheet.Cells[currRow, 1].Value = "IsPowerBinningBinXFail";
            string[] powerBinningBinXFail = [.. _allDiceInfos.Select(x => x.PowerBinningBinXFail)];
            currRow = xlWorkSheet.Cells[currRow, 3].PrintExcelRow(powerBinningBinXFail);

            IExcelConditionalFormattingEqual cond1 = xlWorkSheet.ConditionalFormatting.AddEqual(new ExcelAddress("1:" + currRow));
            cond1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cond1.Style.Fill.BackgroundColor.Color = Color.Red;
            cond1.Style.Font.Color.Color = Color.White;
            cond1.Formula = "\"F\"";

            xlWorkSheet.Cells.TryAutoFitColumns();
        }

        private int PrintData(ExcelWorksheet excelWorksheet, int currRow)
        {
            if (_allDiceInfos.SelectMany(x => x.PowerBinningPTotalSummary).Any())
            {
                int max = _allDiceInfos.Max(x => x.PowerBinningPTotalSummary.Count);
                var sheets = _allDiceInfos.Where(x => x.PowerBinningPTotalSummary.Count == max).ToList();
                for (int i = 0; i < max; i++)
                {
                    string item = sheets[0].PowerBinningPTotalSummary[i].Item1;
                    var arrayPTotal = new List<string> { "P_Total", item };
                    if (!item.Contains("P_TOTAL", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var data = new List<string>();
                    foreach (Base.SiteInfo oneDice in _allDiceInfos)
                    {
                        Tuple<string, string>? powerbonning = oneDice.PowerBinningPTotalSummary.Find(x => x.Item1.EqualsIgnoreCase(item));
                        if (powerbonning != null)
                        {
                            data.Add(powerbonning.Item2);
                        }
                        else
                        {
                            data.Add("N/A");
                        }
                    }
                    arrayPTotal.AddRange(data);
                    currRow = excelWorksheet.Cells[currRow, 1].PrintExcelRow(arrayPTotal.ToArray());
                }
            }

            return currRow;
        }

        private static string GetMergePat(List<string> patterns, string instanceName)
        {
            string mergePat = string.Join("-", patterns);
            if (string.IsNullOrEmpty(mergePat))
            {
                mergePat = instanceName;
            }

            if (mergePat.Contains('+'))
            {
                mergePat = mergePat.Replace('+', '-');
            }

            return mergePat;
        }

        private static List<Tuple<string, int>>? GetGenericPatternList(List<List<PatternRow>> patternRows)
        {
            return SelectPatternList(patternRows, row => row.GenericPatternName);
        }

        private static List<Tuple<string, int>>? GetPatternVerList(List<List<PatternRow>> patternRows)
        {
            return SelectPatternList(patternRows, row => row.PatternName);
        }

        private static List<Tuple<string, int>>? SelectPatternList(List<List<PatternRow>> patternRows, Func<PatternRow, string> nameSelector)
        {
            if (patternRows == null || patternRows.Count == 0)
            {
                return null;
            }

            // 1. Search for the first completely passing block sequence
            foreach (List<PatternRow> datalogPatternRow in patternRows)
            {
                if (datalogPatternRow != null && datalogPatternRow.All(x => !x.IsFail))
                {
                    return [.. datalogPatternRow.Select(x => Tuple.Create(nameSelector(x), x.PatternLine?.LineNo ?? 0))];
                }
            }

            // 2. Fallback to extracting the trailing data segment
            List<PatternRow> lastRow = patternRows.Last();
            return lastRow?
                .Select(x => Tuple.Create(nameSelector(x), x.PatternLine?.LineNo ?? 0))
                .ToList();
        }

        public static string GetMode(string instanceName, List<string> modes)
        {
            foreach (string mode in modes)
            {
                if (instanceName.StartsWithIgnoreCase(mode))
                {
                    return mode;
                }
            }
            return "";
        }

        public void OutputAlarmList(List<Alarm> alarms)
        {
            if (alarms.Count == 0)
            {
                return;
            }

            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("AlarmList");

            //ptring title
            int currRow = 1;
            string title = "Line_No,Alarm_Message,IsBeforeBV";
            currRow = xlWorkSheet.Cells[currRow, 1].PrintExcelRow(title.Split(','));

            //print data
            if (alarms.Count > 0)
            {
                var data = new List<List<string>>
                {
                    alarms.Select(x => x._alarmMessage.LineNo.ToString(CultureInfo.InvariantCulture)).ToList(),
                    alarms.Select(x =>  x._alarmMessage.Line).ToList(),
                    alarms.Select(x => x.IsBeforeBv.ToString()).ToList(),
                };
                xlWorkSheet.Cells[currRow, 1].PrintExcelColByList(data);
                xlWorkSheet.Cells.TryAutoFitColumns();
            }
        }

        public bool IsContainSheet(string sheetName)
        {
            bool flag = false;
            foreach (ExcelWorksheet item in XlWorkBook.Worksheets)
            {
                if (item.Name == sheetName)
                {
                    flag = true;
                }
            }

            return flag;
        }

        public void ExcelClose()
        {
            if (XlWorkBook.Worksheets.Count > 0)
            {
                XlWorkBook.Worksheets[0].Select();
                XlPackage.SaveAs(new FileInfo(DestExcel));
                XlPackage.Dispose();
            }
        }

        private static void ReplaceFlagNameByStatus(ref string expression, Dictionary<string, string> harvesFlags)
        {
            string[] flags = expression.Split(["&", "|", "(", ")", " ", "!"], StringSplitOptions.RemoveEmptyEntries);
            foreach (string flag in flags)
            {
                if (harvesFlags.TryGetValue(flag, out string? harvesFlag))
                {
                    expression = expression.Replace(flag, harvesFlag);
                }
            }
        }

        private static bool ParserExpression(string input)
        {
            input = input.Replace("&&", "&").Replace("||", "|");
            IEnumerable<string> list = Reg.RegexSplit.Split(input).Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim());
            bool result = false;

            var textStack = new Stack<bool>();
            var opStack = new Stack<string>();
            bool fetchNUmber = true;

            foreach (string m in list)
            {
                if (m == "(")
                {
                    opStack.Push(m);
                }
                else if (m == ")")
                {
                    Walk(textStack, opStack, true);
                }
                else
                {
                    if (fetchNUmber)
                    {
                        if (m.StartsWith('!'))
                        {
                            if (m.TrimStart('!') == "T")
                            {
                                textStack.Push(false);
                            }
                            else
                            {
                                textStack.Push(true);
                            }
                        }
                        else
                        {
                            if (m == "T")
                            {
                                textStack.Push(true);
                            }
                            else
                            {
                                textStack.Push(false);
                            }
                        }

                    }
                    else
                    {
                        if (m == "|")
                        {
                            Walk(textStack, opStack);
                            opStack.Push(m);
                        }
                        if (m == "&")
                        {
                            opStack.Push(m);
                        }
                    }
                    fetchNUmber = !fetchNUmber;
                }
            }
            Walk(textStack, opStack);
            result = textStack.Pop();
            return result;
        }

        private static void Walk(Stack<bool> textStack, Stack<string> opStack, bool close = false)
        {
            while (opStack.Count > 0 && textStack.Count > 1)
            {

                if (opStack.Peek() == "(")
                {
                    if (close)
                    {
                        opStack.Pop();
                    }

                    break;
                }
                string operation = opStack.Pop();
                bool d2 = textStack.Pop();
                bool d1 = textStack.Pop();
                bool z1 = false;
                switch (operation)
                {
                    case "&":
                        z1 = d2 & d1;
                        break;
                    case "|":
                        z1 = d2 | d1;
                        break;
                }
                textStack.Push(z1);
            }
        }
    }
}
