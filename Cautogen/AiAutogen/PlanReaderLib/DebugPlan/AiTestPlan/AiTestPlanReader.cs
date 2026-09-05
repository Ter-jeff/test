using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace DebugPlanReaderLib.DebugPlan
{
    public class AiTestPlanReader : MySheetReader<AiTestPlanSheet>
    {
        private const string HeaderUseNotUse = "Use/Not Use";
        private const string HeaderComment = "Comment";
        private const string HeaderTestInstanceName = "Test instance name";
        private const string HeaderAiType = "AI type";
        private const string HeaderDataLoggingSetting = "Data logging setting";
        private const string HeaderTimeset = "Timeset";
        private const string HeaderVoltageCategory = "Voltage Category";
        private const string HeaderOrder = "Order";
        private const string HeaderSearch = "Search";
        private const string HeaderTempCondition = "Temp. Condition";
        private const string HeaderPattern = @"Pattern\s*(?<index>\d+)";
        private const string HeaderStart = "Start";
        private const string HeaderStop = "Stop";
        private const string HeaderStep = "Step";
        private const string HeaderSelsramDssc = "SELSRAM_DSSC";
        private const string HeaderUSL = "USL";
        private const string HeaderLSL = "LSL";
        private const string HeaderAcCategory = "Ac Category";
        private const string HeaderPowerRunScenario = "Power Run Scenario";
        private const string HeaderRetention = "Retention";
        private const string HeaderDigSrc = "DigSrc";
        private const string HeaderFailCyclePoint = "FC on each point";

        private readonly List<int> _indexPatterns = new List<int>();
        private readonly Dictionary<string, string> _indexPins = new Dictionary<string, string>();
        private int _indexAiType = -1;
        private int _indexComment = -1;
        private int _indexDataLoggingSetting = -1;
        private int _indexOrder = -1;
        private int _indexSearch = -1;
        private int _indexSelsramDssc = -1;
        private int _indexTempCondition = -1;
        private int _indexFailCyclePoint = -1;
        private int _indexTestInstanceName = -1;
        private int _indexTimeset = -1;
        private int _indexUseNotUse = -1;
        private int _indexVoltageCategory = -1;
        private int _indexUSL = -1;
        private int _indexLSL = -1;
        private int _indexAcCategory = -1;
        private int _indexPowerRunScenario = -1;
        private int _indexRetention = -1;
        private int _indexDigSrc = -1;

        private AiTestPlanSheet _sheet;

        public override AiTestPlanSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var sheetName = excelWorksheet.Name;

            _sheet = new AiTestPlanSheet(sheetName);

            ExcelWorksheet = excelWorksheet;

            if (!GetDimensions())
            {
                _sheet.AddDimensionError();
                return _sheet;
            }

            if (!GetFirstHeaderPosition())
            {
                _sheet.AddFirstHeaderError(HeaderUseNotUse);
                return _sheet;
            }

            GetHeaderIndex();

            return ReadSheet(sheetName);
        }

        private AiTestPlanSheet ReadSheet(string sheetName)
        {
            var sheet = new AiTestPlanSheet(sheetName);
            for (var i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new AiTestPlanRow(sheetName);
                row.RowNum = i;
                PopulateRowColumns(row, i);
                PopulatePatterns(row, i);
                PopulatePins(row, i);

                if (!string.IsNullOrEmpty(row.UseNotUse) &&
                    row.UseNotUse.Equals("use", StringComparison.OrdinalIgnoreCase))
                {
                    sheet.Rows.Add(row);
                }
            }

            sheet.IndexUseNotUse = _indexUseNotUse;
            sheet.IndexComment = _indexComment;
            sheet.IndexTestInstanceName = _indexTestInstanceName;
            sheet.IndexAiType = _indexAiType;
            sheet.IndexDataLoggingSetting = _indexDataLoggingSetting;
            sheet.IndexTimeset = _indexTimeset;
            sheet.IndexVoltageCategory = _indexVoltageCategory;
            sheet.IndexOrder = _indexOrder;
            sheet.IndexSearch = _indexSearch;
            sheet.IndexTempCondition = _indexTempCondition;
            sheet.IndexFailCyclePoint = _indexFailCyclePoint;
            sheet.IndexSelsramDssc = _indexSelsramDssc;
            sheet.IndexPatternStart = _indexPatterns.First();
            sheet.IndexStartRow = StartRow;
            sheet.IndexPowerRunScenario = _indexPowerRunScenario;
            sheet.IndexAcCategory = _indexAcCategory;
            sheet.IndexRetention = _indexRetention;
            sheet.IndexDigSrc = _indexDigSrc;
            sheet.IndexUSL = _indexUSL;
            sheet.IndexLSL = _indexLSL;

            return sheet;
        }

        private void PopulateRowColumns(AiTestPlanRow row, int i)
        {
            if (_indexUseNotUse != -1)
            {
                row.UseNotUse = ExcelWorksheet.GetMergeCellValue(i, _indexUseNotUse).Trim();
            }

            if (_indexComment != -1)
            {
                row.Comment = ExcelWorksheet.GetMergeCellValue(i, _indexComment).Trim();
            }

            if (_indexTestInstanceName != -1)
            {
                row.TestInstanceName = ExcelWorksheet.GetMergeCellValue(i, _indexTestInstanceName).Trim();
            }

            if (_indexAiType != -1)
            {
                row.AiType = ExcelWorksheet.GetMergeCellValue(i, _indexAiType).Trim();
            }

            if (_indexDataLoggingSetting != -1)
            {
                row.DataLoggingSetting = ExcelWorksheet.GetMergeCellValue(i, _indexDataLoggingSetting).Trim();
            }

            if (_indexTimeset != -1)
            {
                row.Timeset = ExcelWorksheet.GetMergeCellValue(i, _indexTimeset).Trim();
            }

            if (_indexVoltageCategory != -1)
            {
                row.VoltageCategory = ExcelWorksheet.GetMergeCellValue(i, _indexVoltageCategory).Trim();
            }

            if (_indexOrder != -1)
            {
                row.Order = ExcelWorksheet.GetMergeCellValue(i, _indexOrder).Trim().Replace("_", "");
            }

            if (_indexSearch != -1)
            {
                row.Search = ExcelWorksheet.GetMergeCellValue(i, _indexSearch).Trim();
            }

            if (_indexTempCondition != -1)
            {
                row.TempCondition = ExcelWorksheet.GetMergeCellValue(i, _indexTempCondition).Trim();
            }
            if (_indexFailCyclePoint != -1)
            {
                row.FailCycleEachPoint = ExcelWorksheet.GetMergeCellValue(i, _indexFailCyclePoint).Trim();
            }

            if (_indexSelsramDssc != -1)
            {
                row.SelsramDssc = ExcelWorksheet.GetMergeCellValue(i, _indexSelsramDssc).Trim();
            }

            if (_indexUSL != -1)
            {
                row.USL = ExcelWorksheet.GetMergeCellValue(i, _indexUSL).Trim();
            }

            if (_indexLSL != -1)
            {
                row.LSL = ExcelWorksheet.GetMergeCellValue(i, _indexLSL).Trim();
            }

            if (_indexAcCategory != -1)
            {
                row.AcCategory = ExcelWorksheet.GetMergeCellValue(i, _indexAcCategory).Trim();
            }

            if (_indexPowerRunScenario != -1)
            {
                row.PowerRunScenario = ExcelWorksheet.GetMergeCellValue(i, _indexPowerRunScenario).Trim();
            }

            if (_indexRetention != -1)
            {
                row.Retention = ExcelWorksheet.GetMergeCellValue(i, _indexRetention).Trim();
            }

            if (_indexDigSrc != -1)
            {
                row.DigSrc = ExcelWorksheet.GetMergeCellValue(i, _indexDigSrc).Trim();
            }
        }

        private void PopulatePatterns(AiTestPlanRow row, int i)
        {
            var initStop = false;
            var initIndex = 1;
            var plIndex = 1;
            foreach (var index in _indexPatterns)
            {
                var patternCell = ExcelWorksheet.GetMergeCellValue(i, index).Trim();
                var patternIndex = Regex.Match(ExcelWorksheet.GetMergeCellValue(StartRow, index).Trim(), HeaderPattern).Groups["index"].Value;
                if (!string.IsNullOrEmpty(patternCell))
                {
                    foreach (var pattern in patternCell.Split(','))
                    {
                        var applyPattern = new PatternDate(pattern, patternIndex);
                        if (pattern.IsInit() && initStop == false)
                        {
                            applyPattern.SubIndex = "INIT" + initIndex.ToString();
                            row.Inits.Add(applyPattern);
                            row.IndexMappingPattern = index;
                            initIndex++;
                        }
                        else
                        {
                            initStop = true;
                            applyPattern.SubIndex = "PL" + plIndex.ToString();
                            row.Payloads.Add(applyPattern);
                            if (plIndex == 1)
                                row.IndexMappingPattern = index;
                            plIndex++;
                        }
                    }
                }
            }
        }

        private void PopulatePins(AiTestPlanRow row, int i)
        {
            foreach (var index in _indexPins)
            {
                var pin = new Pin();
                pin.Name = index.Value.Replace("_", "");
                var arr = index.Key.Split(';').ToList();
                int num;
                if (int.TryParse(arr.First(), out num))
                {
                    pin.Start = ExcelWorksheet.GetMergeCellValue(i, num).Trim();
                    pin.IndexStart = num;
                }

                if (arr.Count >= 2 && !string.IsNullOrEmpty(arr.ElementAt(1)))
                {
                    if (int.TryParse(arr.ElementAt(1), out num))
                    {
                        pin.Stop = ExcelWorksheet.GetMergeCellValue(i, num).Trim();
                        pin.IndexStop = num;
                    }
                }

                if (arr.Count >= 3 && !string.IsNullOrEmpty(arr.ElementAt(2)))
                {
                    if (int.TryParse(arr.ElementAt(2), out num))
                    {
                        pin.Step = ExcelWorksheet.GetMergeCellValue(i, num).Trim();
                        pin.IndexStep = num;
                    }
                }

                row.Pins.Add(pin);
            }
        }

        private void GetHeaderIndex()
        {
            for (var i = StartCol; i <= EndCol; i++)
            {
                var header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (CheckHeaderString(header, HeaderUseNotUse))
                {
                    _indexUseNotUse = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderComment))
                {
                    _indexComment = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderTestInstanceName))
                {
                    _indexTestInstanceName = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderAiType))
                {
                    _indexAiType = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderDataLoggingSetting))
                {
                    _indexDataLoggingSetting = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderTimeset))
                {
                    _indexTimeset = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderVoltageCategory))
                {
                    _indexVoltageCategory = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderOrder))
                {
                    _indexOrder = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderSearch))
                {
                    _indexSearch = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderTempCondition))
                {
                    _indexTempCondition = i;
                    continue;
                }
                if (CheckHeaderString(header, HeaderFailCyclePoint))
                {
                    _indexFailCyclePoint = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderSelsramDssc))
                {
                    _indexSelsramDssc = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderUSL))
                {
                    _indexUSL = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderLSL))
                {
                    _indexLSL = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderAcCategory))
                {
                    _indexAcCategory = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderPowerRunScenario))
                {
                    _indexPowerRunScenario = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderRetention))
                {
                    _indexRetention = i;
                    continue;
                }

                if (CheckHeaderString(header, HeaderDigSrc))
                {
                    _indexDigSrc = i;
                    continue;
                }

                if (Regex.IsMatch(header, HeaderPattern, RegexOptions.IgnoreCase))
                {
                    _indexPatterns.Add(i);
                    continue;
                }

                var topHeader = ExcelWorksheet.GetCellValue(StartRow - 1, i).Trim();
                if (Regex.IsMatch(header, HeaderStart, RegexOptions.IgnoreCase) &&
                    !string.IsNullOrEmpty(topHeader))
                {
                    var indexStop = "";
                    var header1 = ExcelWorksheet.GetCellValue(StartRow, i + 1).Trim();
                    if (Regex.IsMatch(header1, HeaderStop, RegexOptions.IgnoreCase))
                    {
                        indexStop = (i + 1).ToString();
                    }

                    var indexStep = "";
                    var header2 = ExcelWorksheet.GetCellValue(StartRow, i + 2).Trim();
                    if (Regex.IsMatch(header2, HeaderStep, RegexOptions.IgnoreCase))
                    {
                        indexStep = (i + 2).ToString();
                    }

                    _indexPins.Add(i + ";" + indexStop + ";" + indexStep, topHeader);
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = EndRow > 10 ? 10 : EndRow;
            var colNum = EndCol > 10 ? 10 : EndCol;
            for (var i = 1; i <= rowNum; i++)
                for (var j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim()
                        .Equals(HeaderUseNotUse, StringComparison.OrdinalIgnoreCase))
                    {
                        StartRow = i;
                        return true;
                    }
                }

            return false;
        }

        private bool CheckHeaderString(string header, string headerSyntax)
        {
            header = header.Replace(" ", "").ToLower();
            headerSyntax = headerSyntax.Replace(" ", "").ToLower();
            return header == headerSyntax;
        }
    }
}
