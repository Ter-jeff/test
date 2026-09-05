using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.DataStruct;
using TestPlanLib.Static;

namespace TestPlanLib.Basic
{
    public partial class TestSettingReader(PinMapSheet? pinMapSheet)
    {
        private const string ConTestSettingBodySymbol = "PinName";
        private const string ConFooterGlbSymbol = "NOP";
        private const string ConFooterVrsSymbol = "NOP";

        [GeneratedRegex("\\(\\s*(?<Unit>\\w+)\\s*\\)")]
        private static partial Regex MyRegex();

        private ExcelWorksheet? _currentSheet;
        private readonly PinMapSheet? _pinMapSheet = pinMapSheet;
        private readonly Dictionary<string, string> _redundantDecomposedPinName = new(StringExtensions.IgnoreCase);
        private int _bodyStartRow;
        private int _bodyStartColumn;
        private int _footerStartRow;
        private int _footerStartColumn;
        private int _currentRow;
        private int _currentColumn;

        #region Methods
        public List<TestSettingData> ReadFlow(string currentProject, ExcelWorkbook excelWorkbook, List<string> allJobs)
        {
            if (allJobs.Count == 0)
            {
                allJobs.Add("CP1");
            }

            return ReadFlowByAllJobs(currentProject, excelWorkbook, allJobs);
        }

        private List<TestSettingRow> DecomposeByPinMap(TestSettingRow testSettingRow, PinMapSheet? pinMapSheet, string sheetName, int rowNum)
        {
            List<TestSettingRow> decomposedTestSettingRows = [testSettingRow];
            if (pinMapSheet == null)
            {
                return decomposedTestSettingRows;
            }

            string singlePinName = testSettingRow.PowerPinName;
            bool isValtPin = singlePinName.EndsWithIgnoreCase("_Valt");
            string valtSyntax = "";

            if (isValtPin)
            {
                valtSyntax = singlePinName[^5..];
                singlePinName = singlePinName[..^5];
            }

            if (pinMapSheet.TryGetGroup(singlePinName, out PinGroup? pinGroup))
            {

                decomposedTestSettingRows = [.. pinGroup.PinList.Select(x =>
                {
                    TestSettingRow decomposedRow = testSettingRow.Copy();
                    decomposedRow.PowerPinName = x.PinName + valtSyntax;
                    return decomposedRow;
                })];

                List<string> decomposedNames = [.. decomposedTestSettingRows.Select(x => x.PowerPinName)];
                List<string> intersect = [.. _redundantDecomposedPinName.Keys.ToHashSet(StringExtensions.IgnoreCase).Intersect(decomposedNames)];
                intersect.ForEach(x =>
                {
                    ErrorReportManager.AddError(BasicErrorType.E_DuplicateItems_03, sheetName, rowNum, 0, $"Single pin: {x} is duplicate in pin groups: {_redundantDecomposedPinName[x]} and {testSettingRow.PowerPinName}, please check the sheet: {sheetName}", [x, _redundantDecomposedPinName[x], sheetName]);
                });
            }
            decomposedTestSettingRows.ForEach(x => _redundantDecomposedPinName[x.PowerPinName] = testSettingRow.PowerPinName);
            return decomposedTestSettingRows;
        }

        public List<TestSettingData> ReadFlowByAllJobs(string currentProject, ExcelWorkbook excelWorkbook, List<string> allJobs, List<Tuple<string, string>>? dcCategories = null)
        {
            //Get all TestSetting sheets([projectname]_(TestSetting|valtageTable)_[job])
            var srcSettingSheetsList = excelWorkbook.Worksheets.Where(s => NeededSheets.IsTestSettingSheetName(s.Name, currentProject)).ToList();

            var testSettingSheets = new List<TestSettingData>();
            foreach (string job in allJobs)
            {
                //Find job testSetting sheet
                ExcelWorksheet? testSettingSheet = (srcSettingSheetsList.Find(s => Regex.Match(s.Name, NeededSheets.TestSettingTbl, RegexOptions.IgnoreCase).Groups["job"].ToString().EqualsIgnoreCase(job)) ?? srcSettingSheetsList.Find(s => Regex.Match(s.Name, NeededSheets.TestSettingTbl, RegexOptions.IgnoreCase).Groups["job"].ToString().EqualsIgnoreCase(job[..^1]))) ?? srcSettingSheetsList.Find(s => Regex.Match(s.Name, NeededSheets.TestSettingTbl, RegexOptions.IgnoreCase).Groups["job"].ToString().EqualsIgnoreCase(""));

                if (testSettingSheet != null)
                {
                    string originJob = Regex.Match(testSettingSheet.Name, NeededSheets.TestSettingTbl, RegexOptions.IgnoreCase).Groups["job"].ToString();
                    TestSettingData? preData = testSettingSheets.Find(x => x.OriginJob == originJob);
                    if (preData == null)
                    {
                        Initialize(ConTestSettingBodySymbol, testSettingSheet);
                        var data = new TestSettingData
                        {
                            Job = job,
                            OriginJob = originJob,
                            BasicUnit = GetBasicUnit(),
                            SheetName = "TestSetting_" + job,
                            TestSettingVersion = EpplusExtensions.GetCellValue(testSettingSheet, _bodyStartRow, _bodyStartColumn)
                        };
                        if (dcCategories != null)
                        {
                            data.DcCategorys = ReadSheetHeader([.. dcCategories.Select(x => x.Item1).Distinct()]);
                        }
                        else
                        {
                            data.DcCategorys = ReadSheetHeader();
                        }

                        data.DataRows = ReadSheetBody(data.DcCategorys, testSettingSheet.Name);
                        data.Footer = ReadSheetFooter();
                        testSettingSheets.Add(data);
                    }
                    else
                    {
                        TestSettingData data = preData.Copy();
                        data.Job = job;
                        data.SheetName = "TestSetting_" + job;
                        testSettingSheets.Add(data);
                    }
                }
            }

            return testSettingSheets;
        }

        /// <summary>
        /// Set current sheet and find data start location before read
        /// </summary>
        /// <param name="bodySymbol">cell value ditermines the data location</param>
        /// <param name="excelWorksheet"></param>
        private void Initialize(string bodySymbol, ExcelWorksheet excelWorksheet)
        {
            _currentSheet = excelWorksheet;
            _bodyStartRow = 0;
            _bodyStartColumn = 0;
            _footerStartRow = 0;
            _footerStartColumn = 0;
            _currentRow = 0;
            _currentColumn = 1;

            if (_currentSheet == null)
            {
                return;
            }

            var bodyMather = new Regex(bodySymbol, RegexOptions.IgnoreCase);

            for (int row = 1; row <= _currentSheet.Dimension.Rows; row++)
            {
                for (int col = 1; col <= _currentSheet.Dimension.Columns; col++)
                {
                    if (bodyMather.IsMatch(EpplusExtensions.GetCellValue(_currentSheet, row, col)))
                    {
                        _bodyStartColumn = col;
                        _bodyStartRow = row - 1;
                        return;
                    }
                }
            }
        }

        private List<DcCategoryName> ReadSheetHeader(List<string>? list = null)
        {
            var result = new List<DcCategoryName>();
            int row = _bodyStartRow;
            for (int col = _bodyStartColumn + 1; col <= _currentSheet!.Dimension.Columns; col++)
            {
                string categoryNameContent = EpplusExtensions.GetCellValue(_currentSheet, row, col);
                if (categoryNameContent.Length == 0)
                {
                    continue;
                }

                var categoryName = new DcCategoryName(categoryNameContent);
                if (list != null)
                {
                    if (!list.Exists(x => x.EqualsIgnoreCase(categoryName.CategoryName)))
                    {
                        continue;
                    }
                }
                //ditermine follow datas is HV, LV or NV of current category
                categoryName.ValueType = GetValueType(EpplusExtensions.GetCellValue(_currentSheet, row + 1, col));
                categoryName.ColumnIndex = col;
                result.Add(categoryName);
            }

            return result;
        }

        /// <summary>
        /// Read TSSheet data part and set origin data for each category
        /// </summary>
        /// <param name="dcCategoryNames">title line, ditermine follow data belongs to which category</param>
        /// <returns>categorys with thier data</returns>
        private List<TestSettingRow> ReadSheetBody(List<DcCategoryName> dcCategoryNames, string sheetName)
        {
            var result = new List<TestSettingRow>();
            int readrows = 0;

            FindSheetBody();

            for (int row = _currentRow; row <= _currentSheet!.Dimension.Rows; row++)
            {
                var currentRow = new TestSettingRow();
                string pinName = EpplusExtensions.GetCellValue(_currentSheet, row, _currentColumn);
                if (pinName.Length == 0)
                {
                    //terminate data read process when meet blank line
                    break;
                }

                currentRow.PowerPinName = pinName;
                var groups = dcCategoryNames.GroupBy(x => x.CategoryName).ToList();
                foreach (IGrouping<string, DcCategoryName> group in groups)
                {
                    var dcCategoryValue = new DcCategoryValue(group.First().CategoryName) { ColumnIndex = group.First().ColumnIndex };
                    foreach (DcCategoryName item in group)
                    {
                        CategoryValueType columnValueType = item.ValueType;
                        int col = item.ColumnIndex;
                        SetDcCategoryValue(columnValueType, dcCategoryValue, EpplusExtensions.GetCellText(_currentSheet, row, col), EpplusExtensions.GetCellFormula(_currentSheet, row, col));
                    }
                    currentRow.DcCategoryValues.Add(dcCategoryValue);
                }
                result.AddRange(DecomposeByPinMap(currentRow, _pinMapSheet, sheetName, row));
                readrows++;
            }
            //save current row index after read all body data, perhaps it has other data below
            _currentRow += readrows;
            return result;
        }

        /// <summary>
        /// Read TSSheet footer and get global pin infos
        /// </summary>
        /// <returns>footer pin infos</returns>
        private List<TestSettingFooterRow> ReadSheetFooter()
        {
            var result = new List<TestSettingFooterRow>();
            int readRows = 0;

            //find the location of global pins
            FindNextBody(ConFooterGlbSymbol, ref _footerStartRow, ref _footerStartColumn);
            if (_footerStartRow == 0)
            {
                return result;
            }

            for (int row = _footerStartRow + 1; row <= _currentSheet!.Dimension.Rows; row++)
            {
                var currentGlbRow = new TestSettingFooterRow();
                string pinName = EpplusExtensions.GetCellValue(_currentSheet, row, _footerStartColumn);
                if (pinName.Length == 0)
                {
                    //terminate data read process when meet blank line
                    break;
                }
                currentGlbRow.PinName = pinName;

                currentGlbRow.Value = EpplusExtensions.GetCellText(_currentSheet, row, _footerStartColumn + 1);
                currentGlbRow.Formula = EpplusExtensions.GetCellFormula(_currentSheet, row, _footerStartColumn + 1);
                //save footer pin address because some other footer pin refer the cell addres of current pin
                currentGlbRow.Address = GetCellAddress(_currentSheet, row, _footerStartColumn + 1);
                currentGlbRow.Comment = EpplusExtensions.GetCellValue(_currentSheet, row, _footerStartColumn + 2);
                result.Add(currentGlbRow);

                readRows++;
            }

            _currentRow = _footerStartRow + readRows;
            //find the location of vrs pins
            FindNextBody(ConFooterVrsSymbol, ref _currentRow, ref _currentColumn);

            for (int row = _currentRow + 1; row <= _currentSheet.Dimension.Rows; row++)
            {
                var currentVrsRow = new TestSettingFooterRow();
                string pinName = EpplusExtensions.GetCellValue(_currentSheet, row, _footerStartColumn);
                if (pinName.Length == 0)
                {
                    //terminate data read process when meet blank line
                    break;
                }
                currentVrsRow.PinName = pinName;

                currentVrsRow.Value = EpplusExtensions.GetCellText(_currentSheet, row, _footerStartColumn + 1);
                //save footer pin address because some other footer pin refer the cell addres of current pin
                currentVrsRow.Address = GetCellAddress(_currentSheet, row, _footerStartColumn + 1);
                currentVrsRow.Comment = EpplusExtensions.GetCellValue(_currentSheet, row, _footerStartColumn + 2);
                result.Add(currentVrsRow);
            }
            return result;
        }

        private static void SetDcCategoryValue(CategoryValueType categoryValueType, DcCategoryValue dcCategoryValue, string value, string formula)
        {
            if (categoryValueType == CategoryValueType.NV)
            {
                dcCategoryValue.Nv.OriginValue = value;
                dcCategoryValue.Nv.Value = value;
                dcCategoryValue.Nv.Formula = formula;
            }
            else if (categoryValueType == CategoryValueType.LV)
            {
                dcCategoryValue.Lv.OriginValue = value;
                dcCategoryValue.Lv.Value = value;
                dcCategoryValue.Lv.Formula = formula;
            }
            else
            {
                dcCategoryValue.Hv.OriginValue = value;
                dcCategoryValue.Hv.Value = value;
                dcCategoryValue.Hv.Formula = formula;
            }
        }

        private void FindSheetBody()
        {
            _currentRow = _bodyStartRow + 2;
            _currentColumn = _bodyStartColumn;
        }

        private void FindNextBody(string symbol, ref int rowCoordination, ref int colCoordination)
        {
            if (symbol == "NOP")
            {
                return;
            }
            var symbolMather = new Regex(symbol, RegexOptions.IgnoreCase);

            for (int row = _currentRow; row <= _currentSheet!.Dimension.Rows; row++)
            {
                for (int col = _currentColumn; col <= _currentSheet.Dimension.Columns; col++)
                {
                    if (symbolMather.IsMatch(EpplusExtensions.GetCellValue(_currentSheet, row, col)))
                    {
                        rowCoordination = row;
                        colCoordination = col;
                        return;
                    }
                }
            }
        }

        private static CategoryValueType GetValueType(string value)
        {
            if (value.ContainsIgnoreCase("LV"))
            {
                return CategoryValueType.LV;
            }
            return value.ContainsIgnoreCase("HV") ? CategoryValueType.HV : CategoryValueType.NV;
        }

        private static string GetCellAddress(ExcelWorksheet excelWorksheet, int row, int col)
        {
            string address = excelWorksheet.Cells[row, col].Address;
            return address;
        }

        private EnumTestSettingBasicUnit GetBasicUnit()
        {
            string cellContent = EpplusExtensions.GetCellValue(_currentSheet!, _bodyStartRow + 1, _bodyStartColumn);
            Regex unitReg = MyRegex();
            if (unitReg.IsMatch(cellContent))
            {
                Match matcher = unitReg.Match(cellContent);
                string unit = matcher.Groups["Unit"].Value;
                if (unit.EqualsIgnoreCase("mv"))
                {
                    return EnumTestSettingBasicUnit.mV;
                }

                if (unit.EqualsIgnoreCase("V"))
                {
                    return EnumTestSettingBasicUnit.V;
                }
            }
            return EnumTestSettingBasicUnit.mV;
        }
        #endregion
    }
}
