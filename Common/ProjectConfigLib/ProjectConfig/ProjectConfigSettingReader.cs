using System;
using System.Collections.Generic;
using System.IO;

using CommonLib.Extension;

using OfficeOpenXml;

namespace ProjectConfigLib.ProjectConfig
{
    public class ProjectConfigSettingReader
    {
        public static List<ProjectConfigSettingRow>? Read(string projectName = "")
        {
            string sheetsFormatFile = Path.Combine(AppContext.BaseDirectory, "Config", "ProjectConfigSetting.xlsx");
            if (!string.IsNullOrEmpty(projectName))
            {
                sheetsFormatFile = projectName;
            }

            if (!File.Exists(sheetsFormatFile))
            {
                return null;
            }

            var inputExcel = new ExcelPackage(new FileInfo(sheetsFormatFile));
            ExcelWorkbook inputWBook = inputExcel.Workbook;
            ExcelWorksheet workSheet = inputWBook.Worksheets["ProjectConfig"];
            var projectConfigReader = new ProjectConfigReader();
            ProjectConfigSheet? sheet = projectConfigReader.ReadSheet(workSheet);
            inputExcel.Dispose();
            return sheet?.Rows;
        }
    }

    public class ProjectConfigSettingRow
    {
        public string SourceSheetName { set; get; } = "";
        public int RowNum { get; set; }
        public string Name { set; get; }
        public string GroupName { set; get; }
        public string TabGroup { set; get; }
        public string Description { set; get; }
        public string Type { set; get; }
        public int MaxLength { set; get; }
        public string Default { set; get; }
        public List<string> OptionValue = [];

        public ProjectConfigSettingRow()
        {
            Name = "";
            GroupName = "";
            TabGroup = "";
            Description = "";
            Type = "";
            Default = "";
        }

        public ProjectConfigSettingRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Name = "";
            GroupName = "";
            TabGroup = "";
            Description = "";
            Type = "";
            Default = "";
        }

        public ProjectConfigRow ConvertToProjectConfigRow()
        {
            var projectConfigRow = new ProjectConfigRow
            {
                Group = GroupName,
                Name = Name,
                Value = Default
            };
            return projectConfigRow;
        }
    }

    public class ProjectConfigSheet(string sheetName)
    {
        public string SheetName { get; set; } = sheetName;
        public List<ProjectConfigSettingRow> Rows { get; } = [];
        public Dictionary<string, int> HeaderIndex { get; } = [];

        #region Constructor
        #endregion
    }

    public class ProjectConfigReader
    {
        private const string ConHeaderName = "Name";
        private const string ConHeaderGroupName = "GroupName";
        private const string ConHeaderDevice = "TabGroup";
        private const string ConHeaderDescription = "Description";
        private const string ConHeaderType = "Type";
        private const string ConHeaderMaxlength = "MaxLength";
        private const string ConHeaderDefault = "Default";
        private const string ConHeaderOptionValue1 = "OptionValue1";
        private const string ConHeaderOptionValue2 = "OptionValue2";
        private const string ConHeaderOptionValue3 = "OptionValue3";
        private const string ConHeaderOptionValue4 = "OptionValue4";
        private const string ConHeaderOptionValue5 = "OptionValue5";
        private const string ConHeaderOptionValue6 = "OptionValue6";
        private const string ConHeaderOptionValue7 = "OptionValue7";
        private const string ConHeaderOptionValue8 = "OptionValue8";
        private const string ConHeaderOptionValue9 = "OptionValue9";
        private const string ConHeaderOptionValue10 = "OptionValue10";

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _nameIndex = -1;
        private int _groupNameIndex = -1;
        private int _deviceIndex = -1;
        private int _descriptionIndex = -1;
        private int _typeIndex = -1;
        private int _maxlengthIndex = -1;
        private int _defaultIndex = -1;
        private int _optionValue1Index = -1;
        private int _optionValue2Index = -1;
        private int _optionValue3Index = -1;
        private int _optionValue4Index = -1;
        private int _optionValue5Index = -1;
        private int _optionValue6Index = -1;
        private int _optionValue7Index = -1;
        private int _optionValue8Index = -1;
        private int _optionValue9Index = -1;
        private int _optionValue10Index = -1;

        private ExcelWorksheet _excelWorksheet = null!;
        private string _sheetName = null!;
        private ProjectConfigSheet _projectConfigSheet = null!;

        private readonly Dictionary<string, bool> _headerOptional = new()
        {
            { "Name", true }, { "GroupName", true }, { "Description", true }, { "Type", true }, { "MaxLength", true }, { "Default", true }, { "OptionValue1", true }, { "OptionValue2", true }, { "OptionValue3", true }, { "OptionValue4", true }, { "OptionValue5", true }, { "OptionValue6", true }, { "OptionValue7", true }, { "OptionValue8", true }, { "OptionValue9", true }, { "OptionValue10", true }
        };

        public ProjectConfigSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            _excelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;

            _projectConfigSheet = new ProjectConfigSheet(_sheetName);

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            if (!GetHeaderIndex())
            {
                return null;
            }

            _projectConfigSheet = ReadSheetData();

            return _projectConfigSheet;
        }

        private ProjectConfigSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new ProjectConfigSettingRow(_sheetName)
                {
                    RowNum = i
                };
                if (_nameIndex != -1)
                {
                    row.Name = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _nameIndex).Trim();
                }

                if (_groupNameIndex != -1)
                {
                    row.GroupName = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _groupNameIndex).Trim();
                }

                if (_deviceIndex != -1)
                {
                    row.TabGroup = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _deviceIndex).Trim();
                }

                if (_descriptionIndex != -1)
                {
                    row.Description = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _descriptionIndex).Trim();
                }

                if (_typeIndex != -1)
                {
                    row.Type = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _typeIndex).Trim();
                }

                if (_maxlengthIndex != -1)
                {
                    _ = int.TryParse(EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _maxlengthIndex).Trim(), out int value);
                    row.MaxLength = value;
                }
                if (_defaultIndex != -1)
                {
                    row.Default = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _defaultIndex).Trim();
                }

                if (_optionValue1Index != -1)
                {
                    row.OptionValue.Add(EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _optionValue1Index).Trim());
                }

                if (_optionValue2Index != -1)
                {
                    row.OptionValue.Add(EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _optionValue2Index).Trim());
                }

                if (_optionValue3Index != -1)
                {
                    row.OptionValue.Add(EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _optionValue3Index).Trim());
                }

                if (_optionValue4Index != -1)
                {
                    row.OptionValue.Add(EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _optionValue4Index).Trim());
                }

                if (_optionValue5Index != -1)
                {
                    row.OptionValue.Add(EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _optionValue5Index).Trim());
                }

                if (_optionValue6Index != -1)
                {
                    row.OptionValue.Add(EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _optionValue6Index).Trim());
                }

                if (_optionValue7Index != -1)
                {
                    row.OptionValue.Add(EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _optionValue7Index).Trim());
                }

                if (_optionValue8Index != -1)
                {
                    row.OptionValue.Add(EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _optionValue8Index).Trim());
                }

                if (_optionValue9Index != -1)
                {
                    row.OptionValue.Add(EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _optionValue9Index).Trim());
                }

                if (_optionValue10Index != -1)
                {
                    row.OptionValue.Add(EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _optionValue10Index).Trim());
                }

                _projectConfigSheet.Rows.Add(row);
            }
            return _projectConfigSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = EpplusExtensions.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.EqualsIgnoreCase(ConHeaderName))
                {
                    _nameIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderName, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderGroupName))
                {
                    _groupNameIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderGroupName, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderDevice))
                {
                    _deviceIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderDevice, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderDescription))
                {
                    _descriptionIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderDescription, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderType))
                {
                    _typeIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderType, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderMaxlength))
                {
                    _maxlengthIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderMaxlength, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderDefault))
                {
                    _defaultIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderDefault, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderOptionValue1))
                {
                    _optionValue1Index = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderOptionValue1, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderOptionValue2))
                {
                    _optionValue2Index = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderOptionValue2, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderOptionValue3))
                {
                    _optionValue3Index = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderOptionValue3, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderOptionValue4))
                {
                    _optionValue4Index = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderOptionValue4, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderOptionValue5))
                {
                    _optionValue5Index = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderOptionValue5, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderOptionValue6))
                {
                    _optionValue6Index = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderOptionValue6, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderOptionValue7))
                {
                    _optionValue7Index = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderOptionValue7, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderOptionValue8))
                {
                    _optionValue8Index = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderOptionValue8, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderOptionValue9))
                {
                    _optionValue9Index = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderOptionValue9, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderOptionValue10))
                {
                    _optionValue10Index = i;
                    _projectConfigSheet.HeaderIndex.Add(ConHeaderOptionValue10, i);
                }
            }

            foreach (KeyValuePair<string, int> header in _projectConfigSheet.HeaderIndex)
            {
                if (header.Value == -1 && _headerOptional.TryGetValue(header.Key, out bool value) && value)
                {
                    return false;
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (EpplusExtensions.GetCellValue(_excelWorksheet, i, j).Trim().EqualsIgnoreCase(ConHeaderName))
                    {
                        _startRowNumber = i;
                        return true;
                    }
                }
            }

            return false;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet.Dimension.Start.Row;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }
            return false;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _nameIndex = -1;
            _groupNameIndex = -1;
            _deviceIndex = -1;
            _descriptionIndex = -1;
            _typeIndex = -1;
            _maxlengthIndex = -1;
            _defaultIndex = -1;
            _optionValue1Index = -1;
            _optionValue2Index = -1;
            _optionValue3Index = -1;
            _optionValue4Index = -1;
            _optionValue5Index = -1;
            _optionValue6Index = -1;
            _optionValue7Index = -1;
            _optionValue8Index = -1;
            _optionValue9Index = -1;
            _optionValue10Index = -1;
        }
    }
}
