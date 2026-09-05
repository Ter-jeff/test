using System.Collections.Generic;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.BinNumberLegacy
{
    public partial class SoftBinDataReader(ExcelWorkbook excelWorkbook)
    {
        private const string ConCategory = "Category";
        private const string ConModule = "Module";
        private const string ConSubModule = "SubModule";
        private const string ConBlock = "Block";
        private const string ConLevel = "Level";
        private const string ConRefer = @"See\s*(?<str>.*)";

        [GeneratedRegex(ConRefer, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(ConRefer)]
        private static partial Regex MyRegex1();

        private readonly ExcelWorkbook _binNumerWorkbook = excelWorkbook;

        private List<string> _categoryList = [];
        private List<string> _moduleList = [];
        private List<string> _subModuleList = [];
        private List<string> _blockList = [];
        private List<string> _levelList = [];

        public void ReadModuleList(ExcelWorksheet excelWorksheet)
        {
            int columnModule = -1;
            int columnSubModule = -1;
            int columnBlock = -1;
            int columnLevel = -1;
            int columnCategory = -1;
            _categoryList = [];
            _moduleList = [];
            _subModuleList = [];
            _blockList = [];
            _levelList = [];
            for (int i = 1; i <= excelWorksheet.Dimension.End.Column; i++)
            {
                if (EpplusExtensions.GetCellValue(excelWorksheet, 1, i).EqualsIgnoreCase(ConCategory))
                {
                    columnCategory = i;
                }
                if (EpplusExtensions.GetCellValue(excelWorksheet, 1, i).EqualsIgnoreCase(ConModule))
                {
                    columnModule = i;
                }
                else if (EpplusExtensions.GetCellValue(excelWorksheet, 1, i).EqualsIgnoreCase(ConSubModule))
                {
                    columnSubModule = i;
                }
                else if (EpplusExtensions.GetCellValue(excelWorksheet, 1, i).EqualsIgnoreCase(ConBlock))
                {
                    columnBlock = i;
                }
                else if (EpplusExtensions.GetCellValue(excelWorksheet, 1, i).EqualsIgnoreCase(ConLevel))
                {
                    columnLevel = i;
                }
            }

            for (int i = 2; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                string category = ReadConfigCell(excelWorksheet, i, columnCategory);
                string module = ReadConfigCell(excelWorksheet, i, columnModule);
                string subModule = ReadConfigCell(excelWorksheet, i, columnSubModule);
                string block = ReadConfigCell(excelWorksheet, i, columnBlock);
                string level = ReadConfigCell(excelWorksheet, i, columnLevel);
                if (category.Length != 0)
                {
                    _categoryList.Add(category);
                }
                if (module.Length != 0)
                {
                    _moduleList.Add(module);
                }
                if (subModule.Length != 0)
                {
                    _subModuleList.Add(subModule);
                }
                if (block.Length != 0)
                {
                    _blockList.Add(block);
                }
                if (level.Length != 0)
                {
                    _levelList.Add(level);
                }
            }
        }

        public List<SoftBinDigiDef> ReadSheet(ExcelWorksheet excelWorksheet)
        {
            List<SoftBinDigiDef> softBinData = [];

            for (int i = 2; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                string range = excelWorksheet.MergedCells[i, 1];
                int start = i;
                int end = string.IsNullOrEmpty(range) ? i : new ExcelAddress(range).End.Row;
                SoftBinDigiDef? digitalDef = ReadOneCat(excelWorksheet, start, end, 4);
                if (digitalDef != null)
                {
                    softBinData.Add(digitalDef);
                }
            }
            return softBinData;
        }

        private SoftBinDigiDef? ReadOneCat(ExcelWorksheet excelWorksheet, int startRow, int endRow, int digiNum)
        {
            SoftBinDigiDef digiDef = new SoftBinDigiDef();
            string category = ReadConfigCell(excelWorksheet, startRow, 1);
            digiDef.Category = category;
            if (category.Length == 0)
            {
                return null;
            }

            digiDef.CategoryType = GetBinNumKeyType(category);
            for (int i = 0; i < digiNum; i++)
            {
                DigitalDef digital = new DigitalDef();
                for (int j = startRow; j <= endRow; j++)
                {
                    string keyword = EpplusExtensions.GetMergedCellValue(excelWorksheet, j, (i * 2) + 2).Trim();
                    string number = EpplusExtensions.GetCellValue(excelWorksheet, j, (i * 2) + 3).Trim();

                    if (MyRegex().IsMatch(keyword))
                    {
                        string referSheetName = MyRegex1().Match(keyword).Groups["str"].ToString();
                        if (!digiDef.DigitalList.Exists(p => referSheetName == p.ReferSheet))
                        {
                            digital.ReferSheet = MyRegex1().Match(keyword).Groups["str"].ToString();
                            ExcelWorksheet referSheet = _binNumerWorkbook.Worksheets[digital.ReferSheet];
                            SoftBinDetailReader reader = new SoftBinDetailReader();
                            digiDef.ReferDetail = SoftBinDetailReader.ReadSheet(referSheet);
                        }
                    }
                    if (keyword.Length != 0 && !digital.NumberDefList.Exists(p => p.Keyword.EqualsIgnoreCase(keyword)))
                    {
                        EnumBinNumKeyType type = GetBinNumKeyType(keyword);
                        digital.AddNumDef(keyword, type, number);
                    }
                }
                digiDef.DigitalList.Add(digital);
            }

            return digiDef;
        }

        private EnumBinNumKeyType GetBinNumKeyType(string key)
        {
            if (_categoryList.Exists(p => p.EqualsIgnoreCase(key)))
            {
                return EnumBinNumKeyType.Category;
            }
            if (_moduleList.Exists(p => p.EqualsIgnoreCase(key)))
            {
                return EnumBinNumKeyType.Module;
            }
            if (_subModuleList.Exists(p => p.EqualsIgnoreCase(key)))
            {
                return EnumBinNumKeyType.SubModule;
            }
            if (_blockList.Exists(p => p.EqualsIgnoreCase(key)))
            {
                return EnumBinNumKeyType.Block;
            }
            return _levelList.Exists(p => p.EqualsIgnoreCase(key)) ? EnumBinNumKeyType.Level : EnumBinNumKeyType.NonType;
        }

        private static string ReadConfigCell(ExcelWorksheet excelWorksheet, int row, int column)
        {
            return EpplusExtensions.GetCellValue(excelWorksheet, row, column).Trim();
        }

    }
}
