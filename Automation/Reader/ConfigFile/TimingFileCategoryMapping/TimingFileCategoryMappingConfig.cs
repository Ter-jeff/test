using System.Collections.Generic;

using CommonLib.Extension;

using OfficeOpenXml;

namespace Automation.Reader.ConfigFile.TimingFileCategoryMapping
{
    public class TimingFileCategoryMappingConfig
    {

        private readonly Dictionary<string, string> _categoryMapping = new Dictionary<string, string>();

        public string GetCategoryMapping(string timingSheetName)
        {

            if (_categoryMapping.ContainsKey(timingSheetName.ToUpper()))
            {
                return _categoryMapping[timingSheetName.ToUpper()];
            }

            return "";
        }

        public static TimingFileCategoryMappingConfig LoasConfig(ExcelWorksheet mappingSheet)
        {
            TimingFileCategoryMappingConfig table = new TimingFileCategoryMappingConfig();
            if (mappingSheet == null)
            {
                return table;
            }

            for (int i = 2; i <= mappingSheet.Dimension.End.Row; i++)
            {

                for (int j = 1; j <= mappingSheet.Dimension.End.Column; j++)
                {
                    string timingFile = EpplusExtensions.GetCellValue(mappingSheet, i, j).Trim().ToUpper();
                    string categoryName = EpplusExtensions.GetCellValue(mappingSheet, 1, j).Trim().ToUpper();

                    if (string.IsNullOrEmpty(timingFile) || string.IsNullOrEmpty(categoryName))
                    {
                        continue;
                    }

                    table._categoryMapping[timingFile] = categoryName;
                }
            }

            return table;
        }
    }
}
