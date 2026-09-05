using System.Collections.Generic;

using Automation.InputManager.Data;
using Automation.Static;

using OfficeOpenXml;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.InputReader
{
    public class BlockMappingReader
    {
        public HardIpInputData HardIpInputData { get; }

        public BlockMappingReader(HardIpInputData hardIpInputData)
        {
            HardIpInputData = hardIpInputData;
        }

        public (Dictionary<string, string> VbtNameMapping, Dictionary<string, string> InstSpecialSetting) ConfigReader()
        {
            Dictionary<string, string> vbtNameMapping = null;
            Dictionary<string, string> instSpecialSetting = null;
            if (SettingStatic.BasicConfigWorkbook != null)
            {
                ExcelWorksheet vbtNameMappingSheet = SettingStatic.BasicConfigWorkbook.Worksheets[NeededSheets.VbtNameMap];
                if (vbtNameMappingSheet != null)
                {
                    vbtNameMapping = ReadVbtMapping(vbtNameMappingSheet);
                    HardIpInputData.ConfigData.VbtNameMapping = ReadVbtMapping(vbtNameMappingSheet);
                }

                ExcelWorksheet specialSetupMappingSheet = SettingStatic.BasicConfigWorkbook.Worksheets[NeededSheets.InstSpecialSetupMap];
                if (specialSetupMappingSheet != null)
                {
                    instSpecialSetting = ReadInstSpecialSetupMapping(specialSetupMappingSheet);
                    HardIpInputData.ConfigData.InstSpecialSetting = ReadInstSpecialSetupMapping(specialSetupMappingSheet);
                }
            }

            return (vbtNameMapping, instSpecialSetting);
        }

        private Dictionary<string, string> ReadVbtMapping(ExcelWorksheet sheet)
        {
            var data = new Dictionary<string, string>();
            int endRow = sheet.Dimension.End.Row;
            for (int i = 2; i <= endRow; i++)
            {
                object standardValue = sheet.Cells[i, 1].Value;
                object aliasValue = sheet.Cells[i, 2].Value;
                if (standardValue != null && aliasValue != null)
                {
                    string standardName = standardValue.ToString();
                    string aliasName = aliasValue.ToString();
                    if (standardName != "" && !data.ContainsKey(standardName))
                    {
                        data.Add(standardName, aliasName);
                    }
                }
            }
            return data;
        }

        private Dictionary<string, string> ReadInstSpecialSetupMapping(ExcelWorksheet sheet)
        {
            var data = new Dictionary<string, string>();
            int endRow = sheet.Dimension.End.Row;
            for (int i = 2; i <= endRow; i++)
            {
                object standardValue = sheet.Cells[i, 1].Value;
                object value = sheet.Cells[i, 2].Value;
                if (standardValue != null && value != null)
                {
                    string standardName = standardValue.ToString();
                    string text = value.ToString();
                    if (standardName != "" && !data.ContainsKey(standardName))
                    {
                        data.Add(standardName, text);
                    }
                }
            }
            return data;
        }

    }
}
