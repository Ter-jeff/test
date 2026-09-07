using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.Utility
{
    public static class GetAllPinsInfo
    {
        public class AllPins(Dictionary<string, int> powerPins, Dictionary<string, int> affiliatedPins)
        {
            public Dictionary<string, int> PowerPins { get; private set; } = powerPins;
            public Dictionary<string, int> AffiliatedPins { get; private set; } = affiliatedPins;
        }

        /// <summary>
        /// Get all pins sheet info. <br>
        /// Currently only get `Power Pins` and `Affiliated Pins`
        /// </summary>
        /// <param name="subBinCutSheetInfo"></param>
        /// <returns></returns>
        public static AllPins GetAllPins(SubBinCutSheetInfo subBinCutSheetInfo, int startRow, ExcelWorksheet excelWorksheet)
        {
            int powerRow = startRow + 1;
            var powerPins = new Dictionary<string, int>();
            var affiliatedPins = new Dictionary<string, int>();
            for (int i = subBinCutSheetInfo.JobColumnNum; i <= subBinCutSheetInfo.EndColNum; i++)
            {
                string ingnore = excelWorksheet.GetCellValue(1, i);
                if (ingnore.EqualsIgnoreCase("IGNORE COLUMN"))
                {
                    continue;
                }

                string vddPinsCellValue = excelWorksheet.GetCellValue(powerRow, i).Trim();
                bool hasComma = vddPinsCellValue.Contains(',');
                string powerPin = vddPinsCellValue;
                if (hasComma)
                {
                    string affiliated = vddPinsCellValue.Split(',').Last();

                    affiliatedPins.TryAdd(affiliated, i);
                    powerPin = vddPinsCellValue.Split(',').First();
                }
                if (!powerPin.StartsWithIgnoreCase("VDD_"))
                {
                    continue;
                }

                powerPins.TryAdd(powerPin, i);
            }
            var allPins = new AllPins(powerPins, affiliatedPins);
            return allPins;
        }
    }
}
