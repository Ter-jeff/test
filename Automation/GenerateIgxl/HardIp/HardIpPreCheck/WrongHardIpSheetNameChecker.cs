using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class WrongHardIpSheetNameChecker
    {
        private static readonly Regex _regex = new Regex("HARDIP_|DCTEST_|RTOS_", RegexOptions.IgnoreCase);

        public void CheckWrongHardIpSheetName(ExcelWorkbook planWorkbook)
        {
            foreach (ExcelWorksheet sheet in planWorkbook.Worksheets)
            {
                if (_regex.IsMatch(sheet.Name) ||
                    sheet.Name.StartsWith(NeededSheets.PrefixWireless, StringComparison.OrdinalIgnoreCase) ||
                    sheet.Name.StartsWith(NeededSheets.PrefixLcd, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                Dictionary<string, int> headerOrder = EpplusExtensions.GetHeaderOrder(sheet);
                int matchCount = 0;
                foreach (string header in HardIpPattern.KnownHeaders)
                {
                    int headerIndex = headerOrder.FirstOrDefault(a => Regex.IsMatch(a.Key, header, RegexOptions.IgnoreCase)).Value;
                    if (headerIndex > 0)
                    {
                        matchCount++;
                    }
                }
                if (matchCount * 2 > HardIpPattern.KnownHeaders.Count)
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_MissingHardIpSheet_01,
                        sheet.Name,
                        1,
                        0,
                        [sheet.Name]
                    );
                    Response.Report("Here is one HardIP sheet :\"" + sheet.Name + "\" but sheet name is not start with \"HardIP_\"", EnumMessageLevel.Error, 30);
                }
            }
        }
    }
}
