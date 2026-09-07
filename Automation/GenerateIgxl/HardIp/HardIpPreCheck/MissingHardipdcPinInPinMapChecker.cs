using System;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlSheets;

using LogLib.Static;

using TestPlanLib.HardIpDc.BaseData;
using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class MissingHardIpDcPinInPinMapChecker
    {
        //private const string ErrorMessage = "Missing HardIpDc pin in Pinmap";
        private static readonly Regex _powerPinRegex = new Regex(@"(?<pin>VDD\w+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public void CheckMissingPinInPinMap(HardIpDcSheet dcReader)
        {
            foreach (HardIpCategoryDef category in dcReader.Rows)
            {
                foreach (HardIpDcRow hardIpDcRow in category.DataRows)
                {
                    if (!TestProgram.IgxlWorkBk.PinMapPair.Value.PinList.Any(x => x.PinName.Equals(hardIpDcRow.PinName, StringComparison.OrdinalIgnoreCase))
                        && !TestProgram.IgxlWorkBk.PinMapPair.Value.GroupList.Any(y => y.PinName.Equals(hardIpDcRow.PinName, StringComparison.OrdinalIgnoreCase)))
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_MissingHardIpDcPin_01,
                            NeededSheets.HardIpDc,
                            Convert.ToInt32(hardIpDcRow.RowNum),
                            0,
                            []
                        );
                        Response.Report($"The HardIpDc pin: {hardIpDcRow.PinName} can not find in Pin map!", EnumMessageLevel.Error, 0);
                    }

                    string[] targetValues = { hardIpDcRow.Vil, hardIpDcRow.Vih, hardIpDcRow.Vol, hardIpDcRow.Voh, hardIpDcRow.Iol, hardIpDcRow.Ioh, hardIpDcRow.Vcl, hardIpDcRow.Vch, hardIpDcRow.Vt, hardIpDcRow.Vicm, hardIpDcRow.Vid, hardIpDcRow.Vod };

                    foreach (string value in targetValues)
                    {
                        if (string.IsNullOrEmpty(value))
                        {
                            continue;
                        }

                        Match match = _powerPinRegex.Match(value);
                        if (match.Success)
                        {
                            string grabPin = match.Groups["pin"].Value;
                            PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
                            bool pinExists = pinMap.PinList.Any(x => x.PinName.Equals(grabPin, StringComparison.OrdinalIgnoreCase)) || pinMap.GroupList.Any(y => y.PinName.Equals(grabPin, StringComparison.OrdinalIgnoreCase));

                            if (!pinExists)
                            {
                                int rowNum = int.TryParse(hardIpDcRow.RowNum, out int res) ? res : 0;
                                ErrorReportManager.AddError(
                                    HardIpErrorType.E_MissingHardIpDcPin_01,
                                    NeededSheets.HardIpDc,
                                    rowNum,
                                    0,
                                    []
                                );
                                Response.Report($"The HardIpDc pin: {grabPin} can not find in Pin map!", EnumMessageLevel.Error, 0);
                            }
                        }
                    }
                }
            }
        }
    }
}
