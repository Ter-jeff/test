using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;

using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader
{
    public class PinListReader
    {
        public static void Read(ExcelWorksheet sheet)
        {
            for (int i = 1; i <= sheet.Dimension.End.Column; i++)
            {
                if (sheet.Cells[1, i].Value != null)
                {
                    string groupName = sheet.Cells[1, i].Value.ToString().ToUpper(); // pin group name is defined in the first cell

                    var pinList = new List<string>();
                    int row = 2;

                    while (sheet.Cells[row, i].Value != null)
                    {
                        string pinName = sheet.Cells[row, i].Value.ToString().ToUpper();

                        // check if pin name exists in the PinList from PinMap
                        if (!UtilityMain.UtilityData.PinList.Keys.ToList().Exists(a => a.Equals(pinName.Replace("_", ""), StringComparison.OrdinalIgnoreCase)))
                        {
                            const string errMessage = "Missing (pin name)/(pin group) in PinMap.txt/pinList sheet";
                            ErrorManager.AddError(ErrorType.MissingPinName, sheet.Name, row, i, "Use", errMessage, pinName);
                            ErrorReportManager.AddError(CharErrorType.E_MissingPinName_01, sheet.Name, row, i, [],
                                new ErrorInfo() { Comments = new List<string>() { pinName } });
                        }

                        pinList.Add(pinName);
                        row++;
                    }
                    UtilityMain.UtilityData.PinGroupList.Add(groupName, pinList);
                }
                else
                {
                    return;
                }
            }
        }

        public static void PinListCheck()
        {
            var pinAddLines = new List<string>();

            foreach (string pingroup in UtilityMain.UtilityData.PinGroupList.Keys)
            {
                if (!UtilityMain.UtilityData.PinGroups.ContainsKey(pingroup))
                {
                    string pinType;
                    bool firstGroup = true;
                    if (Regex.IsMatch(pingroup, "^VDD.*"))
                    {
                        pinType = "Power";
                    }
                    else if (Regex.IsMatch(pingroup, "^DIFF.*"))
                    {
                        pinType = "DIFFERENTIAL";
                    }
                    else
                    {
                        pinType = "I/O";
                    }

                    foreach (string pin in UtilityMain.UtilityData.PinGroupList[pingroup])
                    {
                        if (firstGroup)
                        {
                            pinAddLines.Add("\t" + pingroup + "\t" + pin + "\t" + pinType);
                        }
                        else
                        {
                            pinAddLines.Add("\t" + pingroup + "\t" + pin);
                        }

                        firstGroup = false;
                    }
                }
                else if (UtilityMain.UtilityData.PinGroups[pingroup].Count != UtilityMain.UtilityData.PinGroupList[pingroup].Count ||
                    !UtilityMain.UtilityData.PinGroups[pingroup].All(a => UtilityMain.UtilityData.PinGroupList[pingroup].Any(b => b == a)))
                {
                    const string outString = "Pin group in PinList sheet not match pinmap file";
                    ErrorManager.AddError(ErrorType.PinGroupNotMatch, "PinList", 1, 1, "Use", outString);
                    ErrorReportManager.AddError(CharErrorType.E_PinGroupNotMatch_01, "PinList", 1, 1, []);
                }
            }

            if (!pinAddLines.Any())
            {
                return;
            }

            // todo: move pinmap_add.txt to be a sheet in the IM file
            string pinmapAddFile = Path.Combine(UtilityMain.UtilityData.InputParam.TarDic, "pinmap_add.txt");
            var pinmapWriter = new StreamWriter(pinmapAddFile);
            foreach (string line in pinAddLines)
            {
                pinmapWriter.WriteLine(line);
            }

            pinmapWriter.Close();
        }
    }
}
