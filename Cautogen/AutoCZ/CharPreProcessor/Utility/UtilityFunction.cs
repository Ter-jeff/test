using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Static;

namespace Cautogen.AutoCZ.CharPreProcessor.Utility
{
    public class UtilityFunction
    {
        private static readonly Dictionary<string, double> _scaleDict = new Dictionary<string, double>
            {
                {"n", 1e-9},
                {"u", 1e-6},
                {"m", 1e-3},
                {"k", 1e3},
                {"M", 1e6},
                {"G", 1e9}
            };
        private static readonly List<string> _smallUnits = new List<string> { "V", "A", "S" };
        private static readonly List<string> _bigUnits = new List<string> { "OHM", "HZ" };

        public static string GetTimeStamp()
        {
            return "_" + TimeContext.Now.ToString("yyyyMMddHHmmss");
        }

        public string GetPinType(string pinName)
        {
            string name = pinName.ToUpper();

            if (name.Contains("::"))
            {
                name = pinName.Split(':')[0];
            }

            return Regex.IsMatch(name, "VDD")
                ? "power"
                : "I/O";
        }

        public bool PinExistInPinMap(string pinName)
        {
            var pinList = new List<string>();

            if (IsCmPins(pinName))  // e.g. pinCMpin2 in USERDEF4
            {
                string[] array = Regex.Split(pinName, "CM", RegexOptions.IgnoreCase);
                pinList.AddRange(array);
            }
            else if (Regex.IsMatch(pinName, "Diff$", RegexOptions.IgnoreCase))//PCIETX0VDIFF------Diff Pins
            {
                int length = pinName.Length;
                string pinName1 = pinName.Substring(0, ((length - 5) / 2) - 1);
                string pinName2 = pinName.Substring((length - 5) / 2, ((length - 5) / 2) - 1);
                if (pinName1 == pinName2)
                {
                    pinList.Add(pinName1 + "P");
                    pinList.Add(pinName1 + "N");

                }
                else
                {
                    pinList.Add(pinName.Substring(0, length - 5) + "P");
                    pinList.Add(pinName.Substring(0, length - 5) + "N");
                }
            }
            else if (Regex.IsMatch(pinName, "[a-zA-Z]+DIFF[a-zA-Z]+", RegexOptions.IgnoreCase)) // PCIETX0PDIFFPCIETX0N---- Diff Pins
            {
                string[] array = Regex.Split(pinName, "DIFF", RegexOptions.IgnoreCase);
                pinList.AddRange(array);
            }
            else
            {
                pinList.Add(pinName);
            }

            return pinList.All(pin =>
                UtilityMain.UtilityData.PinList.Keys.ToList().Exists(a => a.Equals(pin, StringComparison.OrdinalIgnoreCase)) ||
                UtilityMain.UtilityData.PinGroupList.Keys.ToList().Exists(b => b.Equals(pinName, StringComparison.OrdinalIgnoreCase)));
        }

        public Characterization GetCharItem(string patternName)
        {
            var allPlanItems = new List<Characterization>();
            foreach (List<Characterization> items in CharPlan.CharPlanSheetDict.Values)
            {
                allPlanItems.AddRange(items);
            }

            Characterization resultItem = allPlanItems.FirstOrDefault(
                a => a.Payload1.Equals(patternName, StringComparison.OrdinalIgnoreCase));
            if (resultItem != null)
            {
                return resultItem;
            }

            return new Characterization();
        }

        public string RemoveSuffix(string str)
        {
            return Regex.Match(str, "([^.]+).*").Groups[1].Value;
        }


        /* Convert the Units of USL/LSL */
        public string ConvertUsllsl(string str, string units, string sheetName, int rowNum)
        {
            if (str == "")
            {
                return "";
            }

            if (double.TryParse(str, out double value))
            {
                if (Regex.IsMatch(units, "^m.*"))
                {
                    return (value / 1000).ToString(CultureInfo.InvariantCulture);
                }

                if (Regex.IsMatch(units, "^u.*"))
                {
                    return (value / 1000000).ToString(CultureInfo.InvariantCulture);
                }

                if (Regex.IsMatch(units, "^n.*"))
                {
                    return (value / 1000000000).ToString(CultureInfo.InvariantCulture);
                }

                if (Regex.IsMatch(units, "^K.*"))
                {
                    return (value * 1000).ToString(CultureInfo.InvariantCulture);
                }

                if (Regex.IsMatch(units, "^M.*"))
                {
                    return (value * 1000000).ToString(CultureInfo.InvariantCulture);
                }

                return str;
            }
            ErrorManager.AddError(ErrorType.WrongLimitFormat, sheetName, rowNum, 1, "Use", "Wrong format of USL/LSL", str);
            ErrorReportManager.AddError(CharErrorType.E_WrongLimitFormat_01, sheetName, rowNum, 1, [],
                new ErrorInfo() { Comments = new List<string>() { str } });
            return "";
        }

        /// <summary>
        /// Remove date code of init pattern and pattern payload
        /// </summary>
        /// <param name="patternName"></param>
        /// <returns></returns>
        public static string RemoveDateCode(string patternName)
        {
            if (patternName == "")
            {
                return "";
            }

            if (Regex.IsMatch(patternName, ".*_(\\d{10})_.*"))
            {
                string dateCode = Regex.Match(patternName, ".*(_\\d{10})_.*").Groups[1].Value;
                patternName = Regex.Replace(patternName, dateCode, "");
            }

            List<string> strList = patternName.Split('_').ToList();
            if (strList.Count < 2)
            {
                return patternName;
            }

            if (!UtilityMain.UtilityData.DeviceMapping.ContainsKey(strList[1]))
            {
                return patternName;
            }

            strList[1] = UtilityMain.UtilityData.DeviceMapping[strList[1]];
            string newPattern = strList.Aggregate("", (current, str) => current + str + "_");
            return newPattern.Trim('_');
        }

        public void MoveInitPatternToTheFront(Characterization item)
        {
            string[] initArry = { "", "", "", "", "", "", "", "", "", "" };
            int count = 0;
            if (item.Init1 != "")
            {
                initArry[count] = item.Init1;
                count++;
            }
            if (item.Init2 != "")
            {
                initArry[count] = item.Init2;
                count++;
            }
            if (item.Init3 != "")
            {
                initArry[count] = item.Init3;
                count++;
            }
            if (item.Init4 != "")
            {
                initArry[count] = item.Init4;
                count++;
            }
            if (item.Init5 != "")
            {
                initArry[count] = item.Init5;
                count++;
            }
            if (item.Init6 != "")
            {
                initArry[count] = item.Init6;
                count++;
            }
            if (item.Init7 != "")
            {
                initArry[count] = item.Init7;
                count++;
            }
            if (item.Init8 != "")
            {
                initArry[count] = item.Init8;
                count++;
            }
            if (item.Init9 != "")
            {
                initArry[count] = item.Init9;
                count++;
            }
            if (item.Init10 != "")
            {
                initArry[count] = item.Init10;
            }

            item.Init1 = initArry[0];
            item.Init2 = initArry[1];
            item.Init3 = initArry[2];
            item.Init4 = initArry[3];
            item.Init5 = initArry[4];
            item.Init6 = initArry[5];
            item.Init7 = initArry[6];
            item.Init8 = initArry[7];
            item.Init9 = initArry[8];
            item.Init10 = initArry[9];
        }

        /// <summary>
        /// Deal with Voltage column
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public string ConvertVoltage(string str)
        {
            if (str == "")
            {
                return "";
            }

            if (str == "L")
            {
                return "LV";
            }

            if (str == "H")
            {
                return "HV";
            }

            if (str == "N" || str == "N1" || str == "N2")
            {
                return "NV";
            }

            return str + "_" + "NV";
        }

        /// <summary>
        /// Check power condition（Also Shmoo condition --new request ） of functional header tests to judge skip or combine
        /// </summary>
        /// <param name="tempItem"></param>
        /// <param name="myItem"></param>
        /// <returns></returns>
        public bool CheckPowerCondition(Characterization tempItem, Characterization myItem)
        {
            if (tempItem.PowerSupplyX.Count != myItem.PowerSupplyX.Count ||
                tempItem.PowerSupplyY.Count != myItem.PowerSupplyY.Count ||
                tempItem.PowerSupplyZ.Count != myItem.PowerSupplyZ.Count)
            {
                return false;
            }

            int shmooIndex = 0;
            foreach (ShmooSpec shmoo in tempItem.PowerSupplyX)
            {
                //Remove "shmoo.Start == shmoo.Stop &&" for new request
                if (shmoo.Stop != myItem.PowerSupplyX[shmooIndex].Stop || shmoo.Start != myItem.PowerSupplyX[shmooIndex].Start)
                {
                    return false;
                }

                shmooIndex++;
            }
            shmooIndex = 0;
            foreach (ShmooSpec shmoo in tempItem.PowerSupplyY)
            {
                //Remove "shmoo.Start == shmoo.Stop &&" for new request
                if (shmoo.Stop != myItem.PowerSupplyY[shmooIndex].Stop || shmoo.Start != myItem.PowerSupplyY[shmooIndex].Start)
                {
                    return false;
                }

                shmooIndex++;
            }
            shmooIndex = 0;
            foreach (ShmooSpec shmoo in tempItem.PowerSupplyZ)
            {
                //Remove "shmoo.Start == shmoo.Stop &&" for new request
                if (shmoo.Stop != myItem.PowerSupplyZ[shmooIndex].Stop || shmoo.Start != myItem.PowerSupplyZ[shmooIndex].Start)
                {
                    return false;
                }

                shmooIndex++;
            }
            int pinIndex = 0;
            if (tempItem.Pins.Count != myItem.Pins.Count)
            {
                return false;
            }

            foreach (Pin pin in tempItem.Pins)
            {
                if (pin.Start != myItem.Pins[pinIndex].Start || pin.Stop != myItem.Pins[pinIndex].Stop || pin.PinType != myItem.Pins[pinIndex].PinType)
                {
                    return false;
                }

                pinIndex++;

            }
            return true;
        }

        public void SubInstanceNames(string allInstanceNames, List<string> extraNames)
        {
            List<string> instanceList = allInstanceNames.Split(',').ToList();
            int length = 0;
            var tempList = new List<string>();
            foreach (string instance in instanceList)
            {
                length += instance.Length + 1;
                if (length < 32766)
                {
                    tempList.Add(instance);
                }
                else
                {
                    extraNames.Add(string.Join(",", tempList));
                    length = instance.Length;
                    tempList = new List<string> { instance };
                }
            }
            extraNames.Add(string.Join(",", tempList));
        }

        /* Check if pinName specified in USERDEF4 is in the format of pinACMpinB */
        public static bool IsCmPins(string pinName)
        {
            if (!pinName.Contains("CM", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return pinName.Substring(0, (pinName.Length / 2) - 2) == pinName.Substring((pinName.Length / 2) + 1, (pinName.Length / 2) - 2);
        }

        /* for "pattern:3" or "pattern:DigSrc" return "pattern" */
        public static string TrimPatternName(string patternName)
        {
            return string.IsNullOrEmpty(patternName)
                ? ""
                : string.Join(",", patternName.Split(',').Select(x => x.Split(':')[0]));
        }

        public static string NormalizeUnit(string rawValue)
        {
            List<string> rawValueList = rawValue.Split(',').ToList();
            var resultList = new List<string>();

            const string regScale = @"(?<value>\-*\d+(\.\d+)*(E\-*\d+)*)\s*(?<scale>[numkMG]+)*(?<unit>[VAs(Hz)(Ohm)]+)*";
            foreach (string rawValueItem in rawValueList)
            {
                var result = new List<string>();
                foreach (string subRawValueItem in rawValueItem.Split(':'))
                {
                    string subResult = subRawValueItem;
                    if (subRawValueItem.StartsWith("INIT", StringComparison.OrdinalIgnoreCase) || subRawValueItem.StartsWith("PL", StringComparison.OrdinalIgnoreCase))
                    {
                        result.Add(subRawValueItem);
                        continue;
                    }

                    if (Regex.IsMatch(subRawValueItem, regScale, RegexOptions.IgnoreCase))
                    {
                        string value = Regex.Match(subRawValueItem, regScale, RegexOptions.IgnoreCase).Groups["value"].Value;
                        string scaleStr = Regex.Match(subRawValueItem, regScale, RegexOptions.IgnoreCase).Groups["scale"].Value;
                        string unit = Regex.Match(subRawValueItem, regScale, RegexOptions.IgnoreCase).Groups["unit"].Value;

                        // force casting scale string according to unit
                        if (!string.IsNullOrEmpty(unit))
                        {
                            unit = unit.ToUpper();
                            if (scaleStr.ToUpper() == "M")
                            {
                                if (_smallUnits.Contains(unit))
                                {
                                    scaleStr = "m";
                                }

                                if (_bigUnits.Contains(unit))
                                {
                                    scaleStr = "M";
                                }
                            }
                        }

                        // lookup scale from scaleDict
                        double scale = 1.0;
                        if (_scaleDict.TryGetValue(scaleStr, out double value1))
                        {
                            scale = value1;
                        }

                        subResult = (double.Parse(value) * scale).ToString("G15", CultureInfo.InvariantCulture);
                    }
                    result.Add(subResult);
                }
                resultList.Add(string.Join(":", result));
            }
            return string.Join(",", resultList);
        }

        public static string ConvertUnits(string limitStr)
        {
            if (limitStr.Contains("10^"))
            {
                limitStr = limitStr.Replace("*10^", "E");
            }

            if (limitStr == "" || limitStr.Contains("E") || Regex.IsMatch(limitStr, @"^(\d|\.|-)+$"))//Limit value may be 1.2E-5
            {
                return limitStr;
            }

            if (Regex.IsMatch(limitStr, @"^(\d|\.|-|\*)+(\w)*$"))
            {
                string limitNum = Regex.Match(limitStr, @"(?<num>((\d|\.|-)+))[^\*]*").Groups["num"].ToString();
                if (limitNum == "0")
                {
                    return limitNum;
                }

                string limitUnit = limitStr.Replace(limitNum, "").Trim();
                double rate = 1;
                if (Regex.IsMatch(limitUnit, @"^[\*]?m.*"))
                {
                    rate = 1 / (double)1000;
                }
                else if (Regex.IsMatch(limitUnit, @"^[\*]?u.*"))
                {
                    rate = 1 / (double)1000000;
                }
                else if (Regex.IsMatch(limitUnit, @"^[\*]?n.*"))
                {
                    rate = 1 / (double)1000000000;
                }
                else if (Regex.IsMatch(limitUnit.ToLower(), @"^[\*]?k.*"))
                {
                    rate = 1000;
                }
                else if (Regex.IsMatch(limitUnit, @"^[\*]?M.*"))
                {
                    rate = 1000000;
                }
                else if (Regex.IsMatch(limitUnit, @"^[\*]?G.*"))
                {
                    rate = 1000000000;
                }

                if (double.TryParse(limitNum, out double value))
                {
                    return (value * rate).ToString("G15", CultureInfo.InvariantCulture);
                }
            }
            return limitStr;
        }

        public static string ConvertShiftSpeed(string shiftspeed)
        {
            if (!string.IsNullOrEmpty(shiftspeed) && double.TryParse(shiftspeed, out double speed))
            {
                if (speed > 10000)
                {
                    return (speed / 1000000).ToString();
                }
            }
            return shiftspeed;

        }
    }
}
