using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using EfuseCheckCmdLib.Base;
using EfuseCheckCmdLib.Datalog;
using EfuseCheckCmdLib.EFuse.EFuseApp;
using EfuseCheckCmdLib.Output;

using OfficeOpenXml;

using TestPlanLib.Efuse;

namespace EfuseCheckCmdLib
{
    public static partial class EfuseCmdUtility
    {
        public static readonly Regex RegBin = BinRegex();
        public static readonly Regex RegHex = HexRegex();
        public static readonly Regex RegDec = DecRegex();

        [GeneratedRegex("^(?<format>0*[bB])(?<value>[01]+)$")]
        private static partial Regex BinRegex();

        [GeneratedRegex("^(?<format>0*[xX])(?<value>[a-fA-F0-9]+)$")]
        private static partial Regex HexRegex();

        [GeneratedRegex(@"^(?<value>\d+(?:\.\d+)?)(?<unit>[numkKMgG]?)[a-zA-Z]*$")]
        private static partial Regex DecRegex();

        [GeneratedRegex(@"cfg_condition_\d+_\d+", RegexOptions.IgnoreCase)]
        private static partial Regex CfgConditionRegex();

        public static readonly Dictionary<string, int> JobStages = new()
        {
            { "CFGTABLE", 0},
            { "CP1_EARLY", 0},
            { "CP1", 0},
            { "CP2", 1},
            { "WLFT", 2},
            { "FT1", 3},
            { "FT2", 4},
            { "FT3", 5},
        };
        private static readonly List<(string, string, string, string, string, string, CheckStatus, bool)> _comparisons = [];

        public static void CompareDataWithDictionary(List<EfuseRow> efuseRows, EfuseBitDefTable efuseBitDefTable)
        {
            foreach (EfuseRow logData in efuseRows)
            {
                List<string>? targetBdfRow = efuseBitDefTable.Rows.FirstOrDefault(x => x[efuseBitDefTable.NameIdx].EqualsIgnoreCase(logData.SubConfig));
                if (targetBdfRow != null)
                {
                    bool isDefault = targetBdfRow[efuseBitDefTable.DefaultOrRealIdx].EqualsIgnoreCase("Default");
                    CheckStatus isMatch = CheckStatus.Match;
                    CheckStatus bankMatch = CheckStatus.Match;

                    Console.WriteLine($"\nMatch found for Bank_config: {logData.BankConfig}, Sub_Config: {logData.SubConfig}");

                    // First call includes BankConfig and SubBankConfig
                    CompareAndLog(logData.Site, "MSB BIT", logData.MsbBit.ToString(), targetBdfRow[efuseBitDefTable.MsbBitIdx], ref isMatch, logData.BankConfig, logData.SubConfig);
                    bankMatch = isMatch == CheckStatus.Mismatch || bankMatch == CheckStatus.Mismatch ? CheckStatus.Mismatch : bankMatch;

                    // Subsequent calls omit BankConfig and SubBankConfig (use default `null`)
                    CompareAndLog(logData.Site, "LSB BIT", logData.LsbBit.ToString(), targetBdfRow[efuseBitDefTable.LsbBitIdx], ref isMatch);
                    bankMatch = isMatch == CheckStatus.Mismatch || bankMatch == CheckStatus.Mismatch ? CheckStatus.Mismatch : bankMatch;
                    CompareAndLog(logData.Site, "Programming Stage", logData.ProgrammingStage, targetBdfRow[efuseBitDefTable.PrgStageIdx], ref isMatch);
                    bankMatch = isMatch == CheckStatus.Mismatch || bankMatch == CheckStatus.Mismatch ? CheckStatus.Mismatch : bankMatch;
                    CompareAndLog(logData.Site, "Bit Width", logData.BitWidth.ToString(), targetBdfRow[efuseBitDefTable.BitWidthIdx], ref isMatch);
                    bankMatch = isMatch == CheckStatus.Mismatch || bankMatch == CheckStatus.Mismatch ? CheckStatus.Mismatch : bankMatch;
                    CompareAndLog(logData.Site, "Default Value", logData.HexValue, targetBdfRow[efuseBitDefTable.DefaultValueIdx], ref isMatch, "", "", isDefault);
                    bankMatch = isMatch == CheckStatus.Mismatch || bankMatch == CheckStatus.Mismatch ? CheckStatus.Mismatch : bankMatch;
                    CompareAndLog(logData.Site, "DecValue", logData.DecValue, targetBdfRow[efuseBitDefTable.DefaultValueIdx], ref isMatch, "", "", isDefault);
                    bankMatch = isMatch == CheckStatus.Mismatch || bankMatch == CheckStatus.Mismatch ? CheckStatus.Mismatch : bankMatch;
                    CompareAndLog(logData.Site, "BinValue", logData.BinValue, targetBdfRow[efuseBitDefTable.DefaultValueIdx], ref isMatch, "", "", isDefault);
                    bankMatch = isMatch == CheckStatus.Mismatch || bankMatch == CheckStatus.Mismatch ? CheckStatus.Mismatch : bankMatch;

                    Console.WriteLine(bankMatch == CheckStatus.Match ? "All values match!\n" : "Mismatch in values.\n");
                }
                else
                {
                    Console.WriteLine($"\n❌ No match found for Bank_config: {logData.BankConfig}, Sub_Config: {logData.SubConfig}");
                }
            }
        }

        public static bool AreHexValuesEqual(string hex1, string hex2)
        {
            try
            {
                if (hex1.StartsWithIgnoreCase("0x"))
                {
                    hex1 = hex1[2..];
                }
                else if (hex1.StartsWithIgnoreCase("x"))
                {
                    hex1 = hex1[1..];
                }

                if (hex2.StartsWithIgnoreCase("0x"))
                {
                    hex2 = hex2[2..];
                }
                else if (hex2.StartsWithIgnoreCase("x"))
                {
                    hex2 = hex2[1..];
                }
                ulong value1 = Convert.ToUInt64(hex1, 16);
                ulong value2 = Convert.ToUInt64(hex2, 16);
                return value1 == value2;
            }
            catch
            {
                return false;
            }
        }

        public static string BinaryStringToHexString(string binary)
        {
            if (string.IsNullOrEmpty(binary))
            {
                return binary;
            }

            StringBuilder result = new StringBuilder((binary.Length / 8) + 1);
            int mod4Len = binary.Length % 8;
            if (mod4Len != 0)
            {
                binary = binary.PadLeft(((binary.Length / 8) + 1) * 8, '0');
            }

            for (int i = 0; i < binary.Length; i += 8)
            {
                string eightBits = binary.Substring(i, 8);
                result.AppendFormat("{0:X2}", Convert.ToByte(eightBits, 2));
            }

            return result.ToString();
        }

        public static string BinaryStringToDecimal(string binary)
        {
            if (string.IsNullOrWhiteSpace(binary))
            {
                return "0";
            }

            string result = "0";

            foreach (char bit in binary)
            {
                if (bit != '0' && bit != '1')
                {
                    throw new ArgumentException("Invalid binary digit in input.");
                }

                // Multiply current result by 2
                result = MultiplyDecimalByTwo(result);

                // If bit is 1, add 1
                if (bit == '1')
                {
                    result = AddDecimalOne(result);
                }
            }

            return result;
        }

        private static string MultiplyDecimalByTwo(string number)
        {
            StringBuilder sb = new StringBuilder();
            int carry = 0;

            for (int i = number.Length - 1; i >= 0; i--)
            {
                int digit = ((number[i] - '0') * 2) + carry;
                sb.Insert(0, digit % 10);
                carry = digit / 10;
            }

            if (carry > 0)
            {
                sb.Insert(0, carry);
            }

            return sb.ToString();
        }

        private static string AddDecimalOne(string number)
        {
            StringBuilder sb = new StringBuilder();
            int carry = 1;

            for (int i = number.Length - 1; i >= 0; i--)
            {
                int digit = number[i] - '0' + carry;
                sb.Insert(0, digit % 10);
                carry = digit / 10;
            }

            if (carry > 0)
            {
                sb.Insert(0, carry);
            }

            return sb.ToString();
        }

        private static void CompareAndLog(int site, string propertyName, string objectValue, string dictValue, ref CheckStatus checkStatus, string? bankConfig = null, string? subBankConfig = null, bool isDefault = true)
        {
            if (!string.IsNullOrEmpty(dictValue))
            {
                bool match = objectValue == dictValue;

                if (isDefault)
                {
                    if (match)
                    {
                        Console.WriteLine($"{propertyName}: {objectValue} (Matches Dictionary)");
                        checkStatus = CheckStatus.Match;
                    }
                    else
                    {
                        Console.WriteLine($"{propertyName} Mismatch - Datalog: {objectValue}, BDF: {dictValue}");
                        checkStatus = CheckStatus.Mismatch;
                    }
                }
                else
                {
                    Console.WriteLine($"{propertyName}: {objectValue} (Real value, no need to check)");
                    checkStatus = CheckStatus.NoNeedToCheck;
                }
            }
            else
            {
                Console.WriteLine($"  ❓ {propertyName} missing in dictionary");
                checkStatus = CheckStatus.Mismatch;
            }
            _comparisons.Add((bankConfig ?? "", subBankConfig ?? "", site.ToString(), propertyName, objectValue, dictValue, checkStatus, isDefault));
            // Write to Excel
            //WriteToExcel(comparisons);
        }

        public static void CheckCfgCondition(ExcelWorksheet excelWorksheet, DiceInfo diceInfo, EfuseBitDefTable efuseBitDefTable, EfuseCfgTable efuseCfgTable, int dieIndx)
        {
            int startrow = 3;
            excelWorksheet.Cells[startrow, dieIndx].Value = $"X{diceInfo.XCoor}_Y{diceInfo.YCoor}";
            startrow += 2;
            var allConfigItems = diceInfo.EfuseRows.Where(p => Regex.IsMatch(p.BankConfig, efuseCfgTable.TableName.Replace("_table", ""), RegexOptions.IgnoreCase)).ToList();
            string cfgXxx = GetCfgBin(allConfigItems);
            List<string> tableItem;
            foreach (List<string> item in efuseCfgTable.CfgRows)
            {
                if (item.Count < efuseCfgTable.PrgStageIdx)
                {
                    continue;
                }

                if (item[efuseCfgTable.ConditionIdx].Length == 0)
                {
                    break;
                }

                int msb = Convert.ToInt32(item[efuseCfgTable.MsbBitIdx]);
                int lsb = Convert.ToInt32(item[efuseCfgTable.LsbBitIdx]);
                if (msb > cfgXxx.Length - 1)
                {
                    break;
                }

                string value = GetCfg(cfgXxx, msb, lsb);
                string currentValue = GetCurrentValue(ref value, item, efuseCfgTable);
                int itemIndx = efuseCfgTable.CfgRows.IndexOf(item) + startrow;
                if (efuseCfgTable.Scenario.TryGetValue(diceInfo.Scenario, out int value1))
                {
                    string cFgValue = item[value1];
                    int targetJobStage = -1;
                    if (JobStages.TryGetValue(item[efuseCfgTable.PrgStageIdx], out int value2))
                    {
                        targetJobStage = value2;
                    }
                    else
                    {
                        targetJobStage = 999;
                    }

                    bool isProgrammed = targetJobStage <= diceInfo.JobNum;
                    {
                        if (!CompareCfgTableWithDatalog(ref currentValue, cFgValue, isProgrammed))
                        {
                            WriteExcel.HighLightCell(excelWorksheet, itemIndx, dieIndx, itemIndx, dieIndx, Color.Red);
                        }
                        else
                        {
                            WriteExcel.HighLightCell(excelWorksheet, itemIndx, dieIndx, itemIndx, dieIndx, Color.LightGreen);
                        }
                    }
                }
                else
                {
                    WriteExcel.HighLightCell(excelWorksheet, itemIndx, dieIndx, itemIndx, dieIndx, Color.Yellow);
                }
                excelWorksheet.Cells[itemIndx, dieIndx].Value = currentValue;
            }

            //Search for Others
            foreach (EfuseRow item in allConfigItems)
            {
                string id = item.SubConfig;
                string value = item.BinValue;

                tableItem = efuseCfgTable.CfgRows.FirstOrDefault(p => p[efuseCfgTable.ConditionIdx].EqualsIgnoreCase(id))!;

                if (tableItem != null)
                {
                    string currentValue = GetCurrentValue(ref value, tableItem, efuseCfgTable);
                    int itemIndx = efuseCfgTable.CfgRows.IndexOf(tableItem) + startrow;
                    if (efuseCfgTable.Scenario.TryGetValue(diceInfo.Scenario, out int value1))
                    {
                        string cFgValue = tableItem[value1];
                        int targetJobStage = -1;
                        if (JobStages.TryGetValue(tableItem[efuseCfgTable.PrgStageIdx], out int value2))
                        {
                            targetJobStage = value2;
                        }
                        else
                        {
                            targetJobStage = 999;
                        }

                        bool isProgrammed = targetJobStage <= diceInfo.JobNum;
                        {
                            if (!CompareCfgTableWithDatalog(ref currentValue, cFgValue, isProgrammed))
                            {
                                WriteExcel.HighLightCell(excelWorksheet, itemIndx, dieIndx, itemIndx, dieIndx, Color.Red);
                            }
                            else
                            {
                                WriteExcel.HighLightCell(excelWorksheet, itemIndx, dieIndx, itemIndx, dieIndx, Color.LightGreen);
                            }
                        }
                    }
                    else
                    {
                        WriteExcel.HighLightCell(excelWorksheet, itemIndx, dieIndx, itemIndx, dieIndx, Color.Yellow);
                    }
                    excelWorksheet.Cells[itemIndx, dieIndx].Value = currentValue;
                    //Put in the end to prevent value format difference..
                }
            }
        }

        private static bool CompareCfgTableWithDatalog(ref string realValue, string cFgValue, bool isCompareDefault)
        {
            string regPrefix = @"(?<prefix>(\d*[Xx]+)|(0*[Bb]+0*))";
            string hexPrefix = Regex.Match(cFgValue, regPrefix).Groups["prefix"].ToString();
            if (!isCompareDefault)
            {
                cFgValue = hexPrefix + "0";
            }

            if (!Regex.IsMatch(cFgValue, regPrefix, RegexOptions.IgnoreCase))
            //if not Hex format, it would be binary and convert to hex
            {
                hexPrefix = "";
                try
                {

                    cFgValue = hexPrefix + Convert.ToInt32(cFgValue, 2).ToString("X");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(string.Format(ex.ToString()));
                    cFgValue = hexPrefix + cFgValue;
                }
            }

            realValue = hexPrefix + Convert.ToInt32(realValue, 2).ToString("X");
            return Convert.ToInt32(realValue, 16) == Convert.ToInt32(cFgValue, 16);
        }

        public static string GetCurrentValue(ref string value, List<string> item, EfuseCfgTable efuseCfgTable)
        {
            int bitWidth = int.Parse(item[efuseCfgTable.BitWidthIdx]);
            string realValue;
            if (value.Length >= bitWidth)
            {
                realValue = value[^bitWidth..];
                value = value[..^bitWidth];
            }
            else
            {
                realValue = value;
                for (int i = 0; i <= bitWidth - realValue.Length; i++)
                {
                    realValue = "0" + realValue;
                }
                value = "";
            }
            return realValue;
        }

        private static string GetCfg(string cfxXx, int msb, int lsb)
        {
            int len = cfxXx.Length;
            string cfgStr = cfxXx.Substring(len - msb - 1, msb - lsb + 1);
            // BinStr2Hex(cfgStr);
            return cfgStr;
        }

        private static string GetCfgBin(List<EfuseRow> efuseRows)
        {
            var cfgConditionXx =
                efuseRows.Where(a => CfgConditionRegex().IsMatch(a.SubConfig)).Select(a => a.BinValue).ToList();
            cfgConditionXx.Reverse();
            string xStr = string.Join("", cfgConditionXx);
            return xStr;
        }

        public static bool CheckEfuseDefaultValue(string val, string defValue, bool isNeedReverse = false)
        {
            if (defValue.EqualsIgnoreCase("NA") || defValue.EqualsIgnoreCase("N/A"))
            {
                defValue = "0";
            }

            return ConvertValue(val) == ConvertValue(defValue);
        }

        public static double ConvertValue(string valueStr)
        {
            double retVal = -1.0;
            try
            {
                string format = "";
                string tmpvalue = "";
                valueStr = valueStr.Replace("_", "");
                if (RegBin.IsMatch(valueStr))
                {
                    format = RegBin.Match(valueStr).Groups["format"].ToString();
                    tmpvalue = valueStr.Replace(format, "");
                    if (tmpvalue.Length > 64)
                    {
                        valueStr = tmpvalue;
                    }
                    else
                    {
                        valueStr = Convert.ToString(Convert.ToInt64(tmpvalue, 2));
                    }
                }
                else if (RegHex.IsMatch(valueStr))
                {
                    format = RegHex.Match(valueStr).Groups["format"].ToString();
                    tmpvalue = valueStr.Replace(format, "");
                    if (tmpvalue.Length > 16)
                    {
                        valueStr = tmpvalue;
                    }
                    else
                    {
                        valueStr = Convert.ToString(Convert.ToInt64(tmpvalue, 16));
                    }
                }

                if (RegDec.IsMatch(valueStr))
                {
                    string value = RegDec.Match(valueStr).Groups["value"].Value;
                    string unit = RegDec.Match(valueStr).Groups["unit"].Value;
                    if (!string.IsNullOrEmpty(value))
                    {
                        retVal = double.Parse(value) * ConvertUnitStr2Double(unit);
                    }
                    else
                    {
                        retVal = -999;
                    }
                }
                else
                {
                    retVal = -999;
                }

                /* inputStr Type:
             * 1. N/A
             * 2. Functional P/F
             * 3. BinCutLow
             * 4. numberic value
             */
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format(ex.ToString()));
            }
            return retVal;
        }

        public static double ToDec(string text)
        {
            try
            {
                text = text.ToUpper().Replace("_", "");
                if (string.IsNullOrEmpty(text))
                {
                    throw new ArgumentException("Hex value cannot be null or empty", nameof(text));
                }
                if (text.StartsWith('X'))
                {
                    text = text.TrimStart('X');
                    return Convert.ToInt64(text, 16);
                }
                if (text.StartsWith('B'))
                {
                    text = text.TrimStart('B');
                    if (int.TryParse(text, out int res))
                    {
                        return Convert.ToInt64(text, 2);
                    }
                }
                return double.TryParse(text, out double result) ? result : -999;
            }
            catch
            {
                throw new Exception();
            }
        }

        private static double ConvertUnitStr2Double(string unitStr)
        {
            double retValue = 1;
            if (!string.IsNullOrEmpty(unitStr))
            {
                //Mini unit parser
                switch (unitStr[0])
                {
                    case 'G':
                        retValue *= 1.0e9;
                        break;
                    case 'M':
                        retValue *= 1.0e6;
                        break;
                    case 'K':
                        retValue *= 1.0e3;
                        break;
                    case 'm':
                        //logVal *= 1.0e-3;  //<- for all bincut voltage use mV as base
                        break;
                    case 'u':
                        retValue *= 1.0e-6;
                        break;
                    case 'n':
                        retValue *= 1.0e-9;
                        break;
                }
            }
            return retValue;
        }
    }
}
