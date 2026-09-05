using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Automation.Static;

using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using EfuseCheckCmdLib.AlgorithmCheck;
using EfuseCheckCmdLib.Static;
using EfuseCheckCmdLib.Utility;

using LogLib.Utility;

using OfficeOpenXml;

using TestPlanLib.Efuse;

using Color = System.Drawing.Color;

namespace EfuseCheckCmdLib.EFuse.EFuseApp
{

    /// <summary>
    /// This class used to implement all common functions of Efuse Check Script
    /// </summary>
    public partial class EfuseScriptUtility
    {
        #region Static Property
        public const double ValueOutOfRange = -999;
        #endregion
        public const string Regbin = "(?<format>^0*[bB])[01]+";
        public const string Reghex = "(?<format>^0*[xX])[a-fA-F0-9]+";
        public const string RegDec = @"(?<value>\d+(\.\d+)*)(?<unit>[numkKMgG]*)\w*";

        [GeneratedRegex(Regbin, RegexOptions.IgnoreCase)]
        private static partial Regex BinRegex();

        [GeneratedRegex(Reghex, RegexOptions.IgnoreCase)]
        private static partial Regex HexRegex();

        [GeneratedRegex(RegDec, RegexOptions.IgnoreCase)]
        private static partial Regex DecRegex();

        [GeneratedRegex(@"(?<bank>\w*)\[")]
        private static partial Regex BankBracketRegex();

        [GeneratedRegex(@"\[(?<fieldName>.*)\]", RegexOptions.IgnoreCase)]
        private static partial Regex FieldNameBracketRegex();

        [GeneratedRegex(@"cfg_condition_\d+_\d+", RegexOptions.IgnoreCase)]
        private static partial Regex CfgConditionRegex();

        [GeneratedRegex("config", RegexOptions.IgnoreCase)]
        private static partial Regex ConfigRegex();

        [GeneratedRegex("SVM", RegexOptions.IgnoreCase)]
        private static partial Regex SvmRegex();

        [GeneratedRegex("BASE_VOLTAGE", RegexOptions.IgnoreCase)]
        private static partial Regex BaseVoltageRegex();

        [GeneratedRegex(@"\d\.\d")]
        private static partial Regex DecimalPointRegex();

        [GeneratedRegex("[A-Z]", RegexOptions.IgnoreCase)]
        private static partial Regex UppercaseLetterRegex();

        public static int BaseVoltage { get; set; } = 0;
        public static double BaseVoltageResolution { get; set; } = 0;
        public static readonly List<List<int>> SecKeyList =
                [
                    [0, 1, 2, 3, 4, 7, 8, 9, 12],
                    [0, 1, 3, 5, 7, 8, 10, 13],
                    [0, 2, 4, 6, 7, 9, 11, 14],
                    [1, 2, 5, 6, 7, 10, 11, 15],
                    [3, 4, 5, 6, 7, 12, 13, 14, 15],
                    [8, 9, 10, 11, 12, 13, 14, 15],
                ];
        public static readonly List<List<int>> DecKeyList =
                [
                    [0, 3, 4, 5, 6, 7, 8, 10, 11, 15, 16, 20, 22, 23, 26, 27, 28, 29],
                    [1, 4, 5, 6, 7, 8, 9, 11, 12, 16, 17, 21, 23, 24, 27, 28, 29, 30],
                    [2, 5, 6, 7, 8, 9, 10, 12, 13, 17, 18, 22, 24, 25, 28, 29, 30, 31],
                    [4, 5, 9, 13, 14, 15, 16, 18, 19, 20, 22, 25, 27, 28, 30, 31],
                    [0, 3, 4, 7, 8, 11, 14, 17, 19, 21, 22, 27, 31],
                    [0, 1, 3, 6, 7, 9, 10, 11, 12, 16, 18, 26, 27, 29],
                    [1, 2, 4, 7, 8, 10, 11, 12, 13, 17, 19, 27, 28, 30],
                    [2, 3, 5, 8, 9, 11, 12, 13, 14, 18, 20, 28, 29, 31],
                    [0, 5, 7, 8, 9, 11, 12, 13, 14, 16, 19, 20, 21, 22, 23, 26, 27, 28, 30],
                    [0, 1, 6, 8, 9, 10, 12, 13, 14, 15, 17, 20, 21, 22, 23, 24, 27, 28, 29, 31],
                    [1, 2, 3, 4, 5, 6, 8, 9, 13, 14, 18, 20, 21, 24, 25, 26, 27, 30],
                    [2, 3, 4, 5, 6, 7, 9, 10, 14, 15, 19, 21, 22, 25, 26, 27, 28, 31]
                ];

        //AccessJobOrder
        public static int GetJobOrder(string inJob)
        {
            if (inJob.Contains('(') && inJob.Contains(')'))
            {
                if (EfuseAlgorithmCheck.JobFlow.TryGetValue(inJob, out int order))
                {
                    return order;
                }
            }

            foreach (KeyValuePair<string, int> job in EfuseAlgorithmCheck.JobFlow)
            {
                if (job.Key.Contains('(') && job.Key.Contains(')'))
                {
                    string dbStage = job.Key.Split(['(', ')'], StringSplitOptions.RemoveEmptyEntries).Last();
                    string[] stageList = dbStage.ToUpper().Split([','], StringSplitOptions.RemoveEmptyEntries);
                    if (stageList.Contains(inJob.ToUpper()))
                    {
                        return job.Value;
                    }
                }
            }

            inJob = inJob.Split('_')[0];
            if (inJob.EqualsIgnoreCase("SLT"))
            {
                inJob = "FT3";
            }
            if (EfuseAlgorithmCheck.JobFlow.TryGetValue(inJob, out int jobOrder))
            {
                return jobOrder;
            }

            return 99;
        }

        public static void CheckEquation(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, int row, int col)
        {
            string hipValue = excelWorksheet.Cells[row, col + 1].Value.ToString()!;
            string efuseValue = excelWorksheet.Cells[row, col].Value.ToString()!;
            if (!string.IsNullOrEmpty(hipValue.Trim()) && !string.IsNullOrEmpty(efuseValue.Trim()))
            {
                string equation = excelWorksheet.Cells[row, col + 2].Value.ToString()!;
                equation = equation.Replace("EFUSE_VAL", efuseValue).Replace("HIP_VAL", hipValue);
                excelWorksheet.Cells[row, col + 2].Formula = equation;
                try
                {
                    excelWorksheet.Cells[row, col + 2].Calculate();
                    if (!excelWorksheet.Cells[row, col + 2].Value.ToString()!.EqualsIgnoreCase("true"))
                    {
                        Writer.HighLightCell(excelWorksheet, row, col + 2, row, col + 2, Color.Red);
                        var error = new Error(EfuseCheckCmdLibError.E_MismatchFormat_01, EnumErrorLevel.Error, excelWorksheet.Name, row, col + 2, EfuseCheckCmdLibError.E_MismatchFormat_01.MessageTemplate);
                        EFuseAppExcelWriter.AddComment(appendRichText, excelWorksheet, error);
                    }

                }
                catch (Exception ex)
                {
                    EfuseStatic.Result = EfuseCheckResultType.Exception;
                    if (EfuseStatic.IsCmd)
                    {
                        appendRichText(string.Format(ex.ToString()), "Red");
                    }
                    else
                    {
                        ErrorMessageBox.Show(string.Format(ex.ToString()));
                    }
                }
            }
            if (excelWorksheet.Cells[row, col].Value.ToString()!.EqualsIgnoreCase("false"))
            {

                Writer.HighLightCell(excelWorksheet, row, col, row, col, Color.Red);
                var error = new Error(EfuseCheckCmdLibError.E_MismatchFormat_02, EnumErrorLevel.Error, excelWorksheet.Name, row, col + 2, EfuseCheckCmdLibError.E_MismatchFormat_02.MessageTemplate);
                EFuseAppExcelWriter.AddComment(appendRichText, excelWorksheet, error);
            }
        }

        #region Check Functions
        //Check efuse rule : CRCCheck
        public static void CheckCrc(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, CrcItem crcItem, XDiceInfo xDiceInfo, EfuseBitDefTable efuseBitDefTable, int dieIndx)
        {
            var blockItems = new List<List<string>>();
            var otherBankItems = new List<List<string>>();
            string calcBank = "";
            if (crcItem.Description.ContainsIgnoreCase("BANK:"))
            {
                //calc other bank crc item
                calcBank = BankBracketRegex().Match(crcItem.Description).Groups["bank"].ToString();

                otherBankItems = efuseBitDefTable.Rows.FindAll(p => p[efuseBitDefTable.BlockIdx] == calcBank);
            }
            blockItems = efuseBitDefTable.Rows.FindAll(p => p[efuseBitDefTable.BlockIdx] == crcItem.BlockName);
            int offset = string.IsNullOrEmpty(calcBank) ? int.Parse(blockItems.First()[efuseBitDefTable.LsbBitIdx]) < int.Parse(blockItems.First()[efuseBitDefTable.MsbBitIdx]) ? int.Parse(blockItems.First()[efuseBitDefTable.LsbBitIdx]) : int.Parse(blockItems.First()[efuseBitDefTable.MsbBitIdx]) : int.Parse(otherBankItems.First()[efuseBitDefTable.LsbBitIdx]) < int.Parse(otherBankItems.First()[efuseBitDefTable.MsbBitIdx]) ? int.Parse(otherBankItems.First()[efuseBitDefTable.LsbBitIdx]) : int.Parse(otherBankItems.First()[efuseBitDefTable.MsbBitIdx]);
            List<string> keyList = string.IsNullOrEmpty(calcBank) ? [.. blockItems.Select(p => p[efuseBitDefTable.NameIdx])] : [.. otherBankItems.Select(p => p[efuseBitDefTable.NameIdx])];
            var bitStream = new List<string>();
            var missingitems = new List<string>();
            var wrongBitItems = new Dictionary<string, string>();
            foreach (string key in keyList)
            {
                EfuseDatalogItem? datalogItem =
                    xDiceInfo.AllReadFromDssc.FirstOrDefault(p => p.Block.EqualsIgnoreCase(string.IsNullOrEmpty(calcBank) ? crcItem.BlockName : calcBank) &&
                        p.Id.EqualsIgnoreCase(key));
                if (datalogItem != null)
                {
                    if (!string.IsNullOrEmpty(datalogItem.RawData))
                    {
                        bitStream.Insert(0, datalogItem.RawData);
                    }

                    List<string> bdfItem =
                        efuseBitDefTable.Rows.FirstOrDefault(
                            p => p[0].EqualsIgnoreCase(datalogItem.Id))!;
                    if (int.Parse(bdfItem[efuseBitDefTable.BitWidthIdx]) != datalogItem.RawData.Length)
                    {
                        wrongBitItems.Add(datalogItem.Id, datalogItem.RawData);
                    }

                }
                else
                {
                    missingitems.Add(key);
                }
            }
            if (bitStream.Count != 0 && crcItem.BlockName != "UID")
            {
                crcItem.RawBitstream = string.Join("", bitStream);
                crcItem.Offset = offset;
                string crCresult = "";
                if (crcItem.Description.StartsWith("ONE_COMPLEMENT_TARGET"))
                {
                    string refFieldName = FieldNameBracketRegex().Match(crcItem.Description).Groups["fieldName"].ToString();
                    EfuseDatalogItem? datalogItem = xDiceInfo.AllReadFromDssc.FirstOrDefault(p => p.Block.EqualsIgnoreCase(crcItem.BlockName) &&
                        p.Id.EqualsIgnoreCase(refFieldName));
                    List<string>? bdfItem =
                        efuseBitDefTable.Rows.FirstOrDefault(
                            p => p[0].EqualsIgnoreCase(refFieldName));
                    if (datalogItem != null && bdfItem != null)
                    {
                        crCresult = CrcItem.CalcCrcWithOneComplement(datalogItem.RawData, int.Parse(bdfItem[efuseBitDefTable.BitWidthIdx]));
                    }
                }
                else
                {
                    crCresult = crcItem.CalcCrc();
                }
                string targetCrc = blockItems.First(p => p[efuseBitDefTable.NameIdx].EqualsIgnoreCase(crcItem.FieldName))[efuseBitDefTable.NameIdx];
                XBitDefRow row = XReadEfuseBitDef.GetRow(targetCrc, crcItem.BlockName)!;
                if (targetCrc != null)
                {
                    EfuseDatalogItem datalogCrc = xDiceInfo.AllReadFromDssc.
                        FirstOrDefault(p => p.Block.EqualsIgnoreCase(crcItem.BlockName) && p.Id.EqualsIgnoreCase(targetCrc))!;
                    string crCexpect =
                        datalogCrc.Value;
                    crcItem.ExpectedCrc = datalogCrc.RawData;
                    var crcError = new Error(EfuseCheckCmdLibError.E_MismatchValue_20, EnumErrorLevel.Info, "", row.RowNo, dieIndx, "tool:CRC = " + crCresult);
                    excelWorksheet.Comments.Add(excelWorksheet.Cells[row.RowNo, dieIndx], crcError.Message, "TAutoGen");
                    if (!CompareCRC(crCexpect, crCresult, crcItem.Bitwidth))
                    {
                        Writer.HighLightCell(excelWorksheet, row.RowNo, dieIndx, row.RowNo, dieIndx, Color.Red);
                        var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_20, EnumErrorLevel.Error, excelWorksheet.Name, row.RowNo, dieIndx, "");
                        EFuseAppExcelWriter.AddComment(appendRichText, excelWorksheet, error);
                    }
                }
            }
        }

        private static string GetCfgBin(List<EfuseDatalogItem> efuseDatalogItems)
        {
            var cfgConditionXx =
                efuseDatalogItems.Where(a => CfgConditionRegex().IsMatch(a.Id)).Select(a => a.RawData).ToList();
            cfgConditionXx.Reverse();
            string xStr = string.Join("", cfgConditionXx);
            return xStr;
        }

        private static string GetCfg(string cfxXx, int msb, int lsb)
        {
            int len = cfxXx.Length;
            string cfgStr = cfxXx.Substring(len - msb - 1, msb - lsb + 1);
            // BinStr2Hex(cfgStr);
            return cfgStr;
        }

        //Check efuse rule : Check data with Config Condition table
        public static void CheckCfgCondition(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, XDiceInfo xDiceInfo, EfuseBitDefTable efuseBitDefTable, EfuseCfgTable efuseCfgTable, int dieIndx)
        {
            int startrow = 3;
            excelWorksheet.Cells[startrow, dieIndx].Value = $"X{xDiceInfo.XCoor}_Y{xDiceInfo.YCoor}";
            startrow += 2;
            var allConfigItems = xDiceInfo.AllReadFromDssc.Where(p => ConfigRegex().IsMatch(p.Block)).ToList();
            string cfgXxx = GetCfgBin(allConfigItems);
            List<string>? tableItem;
            EfuseCheckResultType result = EfuseCheckResultType.Pass;
            //Search for CFG_CONDITION

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
                string currentValue = EfuseCmdUtility.GetCurrentValue(ref value, item, efuseCfgTable);
                int itemIndx = efuseCfgTable.CfgRows.IndexOf(item) + startrow;
                if (efuseCfgTable.Scenario.TryGetValue(XParseDatalog.ScenarioInDatalog, out int value1))
                {
                    string cfgValue = item[value1];
                    bool isProgrammed = GetJobOrder(item[efuseCfgTable.PrgStageIdx]) <= XParseDatalog.TestCat;
                    {

                        if (!CompareCFGTableWithDatalog(ref currentValue, cfgValue, isProgrammed))
                        {
                            var cfgError = new Error(EfuseCheckCmdLibError.E_MismatchValue_31, EnumErrorLevel.Error, "", itemIndx, dieIndx, EfuseCheckCmdLibError.E_MismatchValue_31.MessageTemplate);
                            EFuseAppExcelWriter.PrintError(appendRichText, excelWorksheet, cfgError, Color.Red);
                            result = EfuseCheckResultType.Fail;
                        }
                        else
                        {
                            var cfgError = new Error(EfuseCheckCmdLibError.E_MismatchConfig_01, EnumErrorLevel.Error, "", itemIndx, dieIndx, "");
                            Writer.HighLightCell(excelWorksheet, cfgError.RowNum, cfgError.ColNum, cfgError.RowNum, cfgError.ColNum, Color.LightGreen);
                        }
                    }
                }
                else
                {
                    var cfgError = new Error(EfuseCheckCmdLibError.E_MissingValue_03, EnumErrorLevel.Error, "", itemIndx, dieIndx, $"Scenario:{XParseDatalog.ScenarioInDatalog} Not Found");
                    EFuseAppExcelWriter.PrintErrorComment(appendRichText, excelWorksheet, cfgError);
                    result = EfuseCheckResultType.Fail;
                }
                excelWorksheet.Cells[itemIndx, dieIndx].Value = currentValue;
                if (EfuseStatic.Result == EfuseCheckResultType.Pass)
                {
                    EfuseStatic.Result = result;
                }
            }

            //Search for Others
            foreach (EfuseDatalogItem item in allConfigItems)
            {
                string id = item.Id;
                string value = item.RawData;

                tableItem = efuseCfgTable.CfgRows.FirstOrDefault(p => p[efuseCfgTable.ConditionIdx].EqualsIgnoreCase(id));

                if (tableItem != null)
                {
                    string currentValue = EfuseCmdUtility.GetCurrentValue(ref value, tableItem, efuseCfgTable);
                    int itemIndx = efuseCfgTable.CfgRows.IndexOf(tableItem) + startrow;
                    if (efuseCfgTable.Scenario.TryGetValue(XParseDatalog.ScenarioInDatalog, out int value1))
                    {
                        string cfgValue = tableItem[value1];
                        bool isProgrammed = GetJobOrder(tableItem[efuseCfgTable.PrgStageIdx]) <= XParseDatalog.TestCat;
                        {
                            if (!CompareCFGTableWithDatalog(ref currentValue, cfgValue, isProgrammed))
                            {
                                var cfgError1 = new Error(EfuseCheckCmdLibError.E_MismatchValue_31, EnumErrorLevel.Error, "", itemIndx, dieIndx, EfuseCheckCmdLibError.E_MismatchValue_31.MessageTemplate);
                                EFuseAppExcelWriter.PrintError(appendRichText, excelWorksheet, cfgError1, Color.Red);
                                result = EfuseCheckResultType.Fail;
                            }
                            else
                            {
                                var cfgError = new Error(EfuseCheckCmdLibError.E_MismatchConfig_02, EnumErrorLevel.Error, "", itemIndx, dieIndx, "");
                                Writer.HighLightCell(excelWorksheet, cfgError.RowNum, cfgError.ColNum, cfgError.RowNum, cfgError.ColNum, Color.LightGreen);
                            }
                        }
                    }
                    else
                    {
                        var cfgError1 = new Error(EfuseCheckCmdLibError.E_MissingValue_03, EnumErrorLevel.Error, "", itemIndx, dieIndx, $"Scenario:{XParseDatalog.ScenarioInDatalog} Not Found");
                        EFuseAppExcelWriter.PrintErrorComment(appendRichText, excelWorksheet, cfgError1);
                        result = EfuseCheckResultType.Fail;
                    }
                    excelWorksheet.Cells[itemIndx, dieIndx].Value = currentValue;
                    //Put in the end to prevent value format difference..
                }
                if (EfuseStatic.Result == EfuseCheckResultType.Pass)
                {
                    EfuseStatic.Result = result;
                }
            }
        }
        #endregion

        #region Convert
        public static string ConvertToHex(string valueStr)
        {
            string retVal = "";
            valueStr = valueStr.Replace("_", "");
            string format;
            string tmpvalue;
            if (BinRegex().IsMatch(valueStr))
            {
                format = BinRegex().Match(valueStr).Groups["format"].ToString();
                tmpvalue = valueStr.Replace(format, "");
                var result = new StringBuilder((tmpvalue.Length / 8) + 1);

                // TODO: check all 1's or 0's... throw otherwise

                int mod4Len = tmpvalue.Length % 8;
                if (mod4Len != 0)
                {
                    // pad to length multiple of 8
                    tmpvalue = tmpvalue.PadLeft(((tmpvalue.Length / 8) + 1) * 8, '0');
                }

                for (int i = 0; i < tmpvalue.Length; i += 8)
                {
                    string eightBits = tmpvalue.Substring(i, 8);
                    result.AppendFormat("{0:X2}", Convert.ToByte(eightBits, 2));
                }
                retVal = result.ToString();
            }
            else if (HexRegex().IsMatch(valueStr))
            {
                format = HexRegex().Match(valueStr).Groups["format"].ToString();
                tmpvalue = valueStr.Replace(format, "");
                retVal = "X" + tmpvalue;
            }
            return retVal;
        }

        public static double ConvertValue(Action<string, string> appendRichText, string valueStr)
        {
            double retVal = -1.0;
            try
            {
                //const string _regAnyChar = @"[a-zA-Z]+";
                string format = "";
                string tmpvalue = "";
                valueStr = valueStr.Replace("_", "");
                if (BinRegex().IsMatch(valueStr))
                {
                    format = BinRegex().Match(valueStr).Groups["format"].ToString();
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
                else if (HexRegex().IsMatch(valueStr))
                {
                    format = HexRegex().Match(valueStr).Groups["format"].ToString();
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

                if (DecRegex().IsMatch(valueStr))
                {
                    string value = DecRegex().Match(valueStr).Groups["value"].Value;
                    string unit = DecRegex().Match(valueStr).Groups["unit"].Value;
                    if (!string.IsNullOrEmpty(value))
                    {
                        retVal = double.Parse(value) * ConvertUnitStr2Double(unit);
                    }
                    else
                    {
                        retVal = ValueOutOfRange;
                    }
                }
                else
                {
                    retVal = ValueOutOfRange;
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
                EfuseStatic.Result = EfuseCheckResultType.Exception;
                if (EfuseStatic.IsCmd)
                {
                    appendRichText(string.Format(ex.ToString()), "Red");
                }
                else
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }
            return retVal;
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

        public static string CheckAliasBlock(string blockName, EfuseScriptConfig efuseScriptConfig)
        {
            string result = blockName ?? throw new ArgumentNullException(nameof(blockName));

            if (efuseScriptConfig.AliasBlockList.TryGetValue(blockName, out string? value))
            {
                result = value;
            }

            return result;
        }

        public static bool IsPassBin(int bin, EfuseScriptConfig efuseScriptConfig)
        {
            bool result = efuseScriptConfig.PassBin.Contains(bin.ToString("D"));

            return result;
        }

        public static bool CheckEfuseDefaultValue(Action<string, string> appendRichText, string val, string defValue, bool isNeedReverse = false)

        {
            if (defValue.EqualsIgnoreCase("NA") || defValue.EqualsIgnoreCase("N/A"))
            {
                defValue = "0";
            }

            return ConvertValue(appendRichText, val) == ConvertValue(appendRichText, defValue);
        }

        public static bool CheckEfuseLimit(double val, double lLim, double hLim)
        {
            //STEP1. 
            if (val == ValueOutOfRange)
            {
                return true;
            }

            if (lLim != ValueOutOfRange)
            {
                if (val < lLim)
                {
                    return false;
                }
            }

            if (hLim != ValueOutOfRange)
            {
                if (val > hLim)
                {
                    return false;
                }
            }

            return true;
        }
        #endregion

        private static bool CompareCRC(string expect, string real, int bitlength)
        {
            if (expect.ContainsIgnoreCase("0x"))
            {
                return expect == real;
            }

            if (ConvertStr2Hex(expect, 10, bitlength) != real)
            {
                return ConvertStr2Hex(expect, 16, bitlength) == real;
            }

            return true;

        }

        private static string ConvertStr2Hex(string strInput, int strFormat, int bitlength)
        {
            string result = "";
            int intParse;
            string hexFormat = bitlength == 16 ? "X4" : "X8";
            switch (strFormat)
            {
                case 10:

                    if (int.TryParse(strInput, NumberStyles.Integer, CultureInfo.InvariantCulture, out intParse))
                    {
                        result = intParse.ToString(hexFormat);
                    }

                    break;
                case 16:

                    if (int.TryParse(strInput, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out intParse))
                    {
                        result = intParse.ToString(hexFormat);
                    }

                    break;

            }

            return "0x" + result;
        }

        private static bool CompareCFGTableWithDatalog(ref string realValue, string cfgValue, bool isCompareDefault)
        {
            string regPrefix = @"(?<prefix>(\d*[Xx]+)|(0*[Bb]+0*))";
            string hexPrefix = Regex.Match(cfgValue, regPrefix).Groups["prefix"].ToString();
            if (!isCompareDefault)
            {
                cfgValue = hexPrefix + "0";
            }

            if (!Regex.IsMatch(cfgValue, regPrefix, RegexOptions.IgnoreCase))
            //if not Hex format, it would be binary and convert to hex
            {
                hexPrefix = "";
                try
                {

                    cfgValue = hexPrefix + Convert.ToInt32(cfgValue, 2).ToString("X");
                }
                catch (Exception ex)
                {
                    EfuseStatic.Result = EfuseCheckResultType.Exception;
                    if (!EfuseStatic.IsCmd)
                    {
                        ErrorMessageBox.Show(string.Format(ex.ToString()));
                    }

                    cfgValue = hexPrefix + cfgValue;
                }
            }

            realValue = hexPrefix + Convert.ToInt32(realValue, 2).ToString("X");
            return Convert.ToInt32(realValue, 16) == Convert.ToInt32(cfgValue, 16);
        }

        public static EfuseCfgTable CfgTableSel()
        {
            EfuseCfgTable table;
            if (XParseDatalog.IsUseSvm)
            {
                table = EfuseCfgTableReader1.CfgTable.FirstOrDefault(p => SvmRegex().IsMatch(p.Key)).Value;
            }
            else
            {
                table = EfuseCfgTableReader1.CfgTable.ElementAt(0).Value;
            }

            table ??= EfuseCfgTableReader1.CfgTable.ElementAt(0).Value;

            return table;
        }

        public static bool CompareUdrcmpItems(Action<string, string> appendRichText, string itemBlock, string itemId, XDiceInfo xDiceInfo, LoaderEfuseBitDef loaderEfuseBitDef, EfuseScriptConfig efuseScriptConfig)
        {
            string aliasBlock = "";
            if (efuseScriptConfig.AliasBlockList.ContainsValue(itemBlock))
            {
                aliasBlock = efuseScriptConfig.AliasBlockList.FirstOrDefault(p => p.Value == itemBlock).Key;
            }

            string refblock = efuseScriptConfig.UdrcmpMapping[itemBlock];
            try
            {
                double refValue = 0;
                double udrValue = 0;
                double cmpValue = 0;
                EfuseDatalogItem? refItem =
                    xDiceInfo.AllReadFromDssc.FirstOrDefault(p => p.Block == refblock && p.Id == itemId);
                EfuseDatalogItem? cmpItem =
                    xDiceInfo.AllReadFromDssc.FirstOrDefault(
                        p => (p.Block == itemBlock || p.Block == aliasBlock)
                             && p.Id == itemId);
                if (refItem != null && cmpItem != null)
                {
                    List<string> targetRow = loaderEfuseBitDef.BitDefTable.Rows
                        .FirstOrDefault(
                            p =>
                                p[loaderEfuseBitDef.BitDefTable.NameIdx].EqualsIgnoreCase(itemId) &&
                                p[loaderEfuseBitDef.BitDefTable.BlockIdx].EqualsIgnoreCase(refblock))!;

                    cmpValue =
                        double.Parse(ConvertValueToSpecifyFormat(appendRichText, cmpItem.Value, EnumValueType.Dec));
                    udrValue =
                        double.Parse(ConvertValueToSpecifyFormat(appendRichText, refItem.Value, EnumValueType.Dec));
                    if (udrValue == cmpValue)
                    {
                        return true;
                    }

                    if (!targetRow[loaderEfuseBitDef.BitDefTable.AlgorithmIdx].EqualsIgnoreCase("vddbin") &&
                        !BaseVoltageRegex().IsMatch(itemId))
                    {
                        return false;
                    }
                    //compare bincut items
                    double resolution = double.Parse(targetRow[loaderEfuseBitDef.BitDefTable.ResolutionIdx]);

                    if (udrValue == cmpValue)
                    {
                        return true;
                    }

                    refValue = (udrValue - BaseVoltage) / resolution;
                    if (refValue == cmpValue)
                    {
                        return true;
                    }

                    refValue = (udrValue - ((BaseVoltage + 1) * efuseScriptConfig.BaseVoltageResolution)) / resolution;
                    if (refValue == cmpValue)
                    {
                        return true;
                    }

                    refValue = (udrValue / resolution) - 1;
                    return refValue == cmpValue;
                }
            }
            catch (Exception ex)
            {
                EfuseStatic.Result = EfuseCheckResultType.Exception;
                if (EfuseStatic.IsCmd)
                {
                    appendRichText(string.Format(ex.ToString()), "Red");
                }
                else
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }
            return false;
        }

        public static bool CompareUdrcmpItemsMp(Action<string, string> appendRichText, string itemBlock, string itemId, XDiceInfo xDiceInfo, LoaderEfuseBitDef loaderEfuseBitDef, XBitDefRow xBitDefRow, int currentStage, EfuseScriptConfig efuseScriptConfig)
        {
            if (efuseScriptConfig.AliasBlockList.ContainsValue(itemBlock))
            {
            }

            string refblock = efuseScriptConfig.UdrcmpMapping[itemBlock];
            try
            {
                double refValue = 0;
                double udrValue = 0;
                double cmpValue = 0;
                XBitDefRow? defRow = XReadEfuseBitDef.BitDefTable.FirstOrDefault(
                    items => items.Block == refblock && items.Name.Split('#')[1] == itemId);
                if (defRow == null)
                {
                    return false;
                }

                EfuseDatalogItem? refItem = xDiceInfo.FuseMpDataSet!.GetData(refblock, defRow.Msb, defRow.Lsb, defRow.Algorithm, defRow.Resolution, defRow.JobStage <= currentStage ? BaseVoltage.ToString() : "0");
                EfuseDatalogItem? cmpItem = xDiceInfo.FuseMpDataSet!.GetData(itemBlock, xBitDefRow.Msb, xBitDefRow.Lsb, refItem != null && !string.IsNullOrEmpty(defRow.Resolution) ? defRow.Algorithm : xBitDefRow.Algorithm, refItem != null && !string.IsNullOrEmpty(defRow.Algorithm) ? defRow.Resolution : xBitDefRow.Resolution, defRow.JobStage <= currentStage ? BaseVoltage.ToString() : "0");
                if (refItem != null && cmpItem != null)
                {
                    List<string> targetRow = loaderEfuseBitDef.BitDefTable.Rows
                        .FirstOrDefault(
                            p =>
                                p[loaderEfuseBitDef.BitDefTable.NameIdx].EqualsIgnoreCase(itemId) &&
                                p[loaderEfuseBitDef.BitDefTable.BlockIdx].EqualsIgnoreCase(refblock))!;

                    cmpValue =
                        double.Parse(ConvertValueToSpecifyFormat(appendRichText, cmpItem.Value, EnumValueType.Dec));
                    udrValue =
                        double.Parse(ConvertValueToSpecifyFormat(appendRichText, refItem.Value, EnumValueType.Dec));
                    if (udrValue == cmpValue)
                    {
                        return true;
                    }

                    if (!targetRow[loaderEfuseBitDef.BitDefTable.AlgorithmIdx].EqualsIgnoreCase("vddbin") &&
                        !BaseVoltageRegex().IsMatch(itemId))
                    {
                        return false;
                    }
                    //compare bincut items
                    double resolution = double.Parse(targetRow[loaderEfuseBitDef.BitDefTable.ResolutionIdx]);

                    if (udrValue == cmpValue)
                    {
                        return true;
                    }

                    refValue = (udrValue - BaseVoltage) / resolution;
                    if (refValue == cmpValue)
                    {
                        return true;
                    }

                    refValue = (udrValue - ((BaseVoltage + 1) * efuseScriptConfig.BaseVoltageResolution)) / resolution;
                    if (refValue == cmpValue)
                    {
                        return true;
                    }

                    refValue = (udrValue / resolution) - 1;
                    return refValue == cmpValue;
                }
            }
            catch (Exception ex)
            {
                EfuseStatic.Result = EfuseCheckResultType.Exception;
                if (EfuseStatic.IsCmd)
                {
                    appendRichText(string.Format(ex.ToString()), "Red");
                }
                else
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }
            return false;
        }

        public static EnumValueType JudgeValueType(string value)
        {

            EnumValueType type = EnumValueType.Dec;
            if (HexRegex().IsMatch(value))
            {
                return EnumValueType.Hex;
            }

            if (BinRegex().IsMatch(value))
            {
                return EnumValueType.Bin;
            }

            return type;
        }

        public static string ConvertValueToSpecifyFormat(Action<string, string> appendRichText, string value, EnumValueType enumValueType)
        {
            string result = value;
            try
            {
                if (DecimalPointRegex().IsMatch(value))
                {
                    result = double.Parse(value).ToString();
                }

                string format;
                if (enumValueType == EnumValueType.Bin)
                {
                    if (BinRegex().IsMatch(value))
                    {
                        format = BinRegex().Match(value).Groups["format"].ToString();
                        result = value.Replace(format, "");
                        result = Convert.ToString(Convert.ToInt32(result, 2));
                    }
                    try
                    {
                        return "B" + Convert.ToString(Convert.ToInt32(result, 10), 2);
                    }
                    catch (Exception)
                    {
                    }
                }
                if (enumValueType == EnumValueType.Hex)
                {
                    if (HexRegex().IsMatch(value))
                    {
                        format = HexRegex().Match(value).Groups["format"].ToString();
                        result = value.Replace(format, "");
                        return "X" + result;
                    }
                    return "X" + Convert.ToString(Convert.ToInt32(result, 10), 16).ToUpper();
                }

                if (enumValueType == EnumValueType.Dec)
                {
                    if (BinRegex().IsMatch(value))
                    {
                        format = BinRegex().Match(value).Groups["format"].ToString();
                        result = value.Replace(format, "");
                        result = Convert.ToString(Convert.ToInt32(result, 2));
                    }
                    else if (HexRegex().IsMatch(value))
                    {
                        format = HexRegex().Match(value).Groups["format"].ToString();
                        result = value.Replace(format, "");
                        result = Convert.ToString(Convert.ToInt32(result, 16));
                    }
                }
            }
            catch (Exception ex)
            {
                EfuseStatic.Result = EfuseCheckResultType.Exception;
                if (EfuseStatic.IsCmd)
                {
                    appendRichText(string.Format(ex.ToString()), "Red");
                }
                else
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }
            return result;
        }

        public static bool CompareValue(Action<string, string> appendRichText, string val1, string val2)
        {
            double val1Conv = Math.Round(ConvertValue(appendRichText, val1), 6);
            double val2Conv = Math.Round(ConvertValue(appendRichText, val2), 6);
            return val1Conv.Equals(val2Conv);
        }

        public static string GetProberHexcode(XDiceInfo xDiceInfo, bool isNeedReversed = false)
        {
            string binstr = "";

            if (xDiceInfo.AllReadFromDssc == null)
            {
                return binstr;
            }

            if ((xDiceInfo.AllReadFromDssc.FirstOrDefault(p => p.Id.EqualsIgnoreCase("x_coordinate")) == null && xDiceInfo.AllReadFromDssc.FirstOrDefault(p => p.Id.EqualsIgnoreCase("y_coordinate")) == null) || xDiceInfo.AllReadFromDssc.FirstOrDefault(p => p.Id.ContainsIgnoreCase("wafer")) == null || xDiceInfo.AllReadFromDssc.FirstOrDefault(p => p.Id.ContainsIgnoreCase("lot")) == null)
            {
                return binstr;
            }
            string x = xDiceInfo.AllReadFromDssc.FirstOrDefault(p => p.Id.EqualsIgnoreCase("x_coordinate") && XReadEfuseBitDef.GetRow(p.Id, p.Block)!
                            .RealStage.ContainsIgnoreCase("cp1"))!.RawData;
            string y = xDiceInfo.AllReadFromDssc.FirstOrDefault(p => p.Id.EqualsIgnoreCase("y_coordinate") && XReadEfuseBitDef.GetRow(p.Id, p.Block)!
                            .RealStage.ContainsIgnoreCase("cp1"))!.RawData;
            string wafer = xDiceInfo.AllReadFromDssc.FirstOrDefault(p => p.Id.ContainsIgnoreCase("wafer") && XReadEfuseBitDef.GetRow(p.Id, p.Block)!
                            .RealStage.ContainsIgnoreCase("cp1"))!.RawData;
            var lotList =
                xDiceInfo.AllReadFromDssc.FindAll(
                    p =>
                        XReadEfuseBitDef.GetRow(p.Id, p.Block)!
                            .Algorithm.EqualsIgnoreCase("lotid") && XReadEfuseBitDef.GetRow(p.Id, p.Block)!
                            .RealStage.ContainsIgnoreCase("cp1")).Select(q => q.RawData).Distinct().ToList();

            if (isNeedReversed)
            {
                binstr = new string([.. y.Reverse()]) + new string([.. x.Reverse()]) +
                         new string([.. wafer.Reverse()]);
                for (int idx = lotList.Count - 1; idx >= 0; idx--)
                {
                    binstr += new string([.. lotList[idx].Reverse()]);
                }
            }
            else
            {
                binstr = y + x + wafer;
                for (int idx = lotList.Count - 1; idx >= 0; idx--)
                {
                    binstr += new string([.. lotList[idx]]);
                }
            }
            if (binstr.EndsWith("000000"))
            {
                binstr = binstr[..^6];
            }
            binstr = binstr.PadLeft(64, '0');

            return Convert.ToString(Convert.ToInt64(binstr, 2), 16).PadLeft(16, '0').ToUpper();
        }

        public static string GetProberHexcode(string lot, string wafer, string x, string y, bool isNeedReversed = false)
        {
            lot = Hex2Binary(lot, 36, true);
            wafer = Hex2Binary(wafer, 5);
            x = Hex2Binary(x, 6);
            y = Hex2Binary(y, 6);

            string binstr;
            if (isNeedReversed)
            {
                binstr = new string([.. y.Reverse()]) + new string([.. x.Reverse()]) +
                         new string([.. wafer.Reverse()]) + new string([.. lot.Reverse()]);
            }
            else
            {
                binstr = y + x + wafer + new string([.. lot]);
            }

            binstr = binstr.PadLeft(64, '0');

            return Convert.ToString(Convert.ToInt64(binstr, 2), 16).PadLeft(16, '0').ToUpper();
        }

        public static string Hex2Binary(string hex, int bitlength, bool isLotId = false)
        {
            var result = new StringBuilder();
            byte[] byteData = Encoding.ASCII.GetBytes(hex);
            if (isLotId)
            {
                for (int x = 0; x <= byteData.Length - 1; x++)
                {
                    Regex reg = UppercaseLetterRegex();
                    int val = Convert.ToInt32(byteData[x]);
                    string temp;
                    if (reg.IsMatch(hex.AsSpan(x, 1)))
                    {
                        val -= 55;
                        temp = Convert.ToString(val, 2).PadLeft(6, '0');
                    }
                    else
                    {
                        temp = Convert.ToString(Convert.ToInt64(hex.Substring(x, 1), 16), 2).PadLeft(6, '0');
                    }
                    result.Append(temp);
                }
            }
            else
            {
                _ = int.TryParse(hex, out int desc);
                result.Append(Convert.ToString(desc, 2).PadLeft(bitlength, '0'));
            }
            string str = result.ToString().PadLeft(bitlength, '0');

            return str;
        }

        public static string CalcEccValue(bool isMsbFirst, string alg, string rawData)
        {
            if (string.IsNullOrEmpty(alg) && string.IsNullOrEmpty(rawData))
            {
                return "";
            }

            if (isMsbFirst)
            {
                rawData = new string([.. rawData.Reverse()]);
            }

            var dataBytes = new List<byte>();
            string result = "";

            if (alg.EqualsIgnoreCase("SEC"))
            {
                byte[] eccCalcResult = new byte[6];

                foreach (char dataStr in rawData)
                {
                    if (dataStr == '0')
                    {
                        dataBytes.Add(0);
                    }
                    else if (dataStr == '1')
                    {
                        dataBytes.Add(1);
                    }
                }

                for (int i = 0; i < SecKeyList.Count; i++)
                {
                    List<int> currentKeys = SecKeyList[i];
                    byte tmp = dataBytes[currentKeys[0]];
                    for (int j = 1; j < currentKeys.Count; j++)
                    {
                        tmp = (byte)(tmp ^ dataBytes[currentKeys[j]]);
                    }
                    eccCalcResult[i] = tmp;
                    result += tmp;
                }

                return isMsbFirst ? new string([.. result.Reverse()]) : result;
            }
            else if (alg.EqualsIgnoreCase("DEC"))
            {
                byte[] eccCalcResult = new byte[12];

                foreach (char dataStr in rawData)
                {
                    if (dataStr == '0')
                    {
                        dataBytes.Add(0);
                    }
                    else if (dataStr == '1')
                    {
                        dataBytes.Add(1);
                    }
                }

                for (int i = 0; i < DecKeyList.Count; i++)
                {
                    List<int> currentKeys = DecKeyList[i];
                    byte tmp = dataBytes[currentKeys[0]];
                    for (int j = 1; j < currentKeys.Count; j++)
                    {
                        tmp = (byte)(tmp ^ dataBytes[currentKeys[j]]);
                    }
                    eccCalcResult[i] = tmp;
                    result += tmp;
                }

                return isMsbFirst ? new string([.. result.Reverse()]) : result;
            }

            return "";
        }

        public static (string data, string value) GetEccAndData(bool isMsbFirst, string alg, string rawData)
        {
            if (string.IsNullOrEmpty(alg))
            {
                return ("NoAlg", "NoAlg");
            }

            if (alg.EqualsIgnoreCase("SEC"))
            {
                if (rawData.Length < 22)
                {
                    return ("", "");
                }
                string data = isMsbFirst ? rawData.Substring(rawData.Length - 16, 16) : rawData[..16];
                string value = isMsbFirst ? rawData[..6] : rawData.Substring(rawData.Length - 6, 6);
                return (data, value);
            }
            else if (alg.EqualsIgnoreCase("DEC"))
            {
                if (rawData.Length < 44)
                {
                    return ("", "");
                }

                string data = "";
                string value = "";
                int eccbitSize = 6;
                int eccbitsDataSize = 16;
                int count = rawData.Length / (eccbitSize + eccbitsDataSize);
                if (isMsbFirst)
                {
                    for (int i = 1; i <= count; i++)
                    {
                        data += rawData.Substring((i * eccbitSize) + ((i - 1) * eccbitsDataSize), eccbitsDataSize);
                        value = string.Concat(rawData.AsSpan(rawData.Length - (i * eccbitsDataSize) - (i * eccbitSize), eccbitSize), value);
                    }
                }
                else
                {
                    for (int i = 0; i < count; i++)
                    {
                        data += rawData.Substring((i * eccbitSize) + (i * eccbitsDataSize), eccbitsDataSize);
                        value += rawData.Substring((i * eccbitSize) + (i * eccbitsDataSize), eccbitSize);
                    }
                }

                return (data, value);
            }
            return ("", "");
        }

        public static string CalcCheckSum(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, CheckSum checkSum, XDiceInfo xDiceInfo, EfuseBitDefTable efuseBitDefTable, int dieIndx)
        {
            var bitStream = new List<uint>();
            var keyList = new List<string>();
            string hexString = "";
            if (!string.IsNullOrEmpty(checkSum.Range))
            {
                foreach (string range in checkSum.Range.Split(','))
                {
                    int msb = int.Parse(range.Split(':')[0].Replace("[", ""));
                    int lsb = int.Parse(range.Split(':')[1].Replace("]", ""));
                    keyList.AddRange(efuseBitDefTable.Rows.Where(x => int.Parse(x[efuseBitDefTable.MsbBitIdx]) <= msb && int.Parse(x[efuseBitDefTable.LsbBitIdx]) >= lsb).Select(p => p[efuseBitDefTable.NameIdx]));
                }
            }
            string binaryStr = "";
            foreach (string key in keyList)
            {

                EfuseDatalogItem? datalogItem =
                    xDiceInfo.AllReadFromDssc.FirstOrDefault(p => p.Id.EqualsIgnoreCase(key));
                if (datalogItem != null)
                {
                    if (!string.IsNullOrEmpty(datalogItem.RawData))
                    {
                        try
                        {
                            binaryStr = datalogItem.RawData + binaryStr;
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }

            for (int binIdx = binaryStr.Length - 1; binIdx >= 0; binIdx -= 8)
            {
                string tempBin = binaryStr.Substring(binIdx - 7, 8);
                hexString += ConvertToHex("0b" + tempBin);
            }

            int hexIdx = 0;
            string tmp = "";
            foreach (char str in hexString)
            {
                tmp += str;
                if (hexIdx % 2 == 1 || hexIdx == hexString.Length - 1)
                {
                    bitStream.Add(Convert.ToUInt32(tmp, 16));
                    tmp = "";
                }
                hexIdx++;
            }

            long checkSumRet = 0;
            foreach (uint data in bitStream)
            {
                checkSumRet = (checkSumRet >> 1) + ((checkSumRet & 1) << 15);
                checkSumRet += data;
                checkSumRet &= 0xFFFF;
            }

            string checksumHex = checkSumRet.ToString("X2");
            XBitDefRow row = XReadEfuseBitDef.GetRow(checkSum.FieldName, checkSum.Bank)!;

            if (checksumHex != null)
            {
                EfuseDatalogItem datalogChecksum = xDiceInfo.AllReadFromDssc.FirstOrDefault(p => p.Id.EqualsIgnoreCase(checkSum.FieldName))!;
                if (!CompareCRC(checksumHex, datalogChecksum.Value, 16))
                {
                    var crcError = new Error(EfuseCheckCmdLibError.E_MismatchValue_18, EnumErrorLevel.Info, "", row.RowNo, dieIndx, "tool:CheckSum = " + checksumHex);
                    excelWorksheet.Comments.Add(excelWorksheet.Cells[row.RowNo, dieIndx], crcError.Message, "TAutoGen");
                    Writer.HighLightCell(excelWorksheet, row.RowNo, dieIndx, row.RowNo, dieIndx, Color.Red);
                    var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_18, EnumErrorLevel.Error, excelWorksheet.Name, row.RowNo, dieIndx, "");
                    EFuseAppExcelWriter.AddComment(appendRichText, excelWorksheet, error);
                }
            }
            return "";
        }
    }
}
