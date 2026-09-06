using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using TestPlanLib.Efuse.Input;

namespace Automation.GenerateIgxl.EFuse.InputChecker
{
    public class EfuseBitDefTableChecker
    {
        internal double _baseVoltage = double.MinValue;

        public void WorkFlow(List<BitDefTable> tables)
        {
            var udrInfo = tables.Where(x => EFuseConst.BankIsUdr(EFuseConst.GetBankName(x.BlockName))).ToList();
            foreach (BitDefTable table in tables)
            {
                CheckBankType(table);

                foreach (BitDefRow row in table.Rows)
                {
                    CheckNonCpStageInEcid(table, row);
                    CheckIfProgramStageIsFtf(table, row);
                    string width = "";
                    CheckBitFormat(table, row, ref width);
                    CheckLimitFormat(table, row, width);
                    CheckCrcFormat(table, row);
                    CheckDefualtValue(table, row);
                    if (table.BlockName.StartsWith("CMP", StringComparison.CurrentCultureIgnoreCase))
                    {
                        CheckCmpBitWidthInUdr(table, row, udrInfo);
                    }
                }
            }
            CheckShadowFieldStage(tables);
        }

        internal void CheckShadowFieldStage(List<BitDefTable> tables)
        {
            var shadowFields = new List<BitDefRow>();
            var nonShadowFields = new List<BitDefRow>();
            int nameIdx = -1;
            int programmingIdx = -1;
            if (tables.Any())
            {
                //get all shadow and non shadow
                foreach (BitDefTable tbl in tables)
                {
                    nameIdx = tbl.BankEfuseBitDefIdx;
                    programmingIdx = tbl.ProgrammingStageIdx;
                    IEnumerable<BitDefRow> fields = tbl.Rows.Where(x => x.RowData[tbl.BankEfuseBitDefIdx].ToLower().EndsWith("_shadow"));
                    shadowFields.AddRange(fields);
                }
                if (shadowFields.Any())
                {
                    foreach (BitDefRow field in shadowFields)
                    {
                        string nonShadowName = field.RowData[nameIdx].Replace("_shadow", "");
                        foreach (BitDefTable tbl in tables)
                        {
                            IEnumerable<BitDefRow> fields = tbl.Rows.Where(x => x.RowData[tbl.BankEfuseBitDefIdx].Equals(nonShadowName, StringComparison.CurrentCultureIgnoreCase));
                            nonShadowFields.AddRange(fields);
                        }
                    }
                }
            }

            foreach (BitDefRow nonShadowField in nonShadowFields)
            {
                BitDefRow shadowField = shadowFields.FirstOrDefault(x => x.RowData[nameIdx].Equals(nonShadowField.RowData[nameIdx] + "_shadow", StringComparison.CurrentCultureIgnoreCase));
                if (shadowField != null)
                {
                    if (shadowField.RowData[programmingIdx].Equals(nonShadowField.RowData[programmingIdx], StringComparison.CurrentCultureIgnoreCase))
                    {
                        ErrorReportManager.AddError(EFuseErrorType.W_RuleViolationStage_03, shadowField.SheetName, shadowField.RowNum, nameIdx + 1, [shadowField.RowData[nameIdx], nonShadowField.RowData[nameIdx]]);
                    }
                }
            }
        }

        internal void CheckBankType(BitDefTable table)
        {
            if (EFuseConst.GetBankName(table.BlockName).Equals(BankType.Unknow))
            {
                ErrorReportManager.AddError(EFuseErrorType.W_RuleViolationBank_01, table.SheetName, table.HeaderRowNum, table.BankEfuseBitDefIdx + 1, [table.BlockName]);
            }
        }

        internal void CheckNonCpStageInEcid(BitDefTable table, BitDefRow row)
        {
            if (EFuseConst.GetBankName(table.BlockName) != BankType.Ecid)
            {
                return;
            }


            if (table.ProgrammingStageIdx != -1 && row.RowData.Count > table.ProgrammingStageIdx
                                                && !row.RowData[table.ProgrammingStageIdx].StartsWith("CP", StringComparison.CurrentCultureIgnoreCase))
            {
                ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationStage_01, table.SheetName, table.HeaderRowNum, table.ProgrammingStageIdx + 1, [row.RowData[table.ProgrammingStageIdx]]);
            }
        }

        internal void CheckIfProgramStageIsFtf(BitDefTable table, BitDefRow row)
        {
            if (table.ProgrammingStageIdx != -1 && row.RowData.Count > table.ProgrammingStageIdx
                                                && row.RowData[table.ProgrammingStageIdx].Equals("FTF", StringComparison.CurrentCultureIgnoreCase))
            {
                ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationStage_02, table.SheetName, table.HeaderRowNum, table.ProgrammingStageIdx + 1, []);
            }
        }

        internal void CheckBitFormat(BitDefTable table, BitDefRow row, ref string width)
        {
            bool allBitIsNumber = true;
            string lsb = "";
            string msb = "";
            width = "";
            if (table.LsbBitIdx != -1 && row.RowData.Count > table.LsbBitIdx)
            {
                lsb = row.RowData[table.LsbBitIdx];
                if (!lsb.IsNumber())
                {
                    allBitIsNumber = false;
                    ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationBit_01, row.SheetName, row.RowNum, table.LsbBitIdx + 1, [lsb]);
                }
            }

            if (table.MsbBitIdx != -1 && row.RowData.Count > table.MsbBitIdx)
            {
                msb = row.RowData[table.MsbBitIdx];
                if (!msb.IsNumber())
                {
                    allBitIsNumber = false;
                    ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationBit_02, row.SheetName, row.RowNum, table.LsbBitIdx + 1, [msb]);
                }
            }

            if (table.BitWidthIdx != -1 && row.RowData.Count > table.BitWidthIdx)
            {
                width = row.RowData[table.BitWidthIdx];
                if (!width.IsNumber())
                {
                    allBitIsNumber = false;
                    ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationBit_03, row.SheetName, row.RowNum, table.BitWidthIdx + 1, [width]);
                }
            }

            if (allBitIsNumber)
            {
                string calWidth = lsb.CalBitWidth(msb);
                if (!calWidth.Equals(width, StringComparison.CurrentCultureIgnoreCase))
                {
                    ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationBit_04, row.SheetName, row.RowNum, table.BitWidthIdx + 1, [calWidth]);
                }
                else
                {
                    if (double.TryParse(width, out double doubleWidth))
                    {
                        if (doubleWidth > 1024)
                        {
                            ErrorReportManager.AddError(EFuseErrorType.W_InvalidBit_01, row.SheetName, row.RowNum, table.BitWidthIdx + 1, []);
                        }
                    }
                }
            }

        }

        internal void CheckLimitFormat(BitDefTable table, BitDefRow row, string width)
        {
            bool limitIsNummber = true;
            string lowLimit = "";
            string highLimit = "";
            string defaultValue = "";
            if (!EFuseConst.GetBankName(table.BlockName).Equals(BankType.Cfg) &&
                table.DefaultOrRealIdx != -1 && row.RowData[table.DefaultOrRealIdx].Equals("real", StringComparison.CurrentCultureIgnoreCase))
            {
                if (table.LowLimitIdx != -1 && row.RowData.Count > table.LowLimitIdx)
                {
                    lowLimit = row.RowData[table.LowLimitIdx];
                    if (!lowLimit.Trim().EFuseBitToInt().IsNumber())
                    {
                        limitIsNummber = false;
                        ErrorReportManager.AddError(EFuseErrorType.E_InvalidLimit_01, row.SheetName, row.RowNum, table.LowLimitIdx + 1, [lowLimit]);
                    }
                }

                if (table.HighLimitIdx != -1 && row.RowData.Count > table.HighLimitIdx)
                {
                    highLimit = row.RowData[table.HighLimitIdx];
                    if (!row.RowData[table.HighLimitIdx].Trim().EFuseBitToInt().IsNumber())
                    {
                        limitIsNummber = false;
                        ErrorReportManager.AddError(EFuseErrorType.E_InvalidLimit_02, row.SheetName, row.RowNum, table.HighLimitIdx + 1, [highLimit]);
                    }
                }
            }

            if (table.DefaultOrRealIdx != -1 &&
                row.RowData[table.DefaultOrRealIdx].Equals("Default", StringComparison.CurrentCultureIgnoreCase))
            {
                if (table.DefaultValueIdx != -1 && row.RowData.Count > table.DefaultValueIdx)
                {
                    defaultValue = row.RowData[table.DefaultValueIdx];
                    if (!defaultValue.Trim().EFuseBitToInt().IsNumber() && !defaultValue.IsHexOrOctValue())
                    {
                        limitIsNummber = false;
                        ErrorReportManager.AddError(EFuseErrorType.W_InvalidValue_01, row.SheetName, row.RowNum, table.DefaultValueIdx + 1, [defaultValue]);
                    }
                }
            }


            if (limitIsNummber)
            {
                CheckLimitConsistency(table, row, width, lowLimit, highLimit, defaultValue);
            }


        }

        internal void CheckLimitConsistency(BitDefTable table, BitDefRow row, string width, string lowLimit, string highLimit, string defaultValue)
        {

            if (table.DefaultOrRealIdx != -1 && row.RowData[table.DefaultOrRealIdx].Equals("real", StringComparison.CurrentCultureIgnoreCase) &&
                table.AlgorithmIdx != -1 && !row.RowData[table.AlgorithmIdx].Equals("crc", StringComparison.CurrentCultureIgnoreCase))
            {
                if (double.TryParse(highLimit, out double dHigh) &&
                    double.TryParse(lowLimit, out double dLow))
                {
                    if (dHigh <= dLow)
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_InvalidLimit_03, row.SheetName, row.RowNum, table.HighLimitIdx + 1, []);
                    }
                }
            }

            if (table.AlgorithmIdx != -1 && row.RowData.Count > table.AlgorithmIdx && row.RowData[table.AlgorithmIdx].Equals("ids", StringComparison.CurrentCultureIgnoreCase)
                && table.ResolutionIdx != -1 && row.RowData.Count > table.ResolutionIdx &&
                !string.IsNullOrEmpty(width) && double.TryParse(width, out double dWidth))
            {
                if (double.TryParse(row.RowData[table.ResolutionIdx], out double dResolution) && double.TryParse(highLimit, out double dHigh))
                {
                    if ((dResolution * (Math.Pow(2, dWidth) - 1)) + dResolution < dHigh)
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_InvalidLimit_04, row.SheetName, row.RowNum, table.HighLimitIdx + 1, []);
                    }
                }
            }

            CheckBaseAlgorithmLimit(table, row, lowLimit, highLimit, defaultValue);

        }

        internal void CheckBaseAlgorithmLimit(BitDefTable table, BitDefRow row, string lowLimit, string highLimit, string defaultValue)
        {
            if (table.AlgorithmIdx != -1 &&
                row.RowData[table.AlgorithmIdx].Equals("base", StringComparison.CurrentCultureIgnoreCase))
            {
                if (double.TryParse(highLimit, out double dHigh) &&
                    double.TryParse(lowLimit, out double dLow) &&
                    double.TryParse(defaultValue, out double dDefault))
                {
                    if (Equals(_baseVoltage, double.MinValue))
                    {
                        _baseVoltage = dLow;
                    }

                    if (!Equals(dLow, dHigh) || !Equals(dLow, dDefault) || !Equals(dHigh, dDefault))
                    {
                        //string message = "If Algorithm is base, Low Limit and High Limit and Default Value must be same.";
                        int col = table.HighLimitIdx;
                        if (Equals(dLow, dHigh) && !Equals(dLow, dDefault))
                        {
                            col = table.DefaultValueIdx;
                        }
                        else if (Equals(dLow, dDefault) && !Equals(dLow, dHigh))
                        {
                            col = table.HighLimitIdx;
                        }
                        else if (Equals(dHigh, dDefault) && !Equals(dLow, dHigh))
                        {
                            col = table.HighLimitIdx;
                        }

                        ErrorReportManager.AddError(EFuseErrorType.W_RuleViolationLimit_01, row.SheetName, row.RowNum, col + 1, []);
                    }
                    else if (!Equals(_baseVoltage, dLow))
                    {
                        ErrorReportManager.AddError(EFuseErrorType.W_MismatchRow_01, row.SheetName, row.RowNum, table.LowLimitIdx + 1, []);
                    }
                }
            }
        }

        internal void CheckCrcFormat(BitDefTable table, BitDefRow row)
        {
            string regex = @"\[\d+:\d+\]";
            if (table.AlgorithmIdx != -1 && row.RowData.Count > table.AlgorithmIdx &&
                row.RowData[table.AlgorithmIdx].Equals("crc", StringComparison.CurrentCultureIgnoreCase))
            {
                if (table.DescriptionIdx != -1 && row.RowData.Count > table.DescriptionIdx)
                {
                    string[] descriptStr = row.RowData[table.DescriptionIdx].Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string descript in descriptStr)
                    {
                        if (descript.Contains("[") && descript.Contains("]"))
                        {
                            if (!Regex.IsMatch(descript, regex, RegexOptions.IgnoreCase) && descript.IndexOf("one_complement_target", StringComparison.OrdinalIgnoreCase) < 0)
                            {
                                ErrorReportManager.AddError(EFuseErrorType.E_InvalidCRC_01, row.SheetName, row.RowNum, table.DescriptionIdx + 1, []);
                            }
                        }
                    }
                }

                if (table.DefaultValueIdx != -1 && row.RowData.Count > table.DefaultValueIdx)
                {
                    string[] defaultStr = row.RowData[table.DefaultValueIdx].Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string descript in defaultStr)
                    {
                        if (descript.Contains("[") && descript.Contains("]"))
                        {
                            if (!Regex.IsMatch(descript, regex, RegexOptions.IgnoreCase) && descript.IndexOf("one_complement_target", StringComparison.OrdinalIgnoreCase) < 0)
                            {
                                ErrorReportManager.AddError(EFuseErrorType.E_InvalidCRC_01, row.SheetName, row.RowNum, table.DefaultValueIdx + 1, []);
                            }
                        }
                    }
                }
            }
        }

        internal void CheckCmpBitWidthInUdr(BitDefTable table, BitDefRow row, List<BitDefTable> udrItems)
        {
            var udrRows = new List<BitDefRow>();
            foreach (BitDefTable udr in udrItems)
            {
                udrRows.AddRange(udr.Rows.Where(item => item.RowData[udr.BankEfuseBitDefIdx].Equals(row.RowData[table.BankEfuseBitDefIdx], StringComparison.CurrentCultureIgnoreCase)));
            }

            if (!udrRows.TrueForAll(x => x.RowData[udrItems.First().BitWidthIdx].Equals(row.RowData[table.BitWidthIdx])))
            {
                string msg = $"The item {row.RowData[table.BankEfuseBitDefIdx]}'s Bitwidth of the CMP block is not same as UDR Block";
                ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationCMP_01, row.SheetName, row.RowNum, table.BitWidthIdx + 1, [row.RowData[table.BankEfuseBitDefIdx]]);
            }
        }

        internal void CheckDefualtValue(BitDefTable table, BitDefRow row)
        {
            if (table.DefaultValueIdx != -1 && row.RowData.Count > table.DefaultValueIdx &&
                table.BitWidthIdx != -1 && row.RowData.Count > table.BitWidthIdx &&
                row.RowData[table.BitWidthIdx].IsNumber())
            {
                decimal bit = Convert.ToDecimal(row.RowData[table.BitWidthIdx]);
                if (Regex.IsMatch(row.RowData[table.DefaultValueIdx], "^(x|0x|h)", RegexOptions.IgnoreCase))
                {
                    string defaultValue = Regex.Replace(row.RowData[table.DefaultValueIdx], "^(x|0x|h)", "", RegexOptions.IgnoreCase);
                    defaultValue = defaultValue.Replace("_", "");
                    decimal bitSize = Math.Ceiling(bit / 4);
                    if (bitSize < defaultValue.Length)
                    {
                        //string message = $"Incorrect bit width in the column Default Value. Bit Width (Math.Ceiling({bit}/4) = {bitSize}) < Default value ({defaultValue.Length})";
                        ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationBit_05, row.SheetName, row.RowNum, table.DefaultValueIdx + 1, [bit.ToString(), bitSize.ToString(), defaultValue.Length.ToString()]);
                    }
                }
                else if (Regex.IsMatch(row.RowData[table.DefaultValueIdx], "^(b|0b)", RegexOptions.IgnoreCase))
                {
                    string defaultValue = Regex.Replace(row.RowData[table.DefaultValueIdx], "^(b|0b)", "", RegexOptions.IgnoreCase);
                    defaultValue = defaultValue.Replace("_", "");
                    if (bit < defaultValue.Length)
                    {
                        //string message = $"Incorrect bit width in the column Default Value. Bit Width ({bit}) < Default value ({defaultValue.Length})";
                        ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationBit_06, row.SheetName, row.RowNum, table.DefaultValueIdx + 1, [bit.ToString(), defaultValue.Length.ToString()]);
                    }
                }
            }
        }
    }
}
