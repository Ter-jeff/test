using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Business;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using TestPlanLib.Efuse.Input;

namespace Automation.GenerateIgxl.EFuse.InputChecker
{
    public class EfuseConfigureBdfChecker
    {
        internal List<EfuseCrcItem> _ignorBitsList = new List<EfuseCrcItem>();

        public void WorkFlow(EfuseConfigMainSheet sheet, List<BitDefTable> tables, List<EfusePatternRow> efusePatternRows = null)
        {
            BitDefTable cfgTable = tables.Find(x => EFuseConst.GetBankName(x.BlockName).Equals(BankType.Cfg));
            CheckDefaultValue(sheet, cfgTable);
            CheckCrcItems(sheet, tables);
            CheckPatternMode(efusePatternRows, tables);
        }

        internal void CheckPatternMode(List<EfusePatternRow> efusePatternRows, List<BitDefTable> tables)
        {
            if (efusePatternRows != null && efusePatternRows.Any() && tables.Any())
            {
                var dicPatternsByBank = efusePatternRows.GroupBy(x => x.BankName).ToDictionary(x => x.Key, x => x.ToList());
                foreach (KeyValuePair<string, List<EfusePatternRow>> bank in dicPatternsByBank)
                {
                    if (string.IsNullOrEmpty(bank.Key))
                    {
                        continue;
                    }

                    EfusePatternRow firstItem = bank.Value.FirstOrDefault();
                    if (firstItem == null)
                    {
                        continue;
                    }

                    string type = firstItem.PayloadList.LastOrDefault().Split('_')[7];
                    BitDefTable bankBdfTable = tables.FirstOrDefault(x => EFuseConst.GetBankName(x.BlockName).Equals(bank.Key, StringComparison.CurrentCultureIgnoreCase));

                    if (bankBdfTable != null && !CheckPattenTypeWithBdf(type, bankBdfTable))
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_NotMatchToBDF_01, bankBdfTable.SheetName, bankBdfTable.HeaderRowNum - 1, 1, [type, bankBdfTable.AccessMode]);
                    }
                }
            }
        }

        internal bool CheckPattenTypeWithBdf(string currentType, BitDefTable bdfCfgTable)
        {
            var regDaa = new Regex(@"^Direct-Access\sMode");
            var regJtg = new Regex(@"^JTAG-Access\sMode");
            var endReg = new Regex(@"data\)");

            if (string.IsNullOrEmpty(currentType))
            {
                return false;
            }

            if (currentType.ToUpper().Equals("JTG"))
            {
                if (!regJtg.IsMatch(bdfCfgTable.AccessMode) || !endReg.IsMatch(bdfCfgTable.AccessMode))
                {
                    return false;
                }
            }
            else if (currentType.ToUpper().Equals("DAA"))
            {
                if (!regDaa.IsMatch(bdfCfgTable.AccessMode) || !endReg.IsMatch(bdfCfgTable.AccessMode))
                {
                    return false;
                }
            }

            return true;
        }
        internal void CheckDefaultValue(EfuseConfigMainSheet cfgSheet, BitDefTable bdfCfgTable)
        {
            foreach (BitDefRow bdfRow in bdfCfgTable.Rows)
            {
                if (!bdfRow.RowData[bdfCfgTable.AlgorithmIdx].Equals("cond", StringComparison.CurrentCultureIgnoreCase))
                {
                    string registerName = bdfRow.RowData[bdfCfgTable.BankEfuseBitDefIdx].Replace("bank_config_", "");
                    EfuseConfigMainRow cfgData = cfgSheet.Rows.Find(x => x.Description.Equals(registerName, StringComparison.CurrentCultureIgnoreCase));
                    if (cfgData == null)
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_MissingField_01, bdfCfgTable.SheetName, bdfRow.RowNum, bdfCfgTable.DefaultValueIdx + 1, [registerName]);
                    }
                }
            }
        }

        internal void CheckCrcItems(EfuseConfigMainSheet cfgSheet, List<BitDefTable> tables)
        {
            foreach (BitDefTable table in tables)
            {
                CheckCrcIgnorBit(table);
                foreach (BitDefRow row in table.Rows)
                {
                    if (row.RowData[table.AlgorithmIdx].Equals("cond", StringComparison.CurrentCultureIgnoreCase))
                    {
                        double msb = Convert.ToDouble(row.RowData[table.MsbBitIdx]);
                        double lsb = Convert.ToDouble(row.RowData[table.LsbBitIdx]);
                        EfuseConfigMainSheet firstConfigSheet = cfgSheet;
                        var configRows = firstConfigSheet.Rows.Where(x => x.Lsb >= lsb && x.Msb <= msb).ToList();
                        foreach (EfuseConfigMainRow cfgRow in configRows)
                        {
                            CheckCfgJob(cfgSheet, cfgRow);
                        }
                    }
                    else
                    {
                        CheckJob(table, row);
                    }
                }
            }
        }

        internal void CheckCrcIgnorBit(BitDefTable table)
        {
            //check if crc item bits range exist in ignor bits
            _ignorBitsList = new List<EfuseCrcItem>();
            var crcItems = table.Rows.Where(x => x.RowData[table.AlgorithmIdx].Equals("crc", StringComparison.CurrentCultureIgnoreCase)).ToList();
            crcItems = crcItems.Where(x => x.Line.IndexOf("one_complement_target", StringComparison.OrdinalIgnoreCase) < 0).ToList();
            int offset = int.Parse(table.Rows.First().RowData[table.LsbBitIdx]);
            foreach (BitDefRow crcItem in crcItems)
            {
                bool isCalcBits = crcItem.RowData[table.DescriptionIdx].ContainsIgnoreCase("crc_calcbits");
                List<int> ignoreBits = GetIgnoreBitRange(crcItem.RowData[table.DescriptionIdx]);
                if (!ignoreBits.Any())
                {
                    continue;
                }

                int lsb = int.Parse(crcItem.RowData[table.LsbBitIdx]) - offset;
                int msb = int.Parse(crcItem.RowData[table.MsbBitIdx]) - offset;
                if (lsb > msb)
                {
                    lsb = int.Parse(crcItem.RowData[table.MsbBitIdx]);
                    msb = int.Parse(crcItem.RowData[table.LsbBitIdx]);
                }
                for (int i = lsb; i <= msb; i++)
                {
                    if (isCalcBits)
                    {
                        if (ignoreBits.Exists(x => x == i))
                        {
                            string message = $"Item {crcItem.RowData[table.BankEfuseBitDefIdx]} bit range {msb}:{lsb} exist itself";
                            ErrorReportManager.AddError(EFuseErrorType.E_InvalidCRC_02, crcItem.SheetName, crcItem.RowNum, 0, [crcItem.RowData[table.BankEfuseBitDefIdx], msb.ToString(), lsb.ToString()]);
                            break;
                        }
                    }
                    else
                    {
                        if (!ignoreBits.Exists(x => x == i))
                        {
                            string message = $"Item {crcItem.RowData[table.BankEfuseBitDefIdx]} bit range {msb}:{lsb} does not exist in ignore bit";
                            ErrorReportManager.AddError(EFuseErrorType.E_InvalidCRC_03, crcItem.SheetName, crcItem.RowNum, 0, [crcItem.RowData[table.BankEfuseBitDefIdx], msb.ToString(), lsb.ToString()]);
                            break;
                        }
                    }
                }
                var item = new EfuseCrcItem
                {
                    ItemName = crcItem.RowData[table.BankEfuseBitDefIdx],
                    IngorBits = ignoreBits,
                    Job = crcItem.RowData[table.ProgrammingStageIdx],
                    IsCalcBits = isCalcBits,
                    Offset = offset
                };
                _ignorBitsList.Add(item);
            }
        }

        internal void CheckCfgJob(EfuseConfigMainSheet table, EfuseConfigMainRow row)
        {
            EfuseBitDefTableJobSequence currJobseq = GetJobSequence(row.FuseBlowLocation);
            var lsbmsb = new EFuseLsbMsb();
            lsbmsb.SetLsbmsbData(row.Fuse);
            int lsb = int.Parse(lsbmsb.GetLsb());
            int msb = int.Parse(lsbmsb.GetMsb());
            foreach (EfuseCrcItem crcItem in _ignorBitsList)
            {
                EfuseBitDefTableJobSequence crcJobSeq = GetJobSequence(crcItem.Job);
                for (int i = lsb - crcItem.Offset; i <= msb - crcItem.Offset; i++)
                {
                    if (crcItem.IsCalcBits)
                    {

                        if (crcItem.IngorBits.Exists(x => x == i) && currJobseq > crcJobSeq)
                        {
                            ErrorReportManager.AddError(EFuseErrorType.E_InvalidCRC_04, row.SheetName, row.RowNum, table.FuseBlowLocationColNumber, [lsb.ToString(), msb.ToString(), row.FuseBlowLocation, crcItem.ItemName, crcItem.Job]);
                            break;
                        }
                    }
                    else
                    {
                        if (!crcItem.IngorBits.Exists(x => x == i) && currJobseq > crcJobSeq)
                        {
                            ErrorReportManager.AddError(EFuseErrorType.E_InvalidCRC_05, row.SheetName, row.RowNum, table.FuseBlowLocationColNumber, [lsb.ToString(), msb.ToString(), row.FuseBlowLocation, crcItem.ItemName, crcItem.Job]);
                            break;
                        }
                    }
                }
            }

        }

        internal void CheckJob(BitDefTable table, BitDefRow row)
        {
            EfuseBitDefTableJobSequence currJobseq = GetJobSequence(row.RowData[table.ProgrammingStageIdx]);
            if (!int.TryParse(row.RowData[table.LsbBitIdx], out int lsb))
            {
                ErrorReportManager.AddError(EFuseErrorType.E_InvalidLsbMsb_01, row.SheetName, row.RowNum, table.LsbBitIdx + 1, []);
            }

            if (!int.TryParse(row.RowData[table.MsbBitIdx], out int msb))
            {
                ErrorReportManager.AddError(EFuseErrorType.E_InvalidLsbMsb_02, row.SheetName, row.RowNum, table.MsbBitIdx + 1, []);
            }

            foreach (EfuseCrcItem crcItem in _ignorBitsList)
            {
                EfuseBitDefTableJobSequence crcJobSeq = GetJobSequence(crcItem.Job);
                for (int i = lsb - crcItem.Offset; i <= msb - crcItem.Offset; i++)
                {
                    if (crcItem.IsCalcBits)
                    {
                        if (crcItem.IngorBits.Exists(x => x == i) && currJobseq > crcJobSeq)
                        {
                            ErrorReportManager.AddError(EFuseErrorType.E_InvalidCRC_06, row.SheetName, row.RowNum, table.ProgrammingStageIdx + 1, [row.RowData[table.BankEfuseBitDefIdx], msb.ToString(), lsb.ToString(), row.RowData[table.ProgrammingStageIdx], crcItem.ItemName, crcItem.Job]);
                            break;
                        }
                    }
                    else
                    {
                        if (!crcItem.IngorBits.Exists(x => x == i) && currJobseq > crcJobSeq)
                        {
                            ErrorReportManager.AddError(EFuseErrorType.E_InvalidCRC_07, row.SheetName, row.RowNum, table.ProgrammingStageIdx + 1, [row.RowData[table.BankEfuseBitDefIdx], msb.ToString(), lsb.ToString(), row.RowData[table.ProgrammingStageIdx], crcItem.ItemName, crcItem.Job]);
                            break;
                        }
                    }
                }
            }
        }

        internal List<int> GetIgnoreBitRange(string ignoreStr)
        {
            var rangeList = new List<int>();
            ignoreStr = ignoreStr.Replace("_", "");
            ignoreStr = Regex.Replace(ignoreStr, "[a-zA-Z]*", "");
            string[] ignoreList = ignoreStr.Split(new[] { ',', ' ', '=' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string range in ignoreList)
            {
                var lsbmsb = new EFuseLsbMsb();
                lsbmsb.SetLsbmsbData(range);
                int lsb = int.Parse(lsbmsb.GetLsb());
                int msb = int.Parse(lsbmsb.GetMsb());
                for (int i = lsb; i <= msb; i++)
                {
                    rangeList.Add(i);
                }
            }

            return rangeList;
        }

        internal EfuseBitDefTableJobSequence GetJobSequence(string job)
        {
            if (job.ContainsIgnoreCase("CP1"))
            {
                return EfuseBitDefTableJobSequence.Cp1;
            }

            if (job.ContainsIgnoreCase("CP2"))
            {
                return EfuseBitDefTableJobSequence.Cp2;
            }

            if (job.ContainsIgnoreCase("FT1"))
            {
                return EfuseBitDefTableJobSequence.Ft1;
            }

            if (job.ContainsIgnoreCase("FT2"))
            {
                return EfuseBitDefTableJobSequence.Ft2;
            }

            if (job.ContainsIgnoreCase("FT3") || job.ContainsIgnoreCase("FTF"))
            {
                return EfuseBitDefTableJobSequence.Ft3;
            }

            return EfuseBitDefTableJobSequence.Unknown;
        }
    }

    public class EfuseCrcItem
    {
        public string ItemName = "";
        public string Job = "";
        public List<int> IngorBits = new List<int>();
        public bool IsCalcBits;
        public int Offset;
    }

    public enum EfuseBitDefTableJobSequence
    {
        Cp1, Cp2, Ft1, Ft2, Ft3, Unknown
    }
}
