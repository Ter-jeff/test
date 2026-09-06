using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Business;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using ScghLib.Reader;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Efuse.Input;

namespace Automation.GenerateIgxl.EFuse.InputChecker
{
    public class EfusePatternMapChecker
    {
        private readonly Dictionary<string, PatternData> _patternDatas;
        private readonly List<BitDefBankRange> _bankRanges;
        private readonly EfusePatternType _patternType;
        private readonly EfusePatternRow _patternRow;

        public EfusePatternMapChecker() { }

        public EfusePatternMapChecker(Dictionary<string, PatternData> patternDatas, EfusePatternRow patternRow, List<BitDefBankRange> bankRanges)
        {
            _patternDatas = patternDatas;
            _bankRanges = bankRanges;
            _patternRow = patternRow;
            _patternType = patternRow.PatternType;
        }

        public void CheckScghPatterns(ProdCharSheetRow hardIpRow, string bdfAccessName, string bdfBitMode)
        {
            CheckPatternInScgh(hardIpRow);
            CheckWritePatternSendCount(hardIpRow.SourceSheetName, hardIpRow.RowNum, hardIpRow.PayloadValue, bdfBitMode);
            CheckReadPatternStoreCount(hardIpRow.SourceSheetName, hardIpRow.RowNum, hardIpRow.PayloadValue, bdfAccessName, bdfBitMode);
        }

        public void CheckInstancePatterns(BinCutInstanceRow instanceRow, string bdfAccessName, string bdfBitMode)
        {
            CheckPatternInInstance(instanceRow);
            CheckWritePatternSendCount(instanceRow.SheetName, instanceRow.RowNum, instanceRow.PayloadValue, bdfBitMode);
            CheckReadPatternStoreCount(instanceRow.SheetName, instanceRow.RowNum, instanceRow.PayloadValue, bdfAccessName, bdfBitMode);
        }

        private void CheckPatternInScgh(ProdCharSheetRow hardIpRow)
        {
            if (_patternDatas == null || !_patternDatas.Any())
            {
                return;
            }

            #region check payload
            if (!string.IsNullOrEmpty(hardIpRow.PayloadValue))
            {
                if (!_patternDatas.ContainsKey(hardIpRow.PayloadValue.ToLower()))
                {
                    ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_02, hardIpRow.SourceSheetName, hardIpRow.RowNum, 0, [hardIpRow.PayloadValue]);
                }
                else
                {
                    #region check if pattern in CSV
                    //ErrorLevel errorLevel = ErrorLevel.Warning;
                    //if (_patternRow.ReadOrWrite != EfusePatternWorR.Non)
                    //{
                    //    errorLevel = ErrorLevel.Error;
                    //}

                    if (_patternDatas[hardIpRow.PayloadValue.ToLower()].Use.Equals("dont_use", StringComparison.OrdinalIgnoreCase))
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_03, hardIpRow.SourceSheetName, hardIpRow.RowNum, 0, [hardIpRow.PayloadValue]);
                    }
                    else
                    {
                        if (_patternDatas[hardIpRow.PayloadValue.ToLower()].FileVersion.Equals("n/a", StringComparison.OrdinalIgnoreCase))
                        {
                            ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_04, hardIpRow.SourceSheetName, hardIpRow.RowNum, 0, [hardIpRow.PayloadValue]);
                        }
                    }
                    #endregion

                    #region check pattern type
                    if (_patternType.TestMode.Equals(EfuseTestMode.Unknow))
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_MismatchPattern_05, hardIpRow.SourceSheetName, hardIpRow.RowNum, 0, [hardIpRow.PayloadValue]);
                    }
                    #endregion

                    #region check read write pin
                    if (_patternRow.ReadOrWrite != EfusePatternWorR.Non && string.IsNullOrEmpty(_patternRow.ReadWritePin))
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_05, hardIpRow.SourceSheetName, hardIpRow.RowNum, 0, [hardIpRow.PayloadValue]);
                    }
                    #endregion
                }
            }
            #endregion

            #region check init
            for (int i = 0; i < hardIpRow.InitList.Count; i++)
            {
                if (string.IsNullOrEmpty(hardIpRow.InitList[i]))
                {
                    continue;
                }

                string init = hardIpRow.InitList[i];
                if (!_patternDatas.ContainsKey(init.ToLower()))
                {
                    ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_06, hardIpRow.SourceSheetName, hardIpRow.RowNum, 0, [init]);
                }
                else
                {
                    if (_patternDatas[init.ToLower()].Use.Equals("dont_use", StringComparison.OrdinalIgnoreCase))
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_03, hardIpRow.SourceSheetName, hardIpRow.RowNum, 0, [init]);
                    }
                    else
                    {
                        if (_patternDatas[init.ToLower()].FileVersion.Equals("n/a", StringComparison.OrdinalIgnoreCase))
                        {
                            ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_04, hardIpRow.SourceSheetName, hardIpRow.RowNum, 0, [init]);
                        }
                    }
                }
            }
            #endregion
        }

        internal void CheckPatternInInstance(BinCutInstanceRow instanceRow)
        {
            if (_patternDatas == null || !_patternDatas.Any())
            {
                return;
            }

            #region check payload
            if (!string.IsNullOrEmpty(instanceRow.PayloadValue))
            {
                if (!_patternDatas.ContainsKey(instanceRow.PayloadValue.ToLower()))
                {
                    ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_02, instanceRow.SheetName, instanceRow.RowNum, 0, [instanceRow.PayloadValue]);
                }
                else
                {
                    #region check if pattern in CSV
                    //ErrorLevel errorLevel = ErrorLevel.Warning;
                    //if (_patternRow.ReadOrWrite != EfusePatternWorR.Non)
                    //{
                    //    errorLevel = ErrorLevel.Error;
                    //}

                    if (_patternDatas[instanceRow.PayloadValue.ToLower()].Use.Equals("dont_use", StringComparison.OrdinalIgnoreCase))
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_03, instanceRow.SheetName, instanceRow.RowNum, 0, [instanceRow.PayloadValue]);
                    }
                    else
                    {
                        if (_patternDatas[instanceRow.PayloadValue.ToLower()].FileVersion.Equals("n/a", StringComparison.OrdinalIgnoreCase))
                        {
                            //string errorMessage = $"This pattern {instanceRow.PayloadValue} of the FileVersion column is \"n/a\" !!!";
                            ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_04, instanceRow.SheetName, instanceRow.RowNum, 0, [instanceRow.PayloadValue]);
                        }
                    }
                    #endregion

                    #region check pattern type
                    if (_patternType.TestMode.Equals(EfuseTestMode.Unknow))
                    {
                        //string errorMessage = $"This pattern {instanceRow.PayloadValue} can't be discerned in what mode !!!";
                        ErrorReportManager.AddError(EFuseErrorType.E_MismatchPattern_05, instanceRow.SheetName, instanceRow.RowNum, 0, [instanceRow.PayloadValue]);
                    }
                    #endregion

                    #region check read write pin
                    if (_patternRow.ReadOrWrite != EfusePatternWorR.Non && string.IsNullOrEmpty(_patternRow.ReadWritePin))
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_05, instanceRow.SheetName, instanceRow.RowNum, 0, [instanceRow.PayloadValue]);
                    }
                    #endregion
                }
            }
            #endregion

            #region check init
            for (int i = 0; i < instanceRow.InitList.Count; i++)
            {
                if (string.IsNullOrEmpty(instanceRow.InitList[i]))
                {
                    continue;
                }

                string init = instanceRow.InitList[i];
                if (!_patternDatas.ContainsKey(init.ToLower()))
                {
                    ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_02, instanceRow.SheetName, instanceRow.RowNum, 0, [init]);
                }
                else
                {
                    if (_patternDatas[init.ToLower()].Use.Equals("dont_use", StringComparison.OrdinalIgnoreCase))
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_03, instanceRow.SheetName, instanceRow.RowNum, 0, [init]);
                    }
                    else
                    {
                        if (_patternDatas[init.ToLower()].FileVersion.Equals("n/a", StringComparison.OrdinalIgnoreCase))
                        {
                            ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_04, instanceRow.SheetName, instanceRow.RowNum, 0, [init]);
                        }
                    }
                }
            }
            #endregion
        }

        private void CheckWritePatternSendCount(string sheetName, int rowNumber, string pattern, string bdfBitMode)
        {
            if (_patternRow.ReadOrWrite == EfusePatternWorR.Write)
            {
                BitDefBankRange bankRange = _bankRanges.Find(x => x.BankName.Equals(_patternRow.BankName, StringComparison.CurrentCultureIgnoreCase));
                if (bankRange != null)
                {
                    int mode = bdfBitMode.Contains("2") ? 2 : 1;
                    double bits = bankRange.Width * mode;
                    switch (_patternType.TestMode)
                    {
                        case EfuseTestMode.WriteDsscHf:
                            bits /= 2;
                            break;
                        case EfuseTestMode.WriteDsscQ:
                            bits /= 4;
                            break;
                        case EfuseTestMode.WriteDsscH:
                            bits /= 16;
                            break;
                    }
                    if (_patternRow.SendStoreCount != bits)
                    {
                        ErrorReportManager.AddError(EFuseErrorType.E_MismatchPattern_01, sheetName, rowNumber, 0, [pattern, _patternRow.SendStoreCount.ToString(), bits.ToString(), bankRange.BankName, bankRange.Width.ToString(), mode.ToString()]);
                    }
                }
            }
        }

        private void CheckReadPatternStoreCount(string sheetName, int rowNumber, string pattern, string bdfAccessName, string bdfBitMode)
        {
            if (_patternRow.ReadOrWrite == EfusePatternWorR.Read)
            {
                BitDefBankRange bankRange = _bankRanges.Find(x => x.BankName.Equals(_patternRow.BankName, StringComparison.CurrentCultureIgnoreCase));
                if (bankRange.BankName.StartsWith("UDR") && _patternRow.PatternType.TestMode.Equals(EfuseTestMode.Ver2))
                {
                    bankRange =
                        _bankRanges.Find(
                            x => x.BankName.Equals(_patternRow.BankName.Contains("P") ? BankType.CmpP : BankType.CmpE,
                            StringComparison.CurrentCultureIgnoreCase));
                }
                if (bankRange != null)
                {
                    int mode = bdfBitMode.Contains("2") ? 2 : 1;
                    double bits = bankRange.Width * mode;
                    string[] patStr = pattern.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                    if (bdfAccessName.Contains("Direct-Access Mode"))
                    {
                        bits /= 32;
                    }

                    if (patStr.Length > 7)
                    {
                        if (patStr[7].Equals("DAA", StringComparison.CurrentCultureIgnoreCase) &&
                            !bdfAccessName.Contains("Direct-Access Mode"))
                        {
                            string errorMessage =
                                $"The pattern {pattern} access mode (DAA) is not equal to BDF file {bdfAccessName}";
                            ErrorReportManager.AddError(EFuseErrorType.E_MismatchPattern_02, sheetName, rowNumber, 0, [pattern, bdfAccessName]);
                        }
                        if (patStr[7].Equals("JTG", StringComparison.CurrentCultureIgnoreCase) &&
                            !bdfAccessName.Contains("JTAG-Access Mode"))
                        {
                            string errorMessage =
                                $"The pattern {pattern} access mode (JTG) is not equal to BDF file {bdfAccessName}";
                            ErrorReportManager.AddError(EFuseErrorType.E_MismatchPattern_03, sheetName, rowNumber, 0, [pattern, bdfAccessName]);
                        }

                        if (_patternRow.SendStoreCount != bits)
                        {
                            string errorMessage;
                            if (bdfAccessName.Contains("Direct-Access Mode"))
                            {
                                errorMessage =
                                    $"The pattern {pattern} ({patStr[7].ToUpper()}) store count {_patternRow.SendStoreCount} is not equal to {bits} in the bank {bankRange.BankName} ({bdfAccessName}), {bankRange.Width}*{mode}/32 !!!";
                            }
                            else
                            {
                                errorMessage =
                                    $"The pattern {pattern} ({patStr[7].ToUpper()}) store count {_patternRow.SendStoreCount} is not equal to {bits} in the bank {bankRange.BankName} ({bdfAccessName}), {bankRange.Width}*{mode} !!!";
                            }

                            ErrorReportManager.AddError(EFuseErrorType.E_MismatchPattern_04, sheetName, rowNumber, 0, [pattern, patStr[7].ToUpper(), _patternRow.SendStoreCount.ToString(), bits.ToString(), bankRange.BankName, bdfAccessName, bankRange.Width.ToString(), mode.ToString()]);
                        }
                    }
                }
            }
        }

        public void CheckUnusedPatInScgh(ScghData scghData, EfuseFinalInstanceRows finalRows, List<EfusePatternRow> patRows)
        {
            var unUsedPats = patRows.Where(pat => !finalRows.Exists(x => x.EfusePatternRow != null &&
                x.EfusePatternRow.PatternType.IsDvrv == pat.PatternType.IsDvrv &&
                x.EfusePatternRow.PatternType.PatJob.Equals(pat.PatternType.PatJob, StringComparison.CurrentCultureIgnoreCase) &&
                x.EfusePatternRow.PatternType.IsEarly == pat.PatternType.IsEarly &&
                x.EfusePatternRow.PatternType.IsFinal == pat.PatternType.IsFinal &&
                x.EfusePatternRow.PatternType.DvrvType == pat.PatternType.DvrvType &&
                x.EfusePatternRow.PatternType.TestMode.Equals(pat.PatternType.TestMode))).ToList();
            var efuseScghRowList = new List<ProdCharSheetRow>();
            if (scghData.HardIpSheetRowList.Any())
            {
                efuseScghRowList.AddRange(scghData.HardIpSheetRowList);
            }

            if (scghData.ScanScghRowList.Any())
            {
                efuseScghRowList.AddRange(scghData.ScanScghRowList);
            }

            foreach (EfusePatternRow pat in unUsedPats)
            {
                if (pat.PatternType.TestMode.Equals(EfuseTestMode.Unknow) && !pat.PatternType.IsDvrv)
                {
                    continue;
                }

                IEnumerable<ProdCharSheetRow> scghSheet = efuseScghRowList.Where(x => x.SourceSheetName.Equals(pat.SheetName) && x.RowNum == pat.RowNum);
                EfuseTestMode testMode = pat.PatternType.TestMode;
                foreach (ProdCharSheetRow row in scghSheet)
                {
                    ErrorReportManager.AddError(EFuseErrorType.W_RedundantPattern_01, row.SourceSheetName, row.RowNum, 0, [row.PayloadValue, pat.BankName, testMode.ToString()]);
                }
            }
        }
    }
}
