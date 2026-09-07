using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.InputChecker;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using ScghLib.Reader;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Efuse.Input;

namespace Automation.GenerateIgxl.EFuse.Business
{
    public class EfusePatternClassification
    {
        private readonly bool _onlyUdrP;
        private readonly List<BitDefBankRange> _bitDefBankRanges;

        public EfusePatternClassification(bool onlyUdrP, List<BitDefBankRange> bitDefBankRanges)
        {
            _onlyUdrP = onlyUdrP;
            _bitDefBankRanges = bitDefBankRanges;
        }

        public List<EfusePatternRow> ClassificationFromScgh(ScghData scghData, Dictionary<string, PatternData> patternDatas, HashSet<string> patInK, List<BitDefTable> efuseBitDefTables)
        {
            var efusePatRows = new List<EfusePatternRow>();
            var efuseScghRowList = new List<ProdCharSheetRow>();
            if (scghData.HardIpSheetRowList.Any())
            {
                efuseScghRowList.AddRange(scghData.HardIpSheetRowList);
            }

            if (scghData.ScanScghRowList.Any())
            {
                efuseScghRowList.AddRange(scghData.ScanScghRowList);
            }

            foreach (ProdCharSheetRow data in efuseScghRowList)
            {
                var efusePatRow = new EfusePatternRow();
                string[] payloadStr = data.PayloadValue.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                efusePatRow.BankName = EFuseConst.GetBankName(payloadStr[6], payloadStr[5]);
                if (efusePatRow.BankName.Equals(BankType.UdrE) || efusePatRow.BankName.Equals(BankType.UdrP) || efusePatRow.BankName.Equals(BankType.UdrM))
                {
                    efusePatRow.BankName = GetDetailBankName(data.InitList, payloadStr);
                }

                efusePatRow.PatternType = GetPatternType(payloadStr, efusePatRow.BankName.Equals(BankType.Cfg));
                bool isRead = PatternTypeIsRead(efusePatRow.PatternType);
                bool isWrite = PatternTypeIsWrite(efusePatRow.PatternType);

                if (isRead || isWrite)
                {
                    var findPatternPath = patternDatas.Where(x => x.Key.Equals(data.PayloadValue, StringComparison.CurrentCultureIgnoreCase) && x.Value.IsExist).ToList();
                    string patternPath = "";
                    if (findPatternPath.Any())
                    {
                        patternPath = findPatternPath.First().Value.PatternVersion;
                    }

                    if (!string.IsNullOrEmpty(patternPath))
                    {
                        patternPath += ".atp.gz";
                        string file = patInK?.FirstOrDefault(p => string.Equals(Path.GetFileName(p), patternPath, StringComparison.CurrentCultureIgnoreCase));
                        if (file != null)
                        {
                            var patternParser = new EfusePatternParser(file, isRead, isWrite);
                            patternParser.WorkFlow();
                            efusePatRow.ReadWritePin = patternParser.ReadWritePin;
                            efusePatRow.SendStoreCount = isWrite ? patternParser.SendCount : patternParser.StoreCount;
                            efusePatRow.ReadOrWrite = isWrite ? EfusePatternWorR.Write : EfusePatternWorR.Read;
                        }
                        else
                        {
                            ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_01, data.SourceSheetName, data.RowNum, 0, [patternPath]);
                        }
                    }
                }

                efusePatRow.InitList = data.InitList;
                efusePatRow.PayloadList = data.PayloadList;
                efusePatRow.PatList = new List<string>();
                efusePatRow.PatList.AddRange(data.InitList);
                efusePatRow.PatList.AddRange(data.PayloadList);
                efusePatRow.SheetName = data.SourceSheetName;
                efusePatRow.RowNum = data.RowNum;
                efusePatRows.Add(efusePatRow);
                string bdfAccessName = efuseBitDefTables.Find(x => EFuseConst.GetBankName(x.BlockName).Equals(efusePatRow.BankName)).AccessMode;
                string bdfBitMode = efuseBitDefTables.Find(x => EFuseConst.GetBankName(x.BlockName).Equals(efusePatRow.BankName)).BitMode;
                var patternChecker = new EfusePatternMapChecker(patternDatas, efusePatRow, _bitDefBankRanges);
                patternChecker.CheckScghPatterns(data, bdfAccessName, bdfBitMode);
            }
            return efusePatRows;
        }

        public List<EfusePatternRow> ClassificationFromInstanceSheet(BinCutInstanceSheet instanceData, Dictionary<string, PatternData> patternDatas, HashSet<string> patInK, List<BitDefTable> efuseBitDefTables)
        {
            var efusePatRows = new List<EfusePatternRow>();
            foreach (BinCutInstanceRow data in instanceData.Rows)
            {
                bool patFound = false;
                var efusePatRow = new EfusePatternRow();
                if (!data.InitList.Any() && !data.PayloadList.Any())
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(data.PayloadValue))
                {
                    string[] payloadStr = data.PayloadValue.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                    efusePatRow.BankName = EFuseConst.GetBankName(payloadStr[6], payloadStr[5]);
                    if (efusePatRow.BankName.Equals(BankType.UdrE) || efusePatRow.BankName.Equals(BankType.UdrP) || efusePatRow.BankName.Equals(BankType.UdrM))
                    {
                        efusePatRow.BankName = GetDetailBankName(data.InitList, payloadStr);
                    }

                    efusePatRow.PatternType = GetPatternType(payloadStr, efusePatRow.BankName.Equals(BankType.Cfg));
                    bool isRead = PatternTypeIsRead(efusePatRow.PatternType);
                    bool isWrite = PatternTypeIsWrite(efusePatRow.PatternType);

                    if (isRead || isWrite)
                    {
                        var findPatternPath = patternDatas.Where(x => x.Key.Equals(data.PayloadValue, StringComparison.CurrentCultureIgnoreCase) && x.Value.IsExist).ToList();
                        string patternPath = "";
                        if (findPatternPath.Any())
                        {
                            patternPath = findPatternPath.First().Value.PatternVersion;
                        }

                        if (!string.IsNullOrEmpty(patternPath) && patInK != null)
                        {
                            patternPath += ".atp.gz";
                            string file = patInK?.FirstOrDefault(p => string.Equals(Path.GetFileName(p), patternPath, StringComparison.CurrentCultureIgnoreCase));
                            if (file != null)
                            {
                                var patternParser = new EfusePatternParser(file, isRead, isWrite);
                                patternParser.WorkFlow();
                                efusePatRow.ReadWritePin = patternParser.ReadWritePin;
                                efusePatRow.SendStoreCount = isWrite ? patternParser.SendCount : patternParser.StoreCount;
                                efusePatRow.ReadOrWrite = isWrite ? EfusePatternWorR.Write : EfusePatternWorR.Read;
                            }
                            else
                            {
                                ErrorReportManager.AddError(EFuseErrorType.E_MissingPattern_01, data.SheetName, data.RowNum, 0, [patternPath]);
                            }
                        }
                    }
                    if (patFound)
                    {
                        string bdfAccessName = efuseBitDefTables.Find(x => EFuseConst.GetBankName(x.BlockName).Equals(efusePatRow.BankName)).AccessMode;
                        string bdfBitMode = efuseBitDefTables.Find(x => EFuseConst.GetBankName(x.BlockName).Equals(efusePatRow.BankName)).BitMode;
                        var patternChecker = new EfusePatternMapChecker(patternDatas, efusePatRow, _bitDefBankRanges);
                        patternChecker.CheckInstancePatterns(data, bdfAccessName, bdfBitMode);
                    }
                }

                efusePatRow.InitList = data.InitList;
                efusePatRow.PayloadList = data.PayloadList;
                efusePatRow.PatList = new List<string>();
                efusePatRow.PatList.AddRange(data.InitList);
                efusePatRow.PatList.AddRange(data.PayloadList);
                efusePatRow.SheetName = data.SheetName;
                efusePatRow.RowNum = data.RowNum;
                efusePatRows.Add(efusePatRow);
            }
            return efusePatRows;
        }

        private bool PatternTypeIsRead(EfusePatternType patternType)
        {
            return patternType.TestMode == EfuseTestMode.RvRead || patternType.TestMode == EfuseTestMode.BcRead ||
                   patternType.TestMode == EfuseTestMode.ApbRead || patternType.TestMode == EfuseTestMode.JtagRead ||
                   patternType.TestMode == EfuseTestMode.JtagReadDap ||
                   patternType.TestMode == EfuseTestMode.MarginRead ||
                   patternType.TestMode == EfuseTestMode.MarginReadFull ||
                   patternType.TestMode == EfuseTestMode.ReadDap ||
                   patternType.TestMode == EfuseTestMode.Uso || patternType.TestMode == EfuseTestMode.Ver2;
        }

        private bool PatternTypeIsWrite(EfusePatternType patternType)
        {
            return patternType.TestMode == EfuseTestMode.RvWrite ||
                   patternType.TestMode == EfuseTestMode.ApbWrite || patternType.TestMode == EfuseTestMode.JtagWrite ||
                   patternType.TestMode == EfuseTestMode.WriteDssc ||
                   patternType.TestMode == EfuseTestMode.WriteDsscQ ||
                   patternType.TestMode == EfuseTestMode.WriteDsscH ||
                   patternType.TestMode == EfuseTestMode.WriteDsscHf ||
                   patternType.TestMode == EfuseTestMode.Usi;
        }

        private string GetJobFromPayloadPattern(string[] payload)
        {
            string job = "";
            for (int i = payload.Length - 1; i > 7; i--)
            {
                Match regexCondition = Regex.Match(payload[i], @"(?<Job_Cp>CP\d*)|(?<Job_Ft>FTF?\d*)|(?<Job_Slt>SLT)", RegexOptions.IgnoreCase);
                if (regexCondition.Value.Any())
                {
                    job = regexCondition.Value.Equals("FTF", StringComparison.CurrentCultureIgnoreCase) ? "FT3" : regexCondition.Value.ToUpper();
                }
            }
            return job;
        }

        private class PatternValidationInput
        {
            public EfusePatternType PatternType { get; set; }

            public string[] Payload { get; set; }

            public int PatternCount { get; set; }
        }
        private const string BlankCheckRegex = @"(CP|FT|WLFT|SLT|FTF)\d?(bc|b)";
        private const string BlankCheckRegexNew = "blank";
        private const string EarlyBcRegex = @"(CP|FT|WLFT|SLT|FTF)\d?(ebc)";
        private const string DvrvRegex2 = "(dv|rv|default|fixed)";

        private static bool IsStringStartWithCpOrFt(string token)
        {
            return token.StartsWith("CP", StringComparison.CurrentCultureIgnoreCase)
                || token.StartsWith("FT", StringComparison.CurrentCultureIgnoreCase);
        }

        private static readonly (
            Func<PatternValidationInput, bool> ShouldApply,
            Action<PatternValidationInput> Update
        )[] _rules =
        {
            (
                i => i.Payload.Last().Equals("ZEROS", StringComparison.CurrentCultureIgnoreCase),
                i =>
                {
                    i.PatternType.IsEarly = true;
                    i.PatternType.DvrvType = DvRvType.FlatCheck;
                }
            ),
            (
                i => Regex.IsMatch(i.Payload.Last(), BlankCheckRegex, RegexOptions.IgnoreCase),
                i =>
                {
                    i.PatternType.DvrvType = DvRvType.FlatCheck;
                }
            ),
            (
                i => Regex.IsMatch(i.Payload.Last(), EarlyBcRegex, RegexOptions.IgnoreCase),
                i =>
                {
                    i.PatternType.IsEarly = true;
                    i.PatternType.DvrvType = DvRvType.FlatCheck;
                }
            ),
            (
                i => Regex.IsMatch(i.Payload.Last(), BlankCheckRegexNew, RegexOptions.IgnoreCase),
                OnBlankCheckRegexNewMatch
            ),
            (
                IsFldsscRvPatternForCpFt,
                i =>
                {
                    i.PatternType.DvrvType = DvRvType.Rv;
                }
            ),
            (IsCpOrFtPattern, OnCpOrFtMatch),
            (IsMatchDvrvRegex2, OnDvrvRegex2Match),
        };

        private static void OnBlankCheckRegexNewMatch(PatternValidationInput i)
        {
            i.PatternType.DvrvType = DvRvType.FlatCheck;
            if (i.PatternCount <= 12)
            {
                return;
            }
            string target = i.Payload[i.PatternCount - 2];

            bool isCpOrFt = IsStringStartWithCpOrFt(target);
            bool isE = target.EndsWith("E", StringComparison.CurrentCultureIgnoreCase);

            if (isCpOrFt && isE)
            {
                i.PatternType.IsEarly = true;
            }
        }

        private static bool IsFldsscRvPatternForCpFt(PatternValidationInput i)
        {
            if (i.PatternCount <= 12)
            {
                return false;
            }

            string headerToken = i.Payload[i.PatternCount - 2];
            string markerToken = i.Payload[12];

            bool isCpOrFt = IsStringStartWithCpOrFt(headerToken);

            bool isFldssc =
                markerToken.Equals("FLDSSC", StringComparison.CurrentCultureIgnoreCase)
                || Regex.IsMatch(markerToken, @"FLDSSC\d+", RegexOptions.IgnoreCase);

            return isCpOrFt && isFldssc;
        }

        private static bool IsCpOrFtPattern(PatternValidationInput i)
        {
            string target = i.Payload[i.PatternCount - 2];
            return IsStringStartWithCpOrFt(target);
        }

        private static void OnCpOrFtMatch(PatternValidationInput i)
        {
            if (i.Payload[i.PatternCount - 2].EndsWith("E", StringComparison.CurrentCultureIgnoreCase))
            {
                i.PatternType.IsEarly = true;
            }

            if (!Regex.IsMatch(i.Payload.Last(), DvrvRegex2, RegexOptions.IgnoreCase))
            {
                return;
            }

            if (i.Payload.Last().ContainsIgnoreCase("DV"))
            {
                i.PatternType.DvrvType = DvRvType.Dv;
            }
            else if (i.Payload.Last().ContainsIgnoreCase("RV"))
            {
                i.PatternType.DvrvType = DvRvType.Rv;
            }
        }

        private static bool IsMatchDvrvRegex2(PatternValidationInput i)
        {
            string last = i.Payload.Last();
            return Regex.IsMatch(last, DvrvRegex2, RegexOptions.IgnoreCase)
                && !last.Contains("UDRVER");
        }

        private static void OnDvrvRegex2Match(PatternValidationInput i)
        {
            if (i.Payload.Last().Contains("E")) //CP1EDV
            {
                i.PatternType.IsEarly = true;
            }

            if (
                i.Payload.Last().ContainsIgnoreCase("DV")
                || i.Payload.Last().ContainsIgnoreCase("DEFAULT")
                || i.Payload.Last().ContainsIgnoreCase("FIXED"))
            {
                i.PatternType.DvrvType = DvRvType.Dv;
            }
            else
            {
                i.PatternType.DvrvType = DvRvType.Rv;
            }
        }

        private EfusePatternType PopulateCfgEfusePatternType(
            EfusePatternType patType,
            string[] payload,
            int patternCnt
        )
        {
            //Only CFG need to judge DVRV
            patType.IsDvrv = true;
            PatternValidationInput input = new PatternValidationInput()
            {
                PatternType = patType,
                Payload = payload,
                PatternCount = patternCnt
            };
            foreach ((Func<PatternValidationInput, bool> shouldApply, Action<PatternValidationInput> update) in _rules)
            {
                if (shouldApply(input))
                {
                    update(input);
                    return patType;
                }
            }
            patType.IsDvrv = false;
            return patType;
        }

        private void ResolveDaaSns(string[] payload, EfusePatternType patType, int patternCnt)
        {
            if (patternCnt > 12)
            {
                if (Regex.IsMatch(payload[12], @"FLDSSC\d+", RegexOptions.IgnoreCase))
                {
                    patType.TestMode = EfuseTestMode.MarginRead;
                }
                else if (payload[12].Equals("FLDSSC", StringComparison.CurrentCultureIgnoreCase))
                {
                    patType.TestMode = EfuseTestMode.MarginReadFull;
                }
                else if (payload[12].Equals("DAP", StringComparison.CurrentCultureIgnoreCase))
                {
                    patType.TestMode = EfuseTestMode.ReadDap;
                }

                if (payload[patternCnt - 2].StartsWith("CP", StringComparison.CurrentCultureIgnoreCase) ||
                    payload[patternCnt - 2].StartsWith("FT", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (payload[patternCnt - 2].EndsWith("E", StringComparison.CurrentCultureIgnoreCase))
                    {
                        patType.IsEarly = true;
                    }
                }
            }

            if (patType.IsDvrv)
            {
                patType.TestMode = GetDvrvTestMode(patType, true);
            }
        }

        private void ResolveDaaPrg(string[] payload, EfusePatternType patType, int patternCnt)
        {
            if (patternCnt > 12)
            {
                if (Regex.IsMatch(payload[12], @"FLDSSC\d+", RegexOptions.IgnoreCase)
                    || payload[12].Equals("FLDSSC", StringComparison.CurrentCultureIgnoreCase))
                {
                    patType.TestMode = EfuseTestMode.WriteDssc;
                }
                else if (Regex.IsMatch(payload[12], @"Q\d+DSSC", RegexOptions.IgnoreCase))
                {
                    patType.TestMode = EfuseTestMode.WriteDsscQ;
                }
                else if (Regex.IsMatch(payload[12], @"H\d+DSSC", RegexOptions.IgnoreCase))
                {
                    patType.TestMode = EfuseTestMode.WriteDsscH;
                }

                if (payload[patternCnt - 2].StartsWith("CP", StringComparison.CurrentCultureIgnoreCase) ||
                    payload[patternCnt - 2].StartsWith("FT", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (payload[patternCnt - 2].EndsWith("E", StringComparison.CurrentCultureIgnoreCase))
                    {
                        patType.IsEarly = true;
                    }
                }
            }
            if (patternCnt > 11 && payload[11].Equals("HF", StringComparison.CurrentCultureIgnoreCase))
            {
                patType.TestMode = EfuseTestMode.WriteDsscHf;
            }

            if (patType.IsDvrv)
            {
                patType.TestMode = GetDvrvTestMode(patType, false);
            }
        }

        private void ResolveDaaTestMode(string[] payload, EfusePatternType patType, int patternCnt)
        {
            switch (payload[8].ToUpper())
            {
                case "SNS":
                    ResolveDaaSns(payload, patType, patternCnt);
                    break;
                case "PRG":
                    ResolveDaaPrg(payload, patType, patternCnt);
                    break;
            }
        }

        private void ResolveJtgSns(string[] payload, EfusePatternType patType, int patternCnt)
        {
            if (payload.Last().ContainsIgnoreCase("zero")
                || (patternCnt > 11
                && payload[11].Equals("PART", StringComparison.CurrentCultureIgnoreCase)))
            {
                patType.TestMode = EfuseTestMode.Unknow;
            }
            else if (payload[6].Equals("UFA", StringComparison.CurrentCultureIgnoreCase))
            {
                patType.TestMode = EfuseTestMode.Ufr;
            }
            else if (patternCnt > 12
                && payload[12].Equals("DAP", StringComparison.CurrentCultureIgnoreCase))
            {
                patType.TestMode = EfuseTestMode.JtagReadDap;
            }
            else if ((patternCnt > 11
                && payload[11].Equals("APB", StringComparison.CurrentCultureIgnoreCase))
                || payload[6].Equals("SEF", StringComparison.CurrentCultureIgnoreCase))
            {
                patType.TestMode = EfuseTestMode.ApbRead;
            }
            else
            {
                patType.TestMode = EfuseTestMode.JtagRead;
            }
        }

        private void ResolveJtgPrg(string[] payload, EfusePatternType patType, int patternCnt)
        {
            if (payload.Last().ContainsIgnoreCase("zero")
                || (patternCnt > 11
                && payload[11].Equals("PART", StringComparison.CurrentCultureIgnoreCase)))
            {
                patType.TestMode = EfuseTestMode.Unknow;
            }
            else if (payload[6].Equals("UFA", StringComparison.CurrentCultureIgnoreCase))
            {
                patType.TestMode = EfuseTestMode.Ufp;
            }
            else if ((patternCnt > 11
                && payload[11].Equals("APB", StringComparison.CurrentCultureIgnoreCase))
                || payload[6].Equals("SEF", StringComparison.CurrentCultureIgnoreCase))
            {
                patType.TestMode = EfuseTestMode.ApbWrite;
            }
            else
            {
                patType.TestMode = EfuseTestMode.JtagWrite;
            }
        }

        private void ResolveJtgUns(string[] payload, EfusePatternType patType, int patternCnt)
        {
            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                if (payload.Last().ContainsIgnoreCase("fullrd"))
                {
                    patType.TestMode = EfuseTestMode.MarginReadFull;
                }
                else if (payload.Last().ContainsIgnoreCase("fullwr"))
                {
                    patType.TestMode = EfuseTestMode.WriteDssc;
                }
                return;
            }
            if (
                (patternCnt > 12 && payload[12].Equals("UDRVER2", StringComparison.CurrentCultureIgnoreCase))
                || (patternCnt > 11 && payload[11].Equals("UDRVER2", StringComparison.CurrentCultureIgnoreCase)))
            {
                patType.TestMode = EfuseTestMode.Ver2;
            }
            else if (
                (patternCnt > 12 && payload[12].Equals("UDRVER1", StringComparison.CurrentCultureIgnoreCase))
                || (patternCnt > 11 && payload[11].Equals("UDRVER1", StringComparison.CurrentCultureIgnoreCase)))
            {
                patType.TestMode = EfuseTestMode.Ver1;
            }
            else if (payload[6].Equals("USO", StringComparison.CurrentCultureIgnoreCase))
            {
                patType.TestMode = EfuseTestMode.Uso;
            }
            else if (payload[6].Equals("USI", StringComparison.CurrentCultureIgnoreCase))
            {
                patType.TestMode = EfuseTestMode.Usi;
            }
        }

        private void ResolveJtgMode(string[] payload, EfusePatternType patType, int patternCnt)
        {
            switch (payload[8].ToUpper())
            {
                case "SNS":
                    ResolveJtgSns(payload, patType, patternCnt);
                    break;
                case "PRG":
                    ResolveJtgPrg(payload, patType, patternCnt);
                    break;
                case "UNS":
                    ResolveJtgUns(payload, patType, patternCnt);
                    break;
            }
        }

        internal EfusePatternType GetPatternType(string[] payload, bool isCfg)
        {
            var patType = new EfusePatternType
            {
                PatJob = "",
                IsDvrv = false,
                IsEarly = false,
                DvrvType = DvRvType.Not,
                TestMode = EfuseTestMode.Unknow,
            };
            int patternCnt = payload.Length;

            bool HasCrcMonAt(int index) =>
                patternCnt > index
                && payload[index].ContainsIgnoreCase("CRC")
                && payload[index].ContainsIgnoreCase("MON");
            int[] candidateIndexes = { 12, 11 };

            //CRC
            if (candidateIndexes.Any(HasCrcMonAt))
            {
                patType.TestMode = EfuseTestMode.Crc;
            }
            else if (isCfg)
            {
                patType = PopulateCfgEfusePatternType(patType, payload, patternCnt);
            }

            if (patType.TestMode.Equals(EfuseTestMode.Crc) || patType.IsDvrv)
            {
                patType.PatJob = GetJobFromPayloadPattern(payload);
            }

            if (patType.TestMode.Equals(EfuseTestMode.Crc))
            {
                return patType;
            }

            if (payload[7].Equals("DAA", StringComparison.CurrentCultureIgnoreCase))
            {
                ResolveDaaTestMode(payload, patType, patternCnt);
            }
            else if (
                payload[7].Equals("JTG", StringComparison.CurrentCultureIgnoreCase)
                && patternCnt > 8
            )
            {
                ResolveJtgMode(payload, patType, patternCnt);
            }
            return patType;
        }

        private EfuseTestMode GetDvrvTestMode(EfusePatternType patInfo, bool isRead)
        {
            EfuseTestMode testMode = patInfo.TestMode;
            if (!patInfo.IsDvrv)
            {
                return testMode;
            }

            if (isRead)
            {
                if (patInfo.DvrvType.Equals(DvRvType.FlatCheck))
                {
                    testMode = EfuseTestMode.FlatCheck;
                }
                else
                {
                    if (patInfo.DvrvType.Equals(DvRvType.Dv))
                    {
                        testMode = EfuseTestMode.DvRead;
                    }

                    if (patInfo.DvrvType.Equals(DvRvType.Rv))
                    {
                        testMode = EfuseTestMode.RvRead;
                    }
                }
            }
            else
            {
                if (patInfo.DvrvType.Equals(DvRvType.FlatCheck))
                {
                    testMode = EfuseTestMode.Unknow;
                }
                else
                {
                    if (patInfo.DvrvType.Equals(DvRvType.Dv))
                    {
                        testMode = EfuseTestMode.DvWrite;
                    }

                    if (patInfo.DvrvType.Equals(DvRvType.Rv))
                    {
                        testMode = EfuseTestMode.RvWrite;
                    }
                }
            }

            return testMode;
        }

        internal string GetDetailBankName(List<string> initList, string[] payload)
        {
            if (payload.Length > 9)
            {
                if (payload[9].StartsWith("ME", StringComparison.CurrentCultureIgnoreCase))
                {
                    return BankType.UdrE;
                }
                if (payload[9].StartsWith("MP", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (_onlyUdrP)
                    {
                        return BankType.UdrP;
                    }

                    if (initList.Any())
                    {
                        foreach (string init in initList)
                        {
                            string[] initStr = init.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                            if (initStr.Last().Equals("P0", StringComparison.CurrentCultureIgnoreCase))
                            {
                                return BankType.UdrP0;
                            }

                            if (initStr.Last().Equals("P1", StringComparison.CurrentCultureIgnoreCase))
                            {
                                return BankType.UdrP1;
                            }

                            if (initStr.Length > 9)
                            {
                                if (initStr[9].StartsWith("MPX0", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    return BankType.UdrP0;
                                }

                                if (initStr[9].StartsWith("MPX1", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    return BankType.UdrP1;
                                }
                            }
                        }
                    }
                    return BankType.UdrP;
                }
                if (payload[9].StartsWith("MM", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (initList.Any())
                    {
                        foreach (string init in initList)
                        {
                            string[] initStr = init.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                            if (initStr.Last().Equals("M0", StringComparison.CurrentCultureIgnoreCase))
                            {
                                return BankType.UdrM0;
                            }

                            if (initStr.Last().Equals("M1", StringComparison.CurrentCultureIgnoreCase))
                            {
                                return BankType.UdrM1;
                            }

                            if (initStr.Length > 9)
                            {
                                if (initStr[9].StartsWith("MMX0", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    return BankType.UdrM0;
                                }

                                if (initStr[9].StartsWith("MMX1", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    return BankType.UdrM1;
                                }
                            }
                        }
                    }
                    return BankType.UdrM;
                }
            }
            return BankType.Unknow;
        }
    }
}
