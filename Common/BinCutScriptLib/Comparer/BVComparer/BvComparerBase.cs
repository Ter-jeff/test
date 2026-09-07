using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Comparer.BVChecker.GetVoltage;
using BinCutScriptLib.Comparer.Dssc;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;
using BinCutScriptLib.Utility;

using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib.BinCut.BinCutConfig;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc;

namespace BinCutScriptLib.Comparer.BVComparer
{
    public class BvComparerBase(OneGradeSearch oneGradeSearch, EnumJob enumJob, CheckManager checkManager, int tchCnt, EnumSearchType enumSearchType)
    {
        protected readonly EnumJob Job = enumJob;
        protected EnumBinCutTableType BinCutTableType;
        protected EnumPrintType PrintType;
        protected EnumSearchType SearchType = enumSearchType;
        protected BinCutFlowSheetRow? BinCutFlowSheetRow;
        protected GetVoltageBase? GetVoltageBase;
        protected readonly CheckManager CheckManager = checkManager;
        protected readonly InstanceBinCut InstanceBinCut = oneGradeSearch.InstanceBinCut!;
        protected readonly int TchCnt = tchCnt;
        //For debug
        private readonly OneGradeSearch _oneGradeSearch = oneGradeSearch;

        public bool CompareBvString(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, int stepIndex, OneStep oneStep, BvName bvName)
        {
            int lineNumber = 0;
            try
            {
                List<BinCutLineBase> dsscResultLines = oneStep.OneStepDssc;
                var oneStepBvLinesDic = new Dictionary<int, BvResults>();
                List<OffsetData> offsetDatas = GetOffsetData(oneStep.OneStepOffset);

                (bool flowControl, bool value) = CheckByLine(streamWriter, siteInfoArray, stepIndex, oneStep, bvName, ref lineNumber, dsscResultLines, oneStepBvLinesDic, offsetDatas, out bool isSkipCheckPattern, out bool hasPayload);

                if (!flowControl)
                {
                    return value;
                }
                if (oneStep.OneStepDebugErrorLine.Count != 0)
                {
                    foreach (BinCutLineBase errorline in oneStep.OneStepDebugErrorLine)
                    {
                        BinCutPrint.PrintCommomError(streamWriter, errorline, "Vmain Valt debug print line error!!!");
                    }
                }

                if (oneStep.OneStepHarvFailLine.Count != 0)
                {
                    foreach (BinCutLineBase failLine in oneStep.OneStepHarvFailLine)
                    {
                        string siteString = failLine.Line.Split(' ').First().Split('e').Last();
                        _ = int.TryParse(siteString, out int site);
                        siteInfoArray[site].IsHarvFail = true;
                    }
                }

                if (hasPayload)
                {
                    InstanceBinCut.HasPayload = true;
                }

                if (!_oneGradeSearch.IsCallInstance)
                {
                    var dsscCpmparer = new DsscComparer(InstanceBinCut, CheckManager);
                    dsscCpmparer.CompareDssc(streamWriter, siteInfoArray, oneStepBvLinesDic, dsscResultLines);
                }
                else
                {
                    for (int i = 0; i < dsscResultLines.Count; i++)
                    {
                        CheckManager.CheckDsscLine(dsscResultLines[i]);
                    }
                }
                return isSkipCheckPattern;
            }
            catch (Exception)
            {
                throw new Exception($"Exception happened when parsing line number {lineNumber} !!!");
            }
        }

        private (bool flowControl, bool value) CheckByLine(StreamWriter streamWriter, SiteInfo[] siteInfoArray, int stepIndex, OneStep oneStep, BvName bvName, ref int lineNumber, List<BinCutLineBase> binCutLineBases, Dictionary<int, BvResults> oneStepBvLinesDic, List<OffsetData> offsetDatas, out bool isSkipCheckPattern, out bool hasPayload)
        {
            #region check BVLine
            isSkipCheckPattern = false;
            hasPayload = false;
            if (BinCutConfig.FlagSkipPrintingSafeVoltage && oneStep.OneStepBvLineInfo.Count == 0)
            {
                return (flowControl: false, value: true);
            }

            for (int index = 0; index < oneStep.OneStepBvLineInfo.Count; index++)
            {
                string lineOnly = oneStep.OneStepBvLineInfo[index].LineOnly;
                lineNumber = oneStep.OneStepBvLineInfo[index].Line.LineNo;
                string instType = oneStep.OneStepBvLineInfo[index].InstType;
                if (!string.IsNullOrEmpty(instType))
                {
                    InstanceBinCut.CurInstType = instType;
                }

                int site = oneStep.OneStepBvLineInfo[index].Site;
                bool isError = true;
                string lvCmpStr = "";
                string nvCmpStr = "";
                string vrsCmpStr = "";
                if (stepIndex % 2 != 0 || _oneGradeSearch.IsCallInstance || BinCutConfig.FlagSkipPrintingSafeVoltage)
                {
                    hasPayload = true;
                    var offsetDatasbySite = offsetDatas.Where(x => x.Site == site).ToList();
                    BvResults pins = GetPinsData(bvName, siteInfoArray[site], offsetDatasbySite, streamWriter);
                    lvCmpStr = bvName.Name + "," + site + "," + pins.GetComposeString();
                    if (BvLineInfo.CompareBVline(lineOnly, lvCmpStr))
                    {
                        if (binCutLineBases.Count != 0)
                        {
                            int min = binCutLineBases.Min(x => x.LineNo);
                            if (!oneStepBvLinesDic.ContainsKey(site) &&
                                (_oneGradeSearch.IsCallInstance ||
                                (!_oneGradeSearch.IsCallInstance && oneStep.OneStepBvLineInfo[index].Line.LineNo > min)))
                            {
                                pins.Line = oneStep.OneStepBvLineInfo[index].Line;
                                oneStepBvLinesDic.Add(site, pins);
                            }
                        }
                        isError = false;
                    }
                    //Set for pattern coverage
                    isSkipCheckPattern = false;
                }
                else
                {
                    nvCmpStr = bvName.Name + "," + site + "," + BinCutData.NvTable!.GetCompareString(InstanceBinCut, bvName.Name);
                    if (!string.IsNullOrEmpty(BinCutConfig.SafeVoltageFollowPayload))
                    {
                        nvCmpStr = NotSafeVoltageFollowPayload(streamWriter, siteInfoArray, bvName, offsetDatas, site, nvCmpStr);
                    }

                    if (BvLineInfo.CompareBVline(lineOnly, nvCmpStr))
                    {
                        isError = false;
                    }
                    else
                    {
                        vrsCmpStr = bvName.Name + "," + site + "," + BinCutData.VrsTable!.GetCompareString(InstanceBinCut, bvName.Name);
                        if (!string.IsNullOrEmpty(BinCutConfig.SafeVoltageFollowPayload))
                        {
                            vrsCmpStr = NotSafeVoltageFollowPayload1(streamWriter, siteInfoArray, bvName, offsetDatas, site, vrsCmpStr);
                        }

                        if (BvLineInfo.CompareBVline(lineOnly, vrsCmpStr))
                        {
                            isError = false;
                        }
                    }
                    //Set for pattern coverage
                    isSkipCheckPattern = true;
                }

                CheckManager.CheckBvLine(oneStep.OneStepBvLineInfo[index].Line);

                if (isError || BinCutConfig.IsDebugPrint)
                {
                    PrintResult(streamWriter, siteInfoArray, oneStep, index, site, lvCmpStr, nvCmpStr, vrsCmpStr, binCutLineBases, bvName, !isError);
                    InstanceBinCut.IsBvPass = false;
                    siteInfoArray[site].CheckResult.IsBvPass = false;
                }

            }
            #endregion
            return (flowControl: true, value: default);
        }

        private string NotSafeVoltageFollowPayload1(StreamWriter streamWriter, SiteInfo[] siteInfoArray, BvName bvName, List<OffsetData> offsetDatas, int site, string vrsCmpStr)
        {
            string domain = "_" + BinCutConfig.SafeVoltageFollowPayload.Split(':').First().ToUpper() + "_";
            if (InstanceBinCut.CurInstanceName.ContainsIgnoreCase(domain))
            {
                var offsetDatasbySite = offsetDatas.Where(x => x.Site == site).ToList();
                BvResults pins = GetPinsData(bvName, siteInfoArray[site], offsetDatasbySite, streamWriter);
                var bvLine = new BvLine
                {
                    Line = vrsCmpStr
                };
                Dictionary<string, double> bvs = bvLine.GenDictionary();
                BvResults bvResults = BvComparerBaseHelpers.GetBvResults(pins, bvs);
                vrsCmpStr = bvName.Name + "," + site + "," + bvResults.GetComposeString();
            }

            return vrsCmpStr;
        }

        private string NotSafeVoltageFollowPayload(StreamWriter streamWriter, SiteInfo[] siteInfoArray, BvName bvName, List<OffsetData> offsetDatas, int site, string nvCmpStr)
        {
            string domain = "_" + BinCutConfig.SafeVoltageFollowPayload.Split(':').First().ToUpper() + "_";
            if (InstanceBinCut.CurInstanceName.ContainsIgnoreCase(domain))
            {
                var offsetDatasbySite = offsetDatas.Where(x => x.Site == site).ToList();
                BvResults pins = GetPinsData(bvName, siteInfoArray[site], offsetDatasbySite, streamWriter);
                var bvLine = new BvLine
                {
                    Line = nvCmpStr
                };
                Dictionary<string, double> bvs = bvLine.GenDictionary();
                var bvResults = new BvResults();
                List<string> arr = [.. BinCutConfig.SafeVoltageFollowPayload.Split(':').Last().Split(',')];
                foreach (KeyValuePair<string, double> bv in bvs)
                {
                    if (pins.Exists(x => x.PinName == bv.Key) &&
                        arr.Exists(x => x == bv.Key))
                    {
                        var bvResult = new BvResult()
                        {
                            PinName = bv.Key,
                            Voltage = pins.Find(x => x.PinName == bv.Key)!.Voltage
                        };
                        bvResults.Add(bvResult);
                    }
                    else
                    {
                        bvResults.Add(new BvResult() { PinName = bv.Key, Voltage = 1000 * bv.Value });
                    }
                }
                nvCmpStr = bvName.Name + "," + site + "," + bvResults.GetComposeString();
            }

            return nvCmpStr;
        }

        public bool CompareBvStringCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, OneStep oneStep, BvName bvName, string curInstanceName)
        {
            List<BinCutLineBase> dsscResultLines = oneStep.OneStepDssc;
            var oneStepBvLinesDic = new Dictionary<int, BvResults>();
            List<OffsetData> offsetDatas = GetOffsetData(oneStep.OneStepOffset);

            #region check BVLine
            bool isSkipCheckPattern = false;
            bool hasPayload = false;
            for (int index = 0; index < oneStep.SafeVoltageCs.Count; index++)
            {
                string line = oneStep.SafeVoltageCs[index].Line.Line;
                int site = oneStep.SafeVoltageCs[index].Line.GetSite();
                bool isError = true;
                string lvCmpStr = "";
                string vrsCmpStr = "";
                string nvCmpStr = $"[INFO]  [Site {site}] Bincut safe voltage: " + BinCutData.NvAllTable!.GetCompareStringCs(InstanceBinCut, bvName.Name);
                if (BvLineInfo.CompareBVlineCs(line, nvCmpStr))
                {
                    isError = false;
                }

                CheckManager.CheckBvLine(oneStep.SafeVoltageCs[index].Line);

                if (isError || BinCutConfig.IsDebugPrint)
                {
                    if (isError)
                    {
                        PrintSafeVoltageErrorCsharp(streamWriter, siteInfoArray, oneStep, index, site, lvCmpStr, nvCmpStr, vrsCmpStr, dsscResultLines);
                    }
                    else
                    {
                        PrintSafeVoltagePassCsharp(streamWriter, siteInfoArray, oneStep, index, site, lvCmpStr, nvCmpStr, vrsCmpStr, dsscResultLines);
                    }
                }
            }

            for (int index = 0; index < oneStep.OneStepBvLineInfo.Count; index++)
            {
                if (oneStep.OneStepBvLineInfo[index].Line == null)
                {
                    continue;
                }

                string line = oneStep.OneStepBvLineInfo[index].Line.Line;
                int site = oneStep.OneStepBvLineInfo[index].Site;

                bool isError = true;
                string lvCmpStr = "";
                string nvCmpStr = "";
                string vrsCmpStr = "";
                hasPayload = true;
                var offsetDatasbySite = offsetDatas.Where(x => x.Site == site).ToList();
                BvResults pins = GetPinsDataCsharp(bvName, siteInfoArray[site], offsetDatasbySite, streamWriter);

                lvCmpStr = $"[INFO]  [Site {site}] Bincut payload voltage: " + pins.GetComposeStringCs(oneStep.OneStepBvLineInfo[index].IsHardip);

                if (BvLineInfo.CompareBVlineCs(line, lvCmpStr))
                {
                    isError = false;
                }

                if (dsscResultLines.Count != 0)
                {
                    pins.Line = oneStep.OneStepBvLineInfo[index].Line;
                    oneStepBvLinesDic.TryAdd(site, pins);
                }

                isSkipCheckPattern = false;

                CheckManager.CheckBvLine(oneStep.OneStepBvLineInfo[index].Line);

                if (isError || BinCutConfig.IsDebugPrint)
                {
                    PrintResult(streamWriter, siteInfoArray, oneStep, index, site, lvCmpStr, nvCmpStr, vrsCmpStr, dsscResultLines, bvName, !isError);

                    InstanceBinCut.IsBvPass = false;
                    siteInfoArray[site].CheckResult.IsBvPass = false;
                }
                _oneGradeSearch.IsCallInstance = oneStep.OneStepBvLineInfo[index].IsHardip;
            }
            #endregion

            if (oneStep.OneStepDebugErrorLine.Count != 0)
            {
                foreach (BinCutLineBase errorline in oneStep.OneStepDebugErrorLine)
                {
                    BinCutPrint.PrintCommomError(streamWriter, errorline, "Vmain Valt debug print line error!!!");
                }
            }

            if (oneStep.OneStepHarvFailLine.Count != 0)
            {
                foreach (BinCutLineBase failLine in oneStep.OneStepHarvFailLine)
                {
                    string siteString = failLine.Line.Split(' ').First().Split('e').Last();
                    _ = int.TryParse(siteString, out int site);
                    siteInfoArray[site].IsHarvFail = true;
                }
            }

            if (hasPayload)
            {
                InstanceBinCut.HasPayload = true;
            }

            if (!_oneGradeSearch.IsCallInstance)
            {
                var dsscCpmparer = new DsscComparer(InstanceBinCut, CheckManager);
                dsscCpmparer.CompareDsscCs(streamWriter, siteInfoArray, oneStepBvLinesDic, dsscResultLines);
            }
            else
            {
                for (int i = 0; i < dsscResultLines.Count; i++)
                {
                    CheckManager.CheckDsscLine(dsscResultLines[i]);
                }
            }
            return isSkipCheckPattern;
        }

        public BvResults GetPinsData(BvName bvName, SiteInfo siteInfo, List<OffsetData> offsetDatas, StreamWriter streamWriter)
        {
            string t0TxJob = "";
            if (BinCutConfig.FlagT0TxHotFormat || BinCutConfig.FlagT0TxRoomFormat)
            {
                t0TxJob = BinCutConfig.FlagT0TxHotFormat ? "T0TX_HOT" : "T0TX_ROOM";
            }

            string job = !string.IsNullOrEmpty(t0TxJob) ? t0TxJob : GetVoltageBase!.Job.ToString();
            BinCutFlowSheetRow = BvComparerBaseHelpers.GetRow(GetVoltageBase!.InstanceBinCut.CurInstanceName.Contains("outsidebincut", StringComparison.OrdinalIgnoreCase) ? EnumBinCutFlowType.Outside : EnumBinCutFlowType.Normal, job, bvName.Name, BinCutTableType, siteInfo);
            var bvResults = new BvResults();
            if (BinCutFlowSheetRow == null)
            {
                return bvResults;
            }
            for (int index = 0; index < BinCutFlowSheetRow.PinInfos.Count; index++)
            {
                PinInfo pininfo = BinCutFlowSheetRow.PinInfos[index];
                var binCutExpress = new BinCutExpress
                {
                    Title = pininfo.PinName,
                    Express = pininfo.PinContext
                };
                double voltage = GetVoltage(bvName, siteInfo, offsetDatas, binCutExpress, streamWriter);
                var bvResult = new BvResult
                {
                    PinName = pininfo.PinName,
                    Voltage = voltage
                };
                if (binCutExpress.IsEvaluatePin && siteInfo.IsPreVddSearch && BinCutData.HarvPmodeTableOffsetSheet != null)
                {
                    HarvOffsetRow? offsetInfo = BinCutData.HarvPmodeTableOffsetSheet.Rows.Find(x => InstanceBinCut.CurInstanceName.Contains(x.Pmode));
                    if (offsetInfo != null)
                    {
                        List<string> byModes = [.. offsetInfo.ByMode.Split(',')];
                        foreach (string byMode in byModes)
                        {
                            if (!InstanceBinCut.CurInstanceName.Contains(byMode))
                            {
                                continue;
                            }

                            _ = double.TryParse(offsetInfo.Offset, out double offset);
                            bvResult.Voltage += offset;
                        }
                    }
                }
                bvResult.SelSramVoltage = GetSelSramVoltage(binCutExpress, siteInfo, voltage, streamWriter);
                if (InstanceBinCut.SelsrmMappingTalbeRows.Exists(x => x.LogicPins.EqualsIgnoreCase(bvResult.PinName)))
                {
                    bvResult.SelSramVth = BinCutConfig.FlagNewSelsrmVthEnable ? GetSelSramVth(binCutExpress, siteInfo) : GetSelSramVthOld(binCutExpress, siteInfo);
                }

                bvResults.Add(bvResult);
            }
            return bvResults;
        }

        public BvResults GetPinsDataCsharp(BvName bvName, SiteInfo siteInfo, List<OffsetData> offsetDatas, StreamWriter streamWriter)
        {
            string t0TxJob = "";
            if (BinCutConfig.FlagT0TxHotFormat || BinCutConfig.FlagT0TxRoomFormat)
            {
                t0TxJob = BinCutConfig.FlagT0TxHotFormat ? "T0TX_HOT" : "T0TX_ROOM";
            }

            string job = !string.IsNullOrEmpty(t0TxJob) ? t0TxJob : GetVoltageBase!.Job.ToString();
            EnumBinCutFlowType type = GetVoltageBase!.InstanceBinCut.CurInstanceName.Contains("outsidebincut", StringComparison.OrdinalIgnoreCase) ?
            EnumBinCutFlowType.Outside : EnumBinCutFlowType.Normal;
            string name = bvName.Name;
            if (type == EnumBinCutFlowType.Outside)
            {
                name = _oneGradeSearch.Pmode;
            }

            BinCutFlowSheetRow = BvComparerBaseHelpers.GetRowCsharp(type, job, name, BinCutTableType, siteInfo, GetVoltageBase!.InstanceBinCut);
            var bvResults = new BvResults();
            if (BinCutFlowSheetRow == null)
            {
                return bvResults;
            }
            for (int index = 0; index < BinCutFlowSheetRow.PinInfos.Count; index++)
            {
                PinInfo pininfo = BinCutFlowSheetRow.PinInfos[index];
                var binCutExpress = new BinCutExpress
                {
                    Title = pininfo.PinName,
                    Express = pininfo.PinContext
                };

                double voltage = GetVoltage(bvName, siteInfo, offsetDatas, binCutExpress, streamWriter);
                var bvResult = new BvResult
                {
                    PinName = pininfo.PinName,
                    Voltage = voltage
                };
                if (binCutExpress.IsEvaluatePin && siteInfo.IsPreVddSearch && BinCutData.HarvPmodeTableOffsetSheet != null)
                {
                    HarvOffsetRow? offsetInfo = BinCutData.HarvPmodeTableOffsetSheet.Rows.Find(x => InstanceBinCut.CurInstanceName.Contains(x.Pmode));
                    if (offsetInfo != null)
                    {
                        List<string> byModes = [.. offsetInfo.ByMode.Split(',')];
                        foreach (string byMode in byModes)
                        {
                            if (!InstanceBinCut.CurInstanceName.Contains(byMode))
                            {
                                continue;
                            }

                            _ = double.TryParse(offsetInfo.Offset, out double offset);
                            bvResult.Voltage += offset;
                        }
                    }
                }
                bvResult.SelSramVoltage = GetSelSramVoltage(binCutExpress, siteInfo, voltage, streamWriter);
                if (InstanceBinCut.SelsrmMappingTalbeRows.Exists(x => x.LogicPins.EqualsIgnoreCase(bvResult.PinName)))
                {
                    bvResult.SelSramVth = BinCutConfig.FlagNewSelsrmVthEnable ? GetSelSramVth(binCutExpress, siteInfo) : GetSelSramVthOld(binCutExpress, siteInfo);
                }

                bvResults.Add(bvResult);
            }
            foreach (KeyValuePair<string, int> affiliatedPin in BinCutData.BinCutFlowTables.First().AffiliatedPin)
            {
                if (BinCutData.BinCutFlowTables.First().PowerPins.ContainsValue(affiliatedPin.Value))
                {
                    string mainPinName = BinCutData.BinCutFlowTables.First().PowerPins.First(x => x.Value == affiliatedPin.Value).Key;
                    if (bvResults.First(x => x.PinName.EqualsIgnoreCase(mainPinName)) != null)
                    {
                        BvResult mainPinBvResult = bvResults.First(x => x.PinName.EqualsIgnoreCase(mainPinName));
                        double voltage = mainPinBvResult.Voltage;
                        double selsramVth = mainPinBvResult.SelSramVth;
                        double selsramVoltage = mainPinBvResult.SelSramVoltage;
                        var affiliatedBvResult = new BvResult
                        {
                            PinName = affiliatedPin.Key,
                            Voltage = voltage,
                            SelSramVth = selsramVth,
                            SelSramVoltage = selsramVoltage
                        };
                        bvResults.Add(affiliatedBvResult);
                    }
                }
            }
            return bvResults;
        }

        #region GetVoltage
        private double GetVoltage(BvName bvName, SiteInfo siteInfo, List<OffsetData> offsetDatas, BinCutExpress binCutExpress, StreamWriter streamWriter)
        {
            (List<string> expressList, string function) = BvComparerBaseHelpers.GetFunction(binCutExpress);
            double value;
            if (expressList.Count != 0)
            {
                var doubleList = new List<double>();
                foreach (string item in expressList)
                {
                    BinCutExpress newnewbinCutExpress = binCutExpress.Copy();
                    newnewbinCutExpress.Express = item;
                    doubleList.Add(ConvertExpressToValue(newnewbinCutExpress, bvName, siteInfo, offsetDatas, streamWriter));
                }
                value = BvComparerBaseHelpers.GetValueByFormula(function, doubleList);
            }
            else
            {
                value = ConvertExpressToValue(binCutExpress, bvName, siteInfo, offsetDatas, streamWriter);
            }
            return value;
        }

        private double ConvertExpressToValue(BinCutExpress binCutExpress, BvName bvName, SiteInfo siteInfo, List<OffsetData> offsetDatas, StreamWriter streamWriter)
        {
            binCutExpress = ParseExpress(binCutExpress);

            double value = GetVoltageBase!.GetValue(binCutExpress, siteInfo, bvName, streamWriter);

            value = BvComparerBaseHelpers.GetPercentageValue(binCutExpress.Express, value);

            double voltageValue = 0;
            double offsetValue = 0;
            GetOffSetValue(binCutExpress.Title, offsetDatas, ref voltageValue, ref offsetValue);
            if (voltageValue != 0)
            {
                return voltageValue;
            }

            return value + offsetValue;
        }

        private BinCutExpress ParseExpress(BinCutExpress binCutExpress)
        {
            BvComparerBaseHelpers.ReplacePin(binCutExpress);

            ChangeToBinResult(binCutExpress);

            return binCutExpress;
        }

        private void ChangeToBinResult(BinCutExpress binCutExpress)
        {
            if (Job != EnumJob.CP1)
            {
                return;
            }

            if (GetVoltageBase!.InstanceBinCut.CurInstanceName.Contains("_BINRESULT_", StringComparison.OrdinalIgnoreCase) &&
                !Reg.RegexAllmV.IsMatch(binCutExpress.Express))
            {
                if (BinCutData.ModeVsCorePowerName.ContainsValue(binCutExpress.Title))
                {
                    string mode = BinCutAlgorithmService.GetModeByName(binCutExpress.Express);
                    binCutExpress.Express = mode + " Bin Result";
                }
            }
        }

        protected static List<OffsetData> GetOffsetData(List<OffsetLine1> offsetLine1s)
        {
            var offsetDatas = new List<OffsetData>();
            if (offsetLine1s != null && offsetLine1s.Count != 0)
            {
                offsetDatas.AddRange(offsetLine1s.Select(row => row.GetOffsetData()));
            }
            return offsetDatas;
        }

        private static void GetOffSetValue(string power, List<OffsetData> offsetDatas, ref double voltageValue, ref double offsetValue)
        {
            if (offsetDatas != null && offsetDatas.Exists(x => x.PinName.EqualsIgnoreCase(power)))
            {
                voltageValue = offsetDatas.Find(x => x.PinName.EqualsIgnoreCase(power))!.VoltageValue;
                offsetValue = offsetDatas.Find(x => x.PinName.EqualsIgnoreCase(power))!.OffSetValue;
            }
        }
        #endregion

        public static double GetSelSramVoltage(BinCutExpress binCutExpress, SiteInfo siteInfo, double value, StreamWriter streamWriter)
        {
            if (binCutExpress.IsFunction())
            {
                return value;
            }

            if (binCutExpress.PowerType == EnumPowerType.Core_Power)
            {
                PowerZone powerZone = siteInfo.GetPowerZone(binCutExpress);
                string mode = BinCutAlgorithmService.GetModeByName(binCutExpress.Express);
                double productValue;
                double retVal = 0.0;
                //0: MS001_0 for evaluation      ex: MS001 Evaluate Bin
                //1: MS002_1 for result bin      ex: MG004 Bin Result
                //2: MC607_2 for HVCC            ex: HVCC Level at MC607
                //3: MS001_3_E1 for equation E1  ex: MC601 E1 Voltage
                switch (binCutExpress.ExpressType)
                {
                    case EnumExpressType.Product:
                        {
                            productValue = powerZone.IdsValue == 0 || BinCutConfig.IsBinCutJobForStepSearch == false
                                ? SiteInfoHelpers.GetEfuseProductValue(mode, siteInfo.EFuseValues)
                                : powerZone.GetFinalProductValue(streamWriter);
                            retVal = productValue;
                            break;
                        }
                    case EnumExpressType.E1Product:
                        {
                            retVal = BvComparerBaseHelpers.HandleE1Product(binCutExpress, siteInfo, powerZone);
                            break;
                        }
                    case EnumExpressType.E1ProductGb:
                        {
                            retVal = BvComparerBaseHelpers.HandleE1Product(binCutExpress, siteInfo, powerZone);
                            break;
                        }
                    case EnumExpressType.BinProduct:
                        {
                            retVal = BvComparerBaseHelpers.HandleBinProduct(binCutExpress, powerZone, retVal);
                            break;
                        }
                    case EnumExpressType.BinProductGb:
                        {
                            retVal = BvComparerBaseHelpers.HandleBinProduct(binCutExpress, powerZone, retVal);
                            break;
                        }
                    case EnumExpressType.ProductGb:
                        {
                            productValue = powerZone.IdsValue == 0 || BinCutConfig.IsBinCutJobForStepSearch == false
                                ? SiteInfoHelpers.GetEfuseProductValue(mode, siteInfo.EFuseValues)
                                : powerZone.GetFinalProductValue(streamWriter);
                            retVal = productValue;
                            break;
                        }
                    case EnumExpressType.E1Voltage:
                        {
                            retVal = BvComparerBaseHelpers.HandleE1Voltage(binCutExpress, siteInfo, powerZone);
                            break;
                        }
                    case EnumExpressType.Evaluate:
                        {
                            retVal = SiteInfoHelpers.GetVoltageOfEvaluation(powerZone);
                            break;
                        }
                    case EnumExpressType.BinResult:
                        {
                            if (siteInfo.IsPreVddSearch)
                            {
                                int index = powerZone.PossibleSteps.FindIndex(x => x.Bin == siteInfo.Bin && x.EqName == 1);
                                retVal = index != -1 ? powerZone.PossibleSteps[index].Lvcc : powerZone.AllSteps.Last().Lvcc;
                                break;
                            }
                            retVal = powerZone.SearchStatus != EnumSearchStatus.Search
                                ? powerZone.AllSteps.Last().Lvcc
                                : powerZone.GetFinalLvcc();
                            break;
                        }
                }
                return retVal;
            }
            return 0;
        }

        public double GetSelSramVthOld(BinCutExpress binCutExpress, SiteInfo siteInfo)
        {
            string retVal = "";
            if (binCutExpress.PowerType == EnumPowerType.Core_Power)
            {
                PowerZone powerZone = siteInfo.GetPowerZone(binCutExpress);
                if (Job == EnumJob.CP1 && PrintType == EnumPrintType.LV)
                {
                    for (int pStepIdx = 0; pStepIdx < powerZone.GetPosCount(); pStepIdx++)
                    {
                        if (powerZone.PossibleSteps[pStepIdx].Bin == siteInfo.Bin)
                        {
                            retVal = powerZone.PossibleSteps[pStepIdx].SramthreshCp1;
                        }
                    }
                }
                else if (PrintType == EnumPrintType.POST && Job == EnumJob.CP1 && BinCutTableType == EnumBinCutTableType.Lv)
                {
                    for (int pStepIdx = 0; pStepIdx < powerZone.GetPosCount(); pStepIdx++)
                    {
                        if (powerZone.PossibleSteps[pStepIdx].Bin == siteInfo.Bin)
                        {
                            retVal = powerZone.PossibleSteps[pStepIdx].SramthreshCp1;
                        }
                    }
                }
                else
                {
                    for (int pStepIdx = 0; pStepIdx < powerZone.GetPosCount(); pStepIdx++)
                    {
                        if (powerZone.PossibleSteps[pStepIdx].Bin == siteInfo.Bin)
                        {
                            retVal = powerZone.PossibleSteps[pStepIdx].SramthreshProduct;
                        }
                    }
                }
            }
            _ = double.TryParse(retVal, out double outValue);
            return outValue;
        }

        public static double GetSelSramVth(BinCutExpress binCutExpress, SiteInfo siteInfo)
        {
            string retVal = "";
            if (binCutExpress.PowerType == EnumPowerType.Core_Power)
            {
                PowerZone powerZone = siteInfo.GetPowerZone(binCutExpress);
                for (int pStepIdx = 0; pStepIdx < powerZone.GetPosCount(); pStepIdx++)
                {
                    if (powerZone.PossibleSteps[pStepIdx].Bin == siteInfo.Bin)
                    {
                        PowerStep targetStep = powerZone.PossibleSteps[pStepIdx];
                        retVal = binCutExpress.IsProduct ? targetStep.SramthreshProduct : targetStep.SramthreshCp1;
                    }
                }
                _ = double.TryParse(retVal, out double outValue);
                return outValue;
            }
            return 0;
        }

        protected virtual void PrintResult(StreamWriter streamWriter, SiteInfo[] siteInfoArray, OneStep oneStep, int index, int site, string lvCmpStr, string nvCmpStr, string vrsCmpStr, List<BinCutLineBase> binCutLineBases, BvName bvName, bool isPass)
        {
            int eqName = -1;
            int bin = -1;
            if (PrintType == EnumPrintType.LV && SearchType == EnumSearchType.GradeSearch)
            {
                int step = siteInfoArray[site].AllPowers[bvName.Index].Step;
                PowerStep pwrRef = siteInfoArray[site].AllPowers[bvName.Index].PossibleSteps[step];
                eqName = pwrRef.EqName;
                bin = pwrRef.Bin;
            }
            BinCutPrint.PrintBv(PrintType, SearchType, streamWriter, siteInfoArray, TchCnt, InstanceBinCut.CurInstanceName, oneStep.OneStepBvLineInfo[index], site, lvCmpStr, bin, eqName, nvCmpStr, vrsCmpStr, binCutLineBases, isPass);
        }

        private void PrintSafeVoltageErrorCsharp(StreamWriter streamWriter, SiteInfo[] siteInfoArray, OneStep oneStep, int index, int site, string lvCmpStr, string nvCmpStr, string vrsCmpStr, List<BinCutLineBase> binCutLineBases)
        {
            BinCutPrint.PrintSafeVoltage(PrintType, streamWriter, siteInfoArray, TchCnt, InstanceBinCut.CurInstanceName, oneStep.SafeVoltageCs[index], site, lvCmpStr, nvCmpStr, vrsCmpStr, binCutLineBases);
        }

        private void PrintSafeVoltagePassCsharp(StreamWriter streamWriter, SiteInfo[] siteInfoArray, OneStep oneStep, int index, int site, string lvCmpStr, string nvCmpStr, string vrsCmpStr, List<BinCutLineBase> binCutLineBases)
        {
            BinCutPrint.PrintSafeVoltage(PrintType, streamWriter, siteInfoArray, TchCnt, InstanceBinCut.CurInstanceName, oneStep.SafeVoltageCs[index], site, lvCmpStr, nvCmpStr, vrsCmpStr, binCutLineBases, true);
        }
    }
}
