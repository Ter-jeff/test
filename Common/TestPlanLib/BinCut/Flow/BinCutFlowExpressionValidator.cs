using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

namespace TestPlanLib.BinCut.Flow
{
    internal partial class BinCutFlowExpressionValidator
    {
        [GeneratedRegex(BinCutFlowTable.RegexLvCoreVoltage, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex5();
        [GeneratedRegex(BinCutFlowTable.RegexLvCoreEvaluate, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex6();
        [GeneratedRegex(BinCutFlowTable.RegexLvCoreResult, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex7();
        [GeneratedRegex(BinCutFlowTable.RegexCoreHvcc, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex8();

        [GeneratedRegex(BinCutFlowTable.RegexCoreProduct1, RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex9();

        [GeneratedRegex(BinCutFlowTable.RegexCoreE1Product, RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex10();

        [GeneratedRegex(BinCutFlowTable.RegexRam1, RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex11();

        [GeneratedRegex(BinCutFlowTable.RegexLvOTher1, RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex12();

        [GeneratedRegex(BinCutFlowTable.RegexHvOTher1, RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex13();

        [GeneratedRegex(BinCutFlowTable.RegexmVWithProductMode, RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex14();

        [GeneratedRegex(BinCutFlowTable.RegexmVWithBinSearchMode, RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex15();

        [GeneratedRegex(BinCutFlowTable.RegexModeBinProduct, RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex16();

        [GeneratedRegex(BinCutFlowTable.RegexFunction, RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex17();

        [GeneratedRegex(BinCutFlowTable.RegexProductRatioMv, RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex18();

        [GeneratedRegex(BinCutFlowTable.RegexAllmVWithMode, RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex19();

        [GeneratedRegex("^Product$", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex20();

        [GeneratedRegex(@"^(?<voltage>M[a-zA-Z0-9]{4}[a-zA-Z0-9]?\s+E\d\s+Voltage)(?<gbStr>[+|-]?\s*.+\s*[+|-]?\s*.+)", RegexOptions.Compiled)]
        private static partial Regex MyRegex21();

        [GeneratedRegex("^" + BinCutFlowTable.RegexAllmV + "$", RegexOptions.IgnoreCase)]
        private static partial Regex MyRegex22();

        [GeneratedRegex("^" + BinCutFlowTable.RegexPerformance + @"\s+E\d\s+Voltage", RegexOptions.IgnoreCase)]
        private static partial Regex MyRegex23();

        private List<string> _gradeSearchPins = [];

        public void RefreshGradeSearchPins(List<BinCutFlowSheetRow> binCutFlowSheetRows)
        {
            var gradeSerachPins = new List<string>();
            int loopCnt = binCutFlowSheetRows[0].PinInfos.Count;
            for (int i = 0; i < loopCnt; i++)
            {
                foreach (BinCutFlowSheetRow row in binCutFlowSheetRows)
                {
                    if (row.TableType.Equals(EnumBinCutTableType.Lv) &&
                        row.PinInfos[i].PinContext.Contains("evaluate bin", StringComparison.OrdinalIgnoreCase))
                    {
                        gradeSerachPins.Add(row.PinInfos[i].PinName);
                    }
                }
            }
            _gradeSearchPins = [.. gradeSerachPins.Distinct()];
        }

        public bool IsValidExpress(string express, string binningDomain, List<string>? headerList = null, Dictionary<string, List<string>>? domainDic = null)
        {
            if (string.IsNullOrEmpty(express.Trim()))
            {
                return true;
            }

            if (MyRegex5().IsMatch(express))    //LV CORE_POWER => MS001 E1 Voltage
            {
                return true;
            }

            if (MyRegex6().IsMatch(express))   //LV CORE_POWER => MS001 Evaluate Bin
            {
                return true;
            }

            if (MyRegex7().IsMatch(express))     //LV CORE_POWER => MS001 Bin Result
            {
                return true;
            }

            if (MyRegex8().IsMatch(express))         //LV,HV CORE_POWER => HVCC Level at MC607
            {
                return true;
            }

            if (MyRegex9().IsMatch(express))     //CORE_POWER => MC606 Product
            {
                return true;
            }

            if (MyRegex10().IsMatch(express))     //CORE_POWER => MC606 E1 Product
            {
                return true;
            }

            if (MyRegex11().IsMatch(express))             //LV/HV RAM_POWER => MC601 CSRAM Product
            {
                return true;
            }

            if (MyRegex12().IsMatch(express))         //LV Others => CP LVCC
            {
                return true;
            }

            if (MyRegex13().IsMatch(express))         //HV Others => CP HVCC
            {
                return true;
            }

            if (MyRegex20().IsMatch(express))       //Product
            {
                return true;
            }

            if (MyRegex18().IsMatch(express))         //Product + 5.5%
            {
                return true;
            }

            if (MyRegex14().IsMatch(express)) //700.25mV (MS001 Product)
            {
                return true;
            }

            if (MyRegex15().IsMatch(express)) //700.25mV (MS001 BinSearch)
            {
                return true;
            }

            if (MyRegex16().IsMatch(express)) //MC606 BinX Product+10%
            {
                return true;
            }

            if (IsValidPinVoltageSyntax(express, binningDomain, headerList, domainDic))
            {
                return true;
            }

            return MyRegex17().IsMatch(express);
        }

        private bool IsValidPinVoltageSyntax(string express, string binningDomain, List<string>? headerList, Dictionary<string, List<string>>? domainDic)
        {
            if (domainDic != null && domainDic.TryGetValue(binningDomain, out List<string>? modeList))
            {
                if (modeList.Count == 1)
                {
                    if (MyRegex22().IsMatch(express)) //+700.25mV
                    {
                        return true;
                    }

                    if (headerList?.Exists(x => x.EqualsIgnoreCase(express)) == true) //CPVmax
                    {
                        return true;
                    }
                }
                else if (modeList.Count > 1)
                {
                    if (_gradeSearchPins.Any(x => x.EqualsIgnoreCase("VDD_" + binningDomain)))
                    {
                        if (MyRegex19().IsMatch(express)) //+700.25mV (MS001)
                        {
                            return true;
                        }
                    }
                    else
                    {
                        if (MyRegex22().IsMatch(express)) //+700.25mV
                        {
                            return true;
                        }
                    }
                }
            }

            if (headerList != null)
            {
                //MP005 Product-FT_GB_ROOM
                if (headerList.Select(item => BinCutFlowTable.RegexAllProduct + @"\s*[+|-]\s*" + item + @"(\s*[+|-]\s*\d+(\.\d+)?(\%|mV)$)?").Any(regex => Regex.IsMatch(express, regex, RegexOptions.IgnoreCase)))
                {
                    return true;
                }

                //MP005 BinningVmax
                if (headerList.Select(item => "^" + BinCutFlowTable.RegexPerformance + @"\s*" + item + "$").Any(regex => Regex.IsMatch(express, regex, RegexOptions.IgnoreCase)))
                {
                    return true;
                }
                //BinningVmax+5%
                if (headerList.Select(item => item + @"(\s*[+|-]\s*\d+(\.\d+)?(\%|mV)$)?").Any(regex => Regex.IsMatch(express, regex, RegexOptions.IgnoreCase)))
                {
                    return true;
                }

                if (MyRegex23().IsMatch(express)) //MC606 E1 Voltage +GB -GB
                {
                    string gbStrs = MyRegex21().Match(express).Groups["gbStr"].ToString().Trim();
                    if (gbStrs.Length != 0)
                    {
                        List<string> gbarr = [.. gbStrs.Split(['+', '-', ' '], StringSplitOptions.RemoveEmptyEntries)];
                        if (gbarr.Any(gb => !headerList.Exists(x => x.EqualsIgnoreCase(gb))))
                        {
                            return false;
                        }
                    }
                }
            }

            return false;
        }
    }
}
