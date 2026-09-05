using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using TestPlanLib.BinCut;

namespace BinCutScriptLib.Comparer.Dssc
{
    internal class DsscComparer(InstanceBinCut instanceBinCut, CheckManager checkManager)
    {
        private readonly CheckManager _checkManage = checkManager;
        private readonly InstanceBinCut _instanceBinCut = instanceBinCut;

        public void CompareDssc(StreamWriter streamWriter, SiteInfo[] siteInfoArray, Dictionary<int, BvResults> bvResultsDic, List<BinCutLineBase> binCutLineBases)
        {
            if (bvResultsDic.Count != 0)
            {
                int lineIdx;
                for (lineIdx = 0; lineIdx < binCutLineBases.Count; lineIdx++)
                {
                    string line = binCutLineBases[lineIdx].Line;
                    int lineNo = binCutLineBases[lineIdx].LineNo;

                    string[] spt1 = line.Split([','], StringSplitOptions.RemoveEmptyEntries);
                    if (spt1.Length < 3)
                    {
                        continue;
                    }

                    int site = int.Parse(spt1[1]);
                    if (!bvResultsDic.TryGetValue(site, out BvResults? value))
                    {
                        continue;
                    }

                    string refLineNo = bvResultsDic.ContainsKey(site) ? value.Line!.LineNo.ToString() : "";
                    BvResults bvResults = GetDsscBvResults(bvResultsDic[site], _instanceBinCut, line.StartsWith("SELSRAM_DSSC_Bit_Str"));

                    //SELSRAM_Compare_Bit_Str,2,1111111(LSB->MSB),I=1,A=1,G=1,E=1,P=1,D=1,S=1
                    if (line.StartsWithIgnoreCase("SELSRAM_Compare_Bit_Str"))
                    {
                        string dssc = Reg.RegexLsb.Replace(spt1[2], "");
                        string expectDssc = string.Join("", bvResults.Select(x => x.SelsramCompareBitStr));
                        //allDice[site].Currentdssc = dssc;
                        if (!DsscStringCompare(dssc, expectDssc))
                        {
                            BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo + " & " + refLineNo, site, dssc, expectDssc, "[SELSRAM_Compare_Bit_Str]", _instanceBinCut.CurInstanceName);
                            siteInfoArray[site].CheckResult.IsDsscPass = false;
                        }
                        _checkManage.CheckDsscLine(binCutLineBases[lineIdx]);
                    }

                    //SELSRAM_DSSC_Bit_Str,5,00000
                    else if (line.StartsWithIgnoreCase("SELSRAM_DSSC_Bit_Str"))
                    {
                        string dssc = spt1[2];
                        string expectDssc = string.Join("", bvResults.Select(x => x.SelsramDsscBitStr));
                        siteInfoArray[site].Currentdssc = expectDssc;
                        if (!DsscStringCompare(dssc, expectDssc))
                        {
                            BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo + " & " + refLineNo, site, dssc, expectDssc, "[SELSRAM_DSSC_Bit_Str]", _instanceBinCut.CurInstanceName);
                            siteInfoArray[site].CheckResult.IsDsscPass = false;
                        }
                        _checkManage.CheckDsscLine(binCutLineBases[lineIdx]);
                    }

                    //SRAM_Vth(DCVS),4,I=0.725V,A=0.728V,G=0.725V,D=0.725V,S=0.725V
                    else if (line.StartsWithIgnoreCase("SRAM_Vth"))
                    {
                        var vthList = new List<string>();
                        for (int i = 2; i < spt1.Length; i++)
                        {
                            vthList.Add(spt1[i].Split('=')[1].Replace("V", ""));
                        }

                        string vth = string.Join(",", vthList);
                        string expectVth = string.Join(",", bvResults.Select(x => $"{Math.Floor(x.SelSramVth) / 1000:F3}"));
                        if (vth != expectVth)
                        {
                            BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo + " & " + refLineNo, site, vth, expectVth, "[SRAM_Vth]", _instanceBinCut.CurInstanceName);
                            siteInfoArray[site].CheckResult.IsDsscPass = false;
                        }
                        _checkManage.CheckDsscLine(binCutLineBases[lineIdx]);
                    }

                    //SelSram_voltage,3,I=0.681V,A=0.721V,G=0.650V,D=0.709V,S=0.550V
                    else if (line.StartsWithIgnoreCase("SelSram_voltage"))
                    {
                        var selSramVoltageList = new List<string>();
                        for (int i = 2; i < spt1.Length; i++)
                        {
                            selSramVoltageList.Add(spt1[i].Split('=')[1].Replace("V", "").Trim());
                        }

                        string selSramVoltage = string.Join(",", selSramVoltageList);
                        string expectsedSramVoltage = string.Join(",", bvResults.Select(x => $"{Math.Floor(x.SelSramVoltage) / 1000:F3}"));
                        if (!DsscStringCompare(selSramVoltage, expectsedSramVoltage))
                        {
                            BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo + " & " + refLineNo, site, selSramVoltage, expectsedSramVoltage, "[SelSram_voltage]", _instanceBinCut.CurInstanceName);
                            siteInfoArray[site].CheckResult.IsDsscPass = false;
                        }
                        _checkManage.CheckDsscLine(binCutLineBases[lineIdx]);
                    }
                }
                binCutLineBases.RemoveRange(0, lineIdx < binCutLineBases.Count ? lineIdx : binCutLineBases.Count);
            }
        }

        public void CompareDsscCs(StreamWriter streamWriter, SiteInfo[] siteInfoArray, Dictionary<int, BvResults> bvResultsDic, List<BinCutLineBase> binCutLineBases)
        {
            if (bvResultsDic.Count != 0)
            {
                int lineIdx;
                for (lineIdx = 0; lineIdx < binCutLineBases.Count; lineIdx++)
                {
                    string line = binCutLineBases[lineIdx].Line;
                    int lineNo = binCutLineBases[lineIdx].LineNo;
                    string lineContent = line;

                    string[] spt1 = line.Split(['[', ']', ' ', '(', ')'], StringSplitOptions.RemoveEmptyEntries);
                    if (spt1.Length < 3)
                    {
                        continue;
                    }

                    int site = int.Parse(spt1[2]);
                    if (!bvResultsDic.TryGetValue(site, out BvResults? value))
                    {
                        continue;
                    }

                    string refLineNo = bvResultsDic.ContainsKey(site) ? value.Line!.LineNo.ToString() : "";
                    BvResults bvResults = GetDsscBvResults(bvResultsDic[site], _instanceBinCut, line.StartsWith("[INFO]  [Site "));

                    if (line.StartsWithIgnoreCase("[INFO]  [Site "))
                    {
                        string dssc = "";
                        string lineContentForDssc = line.Split(']').Last();
                        //[INFO]  [Site 1] 1(disp_sram_rail),1(gpu_sram_rail),1(dcs_sram_rail),1(soc_sram_rail)
                        if (line.Contains('(') && line.Contains(','))
                        {
                            string[] segments = lineContentForDssc.Split(',');

                            foreach (string segment in segments)
                            {
                                if (segment.Contains("_sram_"))
                                {
                                    int bitEndIndex = segment.IndexOf('(');
                                    if (bitEndIndex > 0)
                                    {
                                        string bit = segment[..bitEndIndex].Trim();
                                        dssc += bit;
                                    }
                                }
                            }
                        }
                        //[INFO]  [Site 1] 1111
                        //[INFO]  [Site 3] 1111(C)
                        else
                        {
                            dssc = spt1[3];
                        }

                        string expectDssc = string.Join("", bvResults.Select(x => x.SelsramDsscBitStr));
                        siteInfoArray[site].Currentdssc = expectDssc;
                        bool fail = !DsscStringCompare(dssc, expectDssc);

                        if (fail || BinCutConfig.IsDebugPrint)
                        {
                            if (fail)
                            {
                                string[] lineCheckFormat = lineContent.Split(['(', ')'], StringSplitOptions.RemoveEmptyEntries);
                                if (lineCheckFormat.Length == 2)
                                {
                                    expectDssc = expectDssc + "(" + lineCheckFormat[1] + ")";
                                }

                                if (lineContent.Contains(','))
                                {
                                    if (lineContent.Contains(',') && !string.IsNullOrEmpty(expectDssc))
                                    {
                                        string datalogContent = line.Split(']').Last();
                                        string[] segments = datalogContent.Split(',');

                                        if (expectDssc.Length == segments.Length)
                                        {
                                            string newExpectDssc = string.Empty;
                                            for (int i = 0; i < expectDssc.Length; i++)
                                            {
                                                string[] parts = segments[i].Split('(');
                                                if (parts.Length > 1)
                                                {
                                                    newExpectDssc += expectDssc[i] + "(" + parts[1];
                                                }
                                                else
                                                {
                                                    newExpectDssc += expectDssc[i] + segments[i];
                                                }

                                                if (i < expectDssc.Length - 1)
                                                {
                                                    newExpectDssc += ",";
                                                }
                                            }

                                            expectDssc = newExpectDssc;
                                        }
                                    }
                                }

                                BinCutPrint.PrintSelsramDifferenceCs(streamWriter, siteInfoArray, line, lineNo + " & " + refLineNo, site, expectDssc, "Selsram", _instanceBinCut.CurInstanceName);
                                siteInfoArray[site].CheckResult.IsDsscPass = false;
                            }
                            else
                            {
                                BinCutPrint.PrintSelsramDifferencePassCs(streamWriter, siteInfoArray, line, lineNo + " & " + refLineNo, site, expectDssc, "Selsram", _instanceBinCut.CurInstanceName);
                            }
                        }
                        _checkManage.CheckDsscLine(binCutLineBases[lineIdx]);
                    }
                }
                binCutLineBases.RemoveRange(0, lineIdx < binCutLineBases.Count ? lineIdx : binCutLineBases.Count);
            }
        }

        private static bool DsscStringCompare(string dssc, string expectDssc)
        {
            bool flag = true;
            if (dssc.Length != expectDssc.Length)
            {
                return false;
            }

            for (int i = 0; i < expectDssc.Length; i++)
            {
                if (expectDssc[i] != 'X')
                {
                    if (expectDssc[i] != dssc[i])
                    {
                        flag = false;
                        break;
                    }
                }
            }
            return flag;
        }

        private static BvResults GetDsscBvResults(BvResults bvResults, InstanceBinCut instanceBinCut, bool dsscTotalBits)
        {
            var newBvResults = new BvResults();
            List<SelsrmMappingTableRow> selsrmMappingTalbeRows = instanceBinCut.SelsrmMappingTalbeRows;
            foreach (SelsrmMappingTableRow row in selsrmMappingTalbeRows)
            {
                if (row.LogicPins.EqualsIgnoreCase("PRESERVED"))
                {
                    if (!dsscTotalBits)
                    {
                        continue;
                    }

                    newBvResults.Add(new BvResult { SelsramDsscBitStr = row.Selsrm1 });
                }

                if (bvResults.Exists(x => x.PinName.EqualsIgnoreCase(row.LogicPins)))
                {
                    BvResult bvResult = bvResults.Find(x => x.PinName.EqualsIgnoreCase(row.LogicPins))!;

                    if (bvResult.SelSramVoltage == bvResult.SelSramVth)
                    {
                        bvResult.SelsramCompareBitStr = "X";
                        bvResult.SelsramDsscBitStr = "X";
                    }
                    else if (bvResult.SelSramVoltage > bvResult.SelSramVth)
                    {
                        bvResult.SelsramCompareBitStr = row.Selsrm0 == "0" ? "0" : "1";
                        bvResult.SelsramDsscBitStr = row.Selsrm1 == "1" ? "0" : "1";
                    }
                    if (bvResult.SelSramVoltage < bvResult.SelSramVth)
                    {
                        bvResult.SelsramCompareBitStr = row.Selsrm0 == "0" ? "1" : "0";
                        bvResult.SelsramDsscBitStr = row.Selsrm1 == "1" ? "1" : "0";
                    }
                    newBvResults.Add(bvResult);
                }
            }
            return newBvResults;
        }

    }
}
