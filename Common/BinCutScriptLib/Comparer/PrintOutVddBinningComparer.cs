using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;

using IgxlLib.Enums;

namespace BinCutScriptLib.Comparer
{
    internal partial class PrintOutVddBinningComparer
    {

        [GeneratedRegex(@"\d+\.[a-z]\d+\w*", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        public static readonly Regex RegxChannel = MyRegex();

        public static void CompareCp1PrintOutCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, ref List<BinCutLineBase> binCutLineBases)
        {
            for (int i = 0; i < binCutLineBases.Count; i++)
            {
                string line = binCutLineBases[i].Line;
                string[] arr = line.Split([" K ", " ", "Define"], StringSplitOptions.RemoveEmptyEntries);

                if (arr.Length < 2)
                {
                    continue;
                }

                if (arr[1].Length != 1)
                {
                    continue;
                }

                int site = int.Parse(arr[1]);
                string pwrTypeLog = arr[3];
                string pwrNameLog = arr[2];
                int channelIndex = GetChannelIndex(arr);
                int measureIndex = GetMeasureIndex(channelIndex, arr);
                _ = double.TryParse(arr[measureIndex], out double log);
                for (int pwrIdx = 0; pwrIdx < siteInfoArray[site].AllPowers.Count; pwrIdx++)
                {
                    if (siteInfoArray[site].AllPowers[pwrIdx].PinMode.Contains(pwrNameLog))
                    {
                        PowerZone powerZone = siteInfoArray[site].AllPowers[pwrIdx];
                        string mode = pwrNameLog.Split('_').Last();
                        if (pwrTypeLog == "Equation_Num")
                        {
                            int eq = powerZone.SearchStatus != EnumSearchStatus.Search ? 1 : powerZone.GetFinalEqName();
                            if (eq != (int)log)
                            {
                                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, binCutLineBases[i], site, log.ToString(), eq.ToString(), "<PrintOutVddBinning> : Equation_Num");
                            }
                        }
                        else if (pwrTypeLog == "Bincut_Num")
                        {
                            int bin = powerZone.SearchStatus != EnumSearchStatus.Search ? 1 : powerZone.GetFinalBin();
                            if (bin != (int)log)
                            {
                                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, binCutLineBases[i], site, log.ToString(), bin.ToString(), "<PrintOutVddBinning> : Bincut_Num");
                            }
                        }
                        else if (pwrTypeLog == "Monotonicity_offset")
                        {
                            double offset = powerZone.GetMonotonicityOffset();
                            if (offset != log)
                            {
                                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, binCutLineBases[i], site, log.ToString(), offset.ToString(), "<PrintOutVddBinning> : Monotonicity_offset");
                            }
                        }
                        else if (pwrTypeLog == "BV")
                        {
                            double pwrLvcc = SiteInfoHelpers.GetLvcc(powerZone, mode, siteInfoArray[site].EFuseValues);
                            if (Math.Abs(pwrLvcc - log) > 0.001)
                            {
                                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, binCutLineBases[i], site, log.ToString(), pwrLvcc.ToString(), "<PrintOutVddBinning> : Lvcc");
                            }
                        }
                        else if (pwrTypeLog == "Product")
                        {
                            double product = powerZone.FinalStep == -1 ? 0 : SiteInfoHelpers.GetProductValue(powerZone, mode, streamWriter, siteInfoArray[site].EFuseValues);
                            if (Math.Abs(product - log) > 0.001)
                            {
                                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, binCutLineBases[i], site, log.ToString(), product.ToString(), "<PrintOutVddBinning> : Product");
                            }
                        }
                        else if (pwrTypeLog == "IDS")
                        {
                            double ids = powerZone.IdsValue;
                            if (Math.Abs(ids - log) > 0.001)
                            {
                                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, binCutLineBases[i], site, log.ToString(), ids.ToString(), "<PrintOutVddBinning> : Product");
                            }
                        }

                        break;
                    }
                }
            }
        }

        public static void CompareCp1PrintOut(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, ref List<BinCutLineBase> binCutLineBases, EnumJob enumJob)
        {
            for (int lnIdx = 0; lnIdx < binCutLineBases.Count; lnIdx++)
            {
                string line = binCutLineBases[lnIdx].Line;
                int lineNo = binCutLineBases[lnIdx].LineNo;
                string[] arr = line.Split([" K ", " ", "Define"], StringSplitOptions.RemoveEmptyEntries);

                if (arr.Length < 2)
                {
                    continue;
                }

                if (arr[1].Length != 1)
                {
                    continue;
                }

                int site = int.Parse(arr[1]);
                //C or V
                char pwrTypeLog = arr[3][0];
                string pwrNameLog = arr[2];
                int channelIndex = GetChannelIndex(arr);
                int measureIndex = GetMeasureIndex(channelIndex, arr);

                _ = double.TryParse(arr[measureIndex], out _);
                double pwrLvccLog = double.Parse(arr[measureIndex]);
                if (arr[3] == enumJob.ToString())
                {
                    pwrTypeLog = 'C';
                }

                if (pwrLvccLog == 0)
                {
                    return;
                }

                if (pwrLvccLog < 5.0)
                {
                    pwrLvccLog *= 1000.0;
                }

                for (int pwrIdx = 0; pwrIdx < siteInfoArray[site].AllPowers.Count; pwrIdx++)
                {
                    if (siteInfoArray[site].AllPowers[pwrIdx].PinMode.Contains(pwrNameLog))
                    {
                        PowerZone powerZone = siteInfoArray[site].AllPowers[pwrIdx];
                        string mode = pwrNameLog.Split('_').Last();
                        double product = SiteInfoHelpers.GetProductValue(powerZone, mode, streamWriter, siteInfoArray[site].EFuseValues);
                        double pwrLvcc = SiteInfoHelpers.GetLvcc(powerZone, mode, siteInfoArray[site].EFuseValues);

                        switch (pwrTypeLog)
                        {
                            case 'C':
                                break;
                            case 'V':
                                pwrLvcc = product;
                                break;
                        }

                        if (pwrLvccLog > 1000.0)
                        {
                            if ((int)pwrLvcc != (int)pwrLvccLog)
                            {
                                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, pwrLvccLog.ToString(), pwrLvcc.ToString(), "<PrintOutVddBinning> - Lvcc");
                            }
                        }
                        else
                        {
                            if (Math.Abs(pwrLvcc - pwrLvccLog) > 0.001)
                            {
                                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, pwrLvccLog.ToString(), pwrLvcc.ToString(), "<PrintOutVddBinning> - Product");
                            }
                        }
                        break;
                    }
                }
            }
        }

        public static void CompareCp2PrintOut(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, ref List<BinCutLineBase> binCutLineBases)
        {
            const double doubleLsb = 1.0e-6;
            int lnIdx;
            //51000000 0     VDD_SOC_MS001 CP2                                                 -1       584.3750       643.7500           690.6250       0.0000         0       
            //51000001 0     VDD_SOC_MS001 VDD Define                                          -1       625.0000       684.3750           731.2500       0.0000         0       
            for (lnIdx = 0; lnIdx < binCutLineBases.Count; lnIdx++)
            {
                string line = binCutLineBases[lnIdx].Line;
                int lineNo = binCutLineBases[lnIdx].LineNo;

                string[] spt = line.Split([" ", " K "], StringSplitOptions.RemoveEmptyEntries);
                _ = int.TryParse(spt[1], out int site);
                string mode = spt[2];

                int regxIdx = -1;
                for (int i = 0; i < spt.Length; i++)
                {
                    Match matchObj = Reg.RegxValue1.Match(spt[i]);
                    if (matchObj.Length != 0)
                    {
                        regxIdx = i;
                        break;
                    }
                }

                int moreTwoStep = 0;
                int step = 0;
                for (int i = regxIdx + 1; i < spt.Length; i++)
                {
                    if (double.TryParse(spt[i], out double _))
                    {
                        moreTwoStep++;
                        if (moreTwoStep == 1)
                        {
                            step = i;
                            break;
                        }
                    }
                }

                bool isFoundK = false;
                //from log
                double dVal = double.Parse(spt[step]);
                if (dVal < 5.0 && dVal > 0)
                {
                    //*1000 accuracy will lost, ex: dVal = 1.001, after multipule by 1000.0, dVal = 1000.9999. for this, add DOUBLE_LSB
                    dVal *= 1000;
                    //if found K, only count on integer part
                    dVal = (int)(dVal + doubleLsb);
                    isFoundK = true;
                }

                //efuse value in CFGFuse GetReadValue
                double expectEfuseVal = 0.0;
                double expectEfuseGb = 0.0;
                EFuseRow? efuse = siteInfoArray[site].EFuseValues.Find(x => mode.Contains(x.Name));
                if (efuse != null)
                {
                    expectEfuseVal = efuse.Value;
                    expectEfuseGb = efuse.Gb;
                }
                bool isError = false;
                double expectDVal = 0.0;
                if (spt[3][0] == 'C' || spt[3][0] == 'F' || spt[3][0] == 'Q') //'C' for efuse - GB 
                {
                    //expect value
                    expectDVal = expectEfuseVal - expectEfuseGb;
                    if (isFoundK)
                    {
                        if (Math.Abs(dVal - (int)expectDVal) > doubleLsb)
                        {
                            isError = true;
                        }
                    }
                    else
                        if (Math.Abs(dVal - expectDVal) > doubleLsb)
                    {
                        isError = true;
                    }
                }
                else
                {
                    expectDVal = expectEfuseVal;
                    if (isFoundK)
                    {
                        if (Math.Abs(dVal - (int)expectDVal) > doubleLsb)
                        {
                            isError = true;
                        }
                    }
                    else
                        if (Math.Abs(dVal - expectDVal) > doubleLsb)
                    {
                        isError = true;
                    }
                }

                if (isError || BinCutConfig.IsDebugPrint)
                {
                    if (isError)
                    {
                        BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{dVal:F4}", $"{expectDVal:F4}", "<PrintOutVddBinning>");
                    }
                    else
                    {
                        BinCutPrint.PrintSame(streamWriter, siteInfoArray, lineNo.ToString(), site, $"{dVal:F4}", $"{expectDVal:F4}", "<PrintOutVddBinning>");
                    }
                }
            }
        }

        public static int GetChannelIndex(string[] spt)
        {
            int channelIndex = 0;
            for (int i = 0; i < spt.Length; i++)
            {
                if (RegxChannel.IsMatch(spt[i]))
                {
                    channelIndex = i;
                    break;
                }
            }
            return channelIndex;
        }

        private static int GetMeasureIndex(int channelIndex, string[] spt)
        {
            int step = 0;
            int moreTwoStep = 0;
            for (int i = channelIndex + 1; i < spt.Length; i++)
            {
                if (double.TryParse(spt[i], out _))
                {
                    moreTwoStep++;
                    if (moreTwoStep == 2)
                    {
                        step = i;
                        break;
                    }
                }
            }
            return step;
        }
    }
}
