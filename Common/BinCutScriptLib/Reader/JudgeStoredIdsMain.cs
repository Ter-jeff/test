using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib.BinCut.BinCutConfig;
using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib.Reader
{
    internal class JudgeStoredIdsMain
    {
        internal class IdsMax
        {
            public string Name = string.Empty;
            public double IdsMaxValue;
            public string Domain = string.Empty;
        }
        public static bool CheckJudgeStoredIds(StreamWriter streamWriter, EnumJob enumJob, ref OneTouchDown oneTouchDown, ref SiteInfo[] siteInfoArray, bool isFoundSetWriteDecimal)
        {
            List<List<IdsMax>> allIdsSpec = GetIdsMaxFromTable();

            bool flag = CheckJudgeStoredIdsAndCreatePowers(streamWriter, enumJob, oneTouchDown, siteInfoArray, isFoundSetWriteDecimal, allIdsSpec);

            CheckIdsMax(ref siteInfoArray, allIdsSpec);

            return flag;
        }

        public static bool CheckJudgeStoredIdsCs(StreamWriter streamWriter, EnumJob enumJob, ref OneTouchDown oneTouchDown, ref SiteInfo[] siteInfoArray, bool isFoundSetWriteDecimal, List<PinInfo> pinInfos)
        {
            List<List<IdsMax>> allIdsSpec = GetIdsMaxFromTableCs();

            bool flag = CheckJudgeStoredIdsAndCreatePowersCs(streamWriter, enumJob, oneTouchDown, siteInfoArray, isFoundSetWriteDecimal, allIdsSpec, pinInfos);

            CheckIdsMax(ref siteInfoArray, allIdsSpec);

            return flag;
        }

        private static bool CheckJudgeStoredIdsAndCreatePowers(StreamWriter streamWriter, EnumJob enumJob, OneTouchDown oneTouchDown, SiteInfo[] siteInfoArray, bool isFoundSetWriteDecimal, List<List<IdsMax>> allIdsSpec)
        {
            //<Judge_stored_IDS>
            //24287900     0     VDD_PCPU0_MP001 BinCut1 IDS                                                                                    VDD_PCPU0                                8.e101   0.2000 mA      142.8000 mA          425.4000 mA    0.0000 A       0       
            //24287901     0     VDD_PCPU0_MP001 BinCut2 IDS                                                                                    VDD_PCPU0                                8.e101   0.2000 mA      142.8000 mA          553.0000 mA    0.0000 A       0       
            bool isFoundIds = false;
            int oneTouchIndex;
            bool isIdsHotStart = false;
            bool isIdsHotEnd = false;
            for (oneTouchIndex = 0; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("<Judge_stored_IDS>"))
                {
                    isFoundIds = true;
                    break;
                }
            }
            if (!isFoundIds)
            {
                return false;
            }

            oneTouchIndex++;
            for (; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                string[] spt = oneTouchDown.Lines[oneTouchIndex].Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                int regxIdx = -1;
                if (oneTouchDown.Lines[oneTouchIndex].Line == "------Check IDS Hot Start------")
                {
                    isIdsHotStart = true;
                    continue;
                }
                if (oneTouchDown.Lines[oneTouchIndex].Line == "------Check IDS Hot End------")
                {
                    isIdsHotEnd = true;
                    continue;
                }
                if (isIdsHotStart && !isIdsHotEnd)
                {
                    if (spt.Length >= 11)
                    {
                        if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("(F)"))
                        {
                            siteInfoArray[int.Parse(spt[1])].IsIdsHotBinOut = true;
                        }
                    }
                }
                for (int i = 0; i < spt.Length; i++)
                {
                    Match matchObj = Reg.RegxChannel.Match(spt[i]);
                    if (matchObj.Length != 0)
                    {
                        regxIdx = i;
                        break;
                    }
                    if (spt[i] == "-1")
                    {
                        regxIdx = i;
                        break;
                    }
                }

                if (regxIdx == -1)
                {
                    //ex: 1580101 1 Check IDS Error -1 1.000 999.000    (F)    1.000   0.000   0
                    if ((spt.Length > 2 && spt[1].Length == 1) ||
                        oneTouchDown.Lines[oneTouchIndex].Line.Contains("ERROR", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    break;
                }

                //!!ids fail line, this is vary important, once ids fail, all performance power will be remove later and this site should be disabled!
                if (!BinCutConfig.IsDoAll)
                {
                    if (oneTouchDown.Lines[oneTouchIndex].Line.Contains("(F)"))
                    {
                        continue;
                    }
                }

                CheckJudgeStoredIds(streamWriter, enumJob, oneTouchDown, siteInfoArray, isFoundSetWriteDecimal, allIdsSpec, oneTouchIndex, spt, regxIdx);
            }
            return true;
        }

        private static void CheckJudgeStoredIds(StreamWriter streamWriter, EnumJob enumJob, OneTouchDown oneTouchDown, SiteInfo[] siteInfoArray, bool isFoundSetWriteDecimal, List<List<IdsMax>> allIdsSpec, int oneTouchIndex, string[] spt, int regxIdx)
        {
            //Create performance power and IDS value
            int site = int.Parse(spt[1]);
            string pinMode = spt[2].ToUpper();
            string mode = Reg.RegexRegexPerformance.IsMatch(pinMode) ? Reg.RegexRegexPerformance.Match(pinMode).Groups["pmode"].ToString() : "";
            string binCutBin = spt[3];

            var powerTmp = new PowerZone
            {
                PinMode = pinMode,
                Mode = mode
            };
            if (string.IsNullOrEmpty(mode))
            {
                powerTmp.Pin = pinMode;
            }
            else
            {
                powerTmp.Pin = powerTmp.PinMode.Replace("_" + powerTmp.Mode, "");
            }

            int moreTwoStep = 0;
            int step = 0;
            // todo jeff
            double value;
            for (int i = regxIdx + 1; i < spt.Length; i++)
            {
                if (double.TryParse(spt[i], out value))
                {
                    moreTwoStep++;
                    if (moreTwoStep == 2)
                    {
                        step = i;
                        break;
                    }
                }
            }

            if (step != 0)
            {
                if (spt[step] != null && double.TryParse(spt[step], out value))
                {
                    powerTmp.IdsValue = value;
                }
                else
                {
                    BinCutPrint.PrintCommomError(streamWriter, oneTouchDown.Lines[oneTouchIndex], "IDS value should not be empty");
                }
            }
            string unitTmp = spt[regxIdx + 4].ToLower();
            if (unitTmp == "ua")
            {
                powerTmp.IdsValue /= 1000.0;
            }

            powerTmp.IdsValueReal = powerTmp.IdsValue;

            ChekHighLimit(streamWriter, oneTouchDown, siteInfoArray, allIdsSpec, oneTouchIndex, spt, regxIdx, site, pinMode, mode, binCutBin);

            if (isFoundSetWriteDecimal)
            {
                HandleSetWriteDecimal(streamWriter, enumJob, oneTouchDown, siteInfoArray, oneTouchIndex, site, powerTmp);
            }

            if (siteInfoArray[site].AllPowers.Exists(x => x.PinMode.EqualsIgnoreCase(powerTmp.PinMode)))
            {
                powerTmp.IdsValuePowerBinning = powerTmp.IdsValue;
                for (int index = 0; index < siteInfoArray[site].AllPowers.Count; index++)
                {
                    if (siteInfoArray[site].AllPowers[index].PinMode.EqualsIgnoreCase(powerTmp.PinMode))
                    {
                        siteInfoArray[site].AllPowers[index] = powerTmp;
                    }
                }
            }
            else
            {
                if (siteInfoArray[site].AllPowers.Exists(x => x.Pin.EqualsIgnoreCase(powerTmp.Pin)))
                {
                    //for Sram power
                    powerTmp.IdsValuePowerBinning = powerTmp.IdsValue;
                    var powers = siteInfoArray[site].AllPowers.Where(x => x.Pin.EqualsIgnoreCase(powerTmp.Pin)).ToList();
                    for (int index = 0; index < powers.Count; index++)
                    {
                        powers[index] = powerTmp.Copy();
                    }
                }
            }
        }

        private static void HandleSetWriteDecimal(StreamWriter streamWriter, EnumJob enumJob, OneTouchDown oneTouchDown, SiteInfo[] siteInfoArray, int oneTouchIndex, int site, PowerZone powerZone)
        {
            string powerName = powerZone.Mode.Length != 0 ? powerZone.PinMode.Replace("_" + powerZone.Mode, "") : powerZone.PinMode;
            string idsName = GetIdsName(powerName, siteInfoArray[site].HarvesFlags);
            IdsData? oneRealIds = siteInfoArray[site].RealIds.Find(x => x.IdsName == idsName);
            if (oneRealIds != null)
            {
                if (enumJob != EnumJob.CP1) //non-cp1, ignore compare judge stroe IDS. Check non-empty only
                {
                }
                else if (powerZone.IdsValue != oneRealIds.EfuseVal && oneRealIds.EfuseVal != 0)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, oneTouchDown.Lines[oneTouchIndex].LineNo.ToString(), site, powerZone.IdsValue.ToString(), oneRealIds.EfuseVal.ToString(), "Judge_stored_IDS - " + powerZone.PinMode + " - " + oneRealIds.IdsType);
                    oneRealIds.EfuseVal = powerZone.IdsValue;
                    oneRealIds.RealVal = powerZone.IdsValue;
                }
            }
            else if (enumJob == EnumJob.CP1)
            {
                BinCutPrint.PrintCommomError(streamWriter, oneTouchDown.Lines[oneTouchIndex], " The data of IDS_" + powerName + " can not be found !!!");
            }
        }

        private static void ChekHighLimit(StreamWriter streamWriter, OneTouchDown oneTouchDown, SiteInfo[] siteInfoArray, List<List<IdsMax>> allIdsSpec, int oneTouchIndex, string[] spt, int regxIdx, int site, string pinMode, string mode, string binCutBin)
        {
            #region check high limit
            string limitHigh = spt[regxIdx + 5];
            if (limitHigh == "(F)")
            {
                limitHigh = spt[regxIdx + 6];
            }

            _ = double.TryParse(limitHigh, out double limitHighvalue);
            int binIndex = GetBinIndex(binCutBin);
            string name = string.IsNullOrEmpty(mode) ? pinMode : mode[..4];
            if (allIdsSpec[binIndex].Exists(x => x.Name.EqualsIgnoreCase(name)))
            {
                double expectlimitHigh = allIdsSpec[binIndex].Find(x => x.Name.EqualsIgnoreCase(name))!.IdsMaxValue;
                if (Math.Abs(expectlimitHigh - limitHighvalue) > 0.002)
                {
                    BinCutPrint.PrintDifference(streamWriter, siteInfoArray, oneTouchDown.Lines[oneTouchIndex].LineNo.ToString(), site, limitHighvalue.ToString(), expectlimitHigh.ToString(), "Judge_stored_IDS - Limit high Mismatch");
                }
            }
            else
            {
                BinCutPrint.PrintCommomError(streamWriter, oneTouchDown.Lines[oneTouchIndex], $"Can not find high limit for {pinMode}!!!");
            }
            #endregion
        }

        private static bool CheckJudgeStoredIdsAndCreatePowersCs(StreamWriter streamWriter, EnumJob enumJob, OneTouchDown oneTouchDown, SiteInfo[] siteInfoArray, bool isFoundSetWriteDecimal, List<List<IdsMax>> allIdsSpec, List<PinInfo> pinInfos)
        {
            //<Judge_stored_IDS_Csharp>
            //41799000     0     VDD_PCPU Bincut1 IDS VDD_PCPU                       7.e201   0.00 mA        174.40 mA            655.80 mA      0.00           0
            //41799000     1     VDD_PCPU Bincut1 IDS VDD_PCPU                       7.e401   0.00 mA        168.00 mA            655.80 mA      0.00           0
            //41799000     2     VDD_PCPU Bincut1 IDS VDD_PCPU                       6.e201   0.00 mA        135.00 mA            655.80 mA      0.00           0
            //41799000     3     VDD_PCPU Bincut1 IDS VDD_PCPU                       6.e401   0.00 mA        139.40 mA            655.80 mA      0.00           0
            //41799001     0     VDD_ECPU Bincut1 IDS VDD_ECPU                       11.e201  0.00 mA        51.20 mA             235.00 mA      0.00           0
            //41799001     1     VDD_ECPU Bincut1 IDS VDD_ECPU                       11.e401  0.00 mA        50.20 mA             235.00 mA      0.00           0
            //41799001     2     VDD_ECPU Bincut1 IDS VDD_ECPU                       10.e201  0.00 mA        41.00 mA             235.00 mA      0.00           0
            //41799001     3     VDD_ECPU Bincut1 IDS VDD_ECPU                       10.e412  0.00 mA        45.80 mA             235.00 mA      0.00           0

            bool isFoundIds = false;
            var rows = new List<JudgestoredIdsRow>();
            int i;
            int startIndex = -1;
            int idsEnd = oneTouchDown.Lines.FindIndex(x => x.Line.StartsWith("print: Judge_stored_IDS end"));
            i = CheckJudge_stored_IDS(oneTouchDown, ref isFoundIds, ref rows, ref startIndex, idsEnd);

            var groups = pinInfos.GroupBy(x => x.Pin).ToList();
            var judgestoredIdsRows = new List<JudgestoredIdsRow>();
            foreach (JudgestoredIdsRow row in rows)
            {
                int site = row.Site;
                _ = double.TryParse(row.Measured, out double idsValue);
                string binCutBin = row.TestBinCutBin;
                string pinMode = row.TestName;
                string domain = pinMode.Replace("VDD_", "");

                ChekHighLimitCs(streamWriter, siteInfoArray, allIdsSpec, row, site, binCutBin, pinMode, domain);

                if (isFoundSetWriteDecimal)
                {
                    SetWriteDecimalCs(streamWriter, enumJob, siteInfoArray, row, site, idsValue, pinMode);
                }

                string pin = row.Pin;
                IGrouping<string, PinInfo>? find = groups.Find(x => x.Key.EqualsIgnoreCase(pin));
                if (find != null)
                {
                    foreach (PinInfo item in find)
                    {
                        JudgestoredIdsRow judgestoredIdsRow = row.Copy();
                        row.TestName = item.Pin.Contains(item.Mode) ? item.Pin : item.Pin + "_" + item.Mode;
                        judgestoredIdsRow.TestName = row.TestName;
                        judgestoredIdsRow.Line.Line = row.Print();
                        judgestoredIdsRows.Add(judgestoredIdsRow);
                        row.TestName = "";
                    }
                }
            }

            if (BinCutConfig.IsDebugPrint)
            {
                streamWriter.WriteLine("");
                streamWriter.Flush();
            }

            if (!isFoundIds)
            {
                return false;
            }

            CompareJudge_stored_IDS(streamWriter, siteInfoArray, judgestoredIdsRows);
            return true;
        }

        private static void CompareJudge_stored_IDS(StreamWriter streamWriter, SiteInfo[] siteInfoArray, List<JudgestoredIdsRow> judgestoredIdsRows)
        {
            for (int j = 0; j < judgestoredIdsRows.Count; j++)
            {
                string[] spt = judgestoredIdsRows[j].Line.Line.Split(['\t'], StringSplitOptions.RemoveEmptyEntries);

                int regxIdx = -1;
                //find eg. 7.e201
                for (int k = 0; k < spt.Length; k++)
                {
                    Match matchObj = Reg.RegxChannel.Match(spt[k]);
                    if (matchObj.Length != 0)
                    {
                        regxIdx = k;
                        break;
                    }
                    if (spt[k] == "-1")
                    {
                        regxIdx = k;
                        break;
                    }
                }

                if (regxIdx == -1)
                {
                    //ex: 1580101 1 Check IDS Error -1 1.000 999.000    (F)    1.000   0.000   0
                    if ((spt.Length > 2 && spt[1].Length == 1) ||
                        judgestoredIdsRows[j].Line.Line.Contains("ERROR", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    break;
                }

                //!!ids fail line, this is vary important, once ids fail, all performance power will be remove later and this site should be disabled!
                if (!BinCutConfig.IsDoAll)
                {
                    if (judgestoredIdsRows[j].Line.Line.Contains("(F)"))
                    {
                        continue;
                    }
                }

                int site = int.Parse(spt[1]);
                string pinMode = spt[2].ToUpper();
                string mode = Reg.RegexRegexPerformance.IsMatch(pinMode) ?
                    Reg.RegexRegexPerformance.Match(pinMode).Groups["pmode"].ToString() : "";

                var powerTmp = new PowerZone
                {
                    PinMode = pinMode,
                    Mode = mode
                };
                if (string.IsNullOrEmpty(mode))
                {
                    powerTmp.Pin = pinMode;
                }
                else
                {
                    powerTmp.Pin = powerTmp.PinMode.Replace("_" + powerTmp.Mode, "");
                }

                int moreTwoStep = 0;
                int step = 0;
                double value;
                for (int k = regxIdx + 1; k < spt.Length; k++)
                {
                    if (double.TryParse(spt[k], out value))
                    {
                        moreTwoStep++;
                        if (moreTwoStep == 2)
                        {
                            step = k;
                            break;
                        }
                    }
                }

                if (step != 0)
                {
                    if (spt[step] != null && double.TryParse(spt[step], out value))
                    {
                        powerTmp.IdsValue = value;
                    }
                    else
                    {
                        BinCutPrint.PrintCommomErrorForIdsCsharp(streamWriter, judgestoredIdsRows[j].Line, "IDS value should not be empty");
                    }
                }
                string unitTmp = spt[regxIdx + 4].ToLower();
                if (unitTmp == "ua")
                {
                    powerTmp.IdsValue /= 1000.0;
                }

                powerTmp.IdsValueReal = powerTmp.IdsValue;

                if (siteInfoArray[site].AllPowers.Exists(x => x.PinMode.EqualsIgnoreCase(powerTmp.PinMode)))
                {
                    powerTmp.IdsValuePowerBinning = powerTmp.IdsValue;
                    for (int index = 0; index < siteInfoArray[site].AllPowers.Count; index++)
                    {
                        if (siteInfoArray[site].AllPowers[index].PinMode.EqualsIgnoreCase(powerTmp.PinMode))
                        {
                            siteInfoArray[site].AllPowers[index] = powerTmp;
                        }
                    }
                }
                else
                {
                    if (siteInfoArray[site].AllPowers.Exists(x => x.Pin.EqualsIgnoreCase(powerTmp.Pin)))
                    {
                        //for Sram power
                        powerTmp.IdsValuePowerBinning = powerTmp.IdsValue;
                        var powers = siteInfoArray[site].AllPowers.Where(x => x.Pin.EqualsIgnoreCase(powerTmp.Pin)).ToList();
                        for (int index = 0; index < powers.Count; index++)
                        {
                            powers[index] = powerTmp.Copy();
                        }
                    }
                }
            }
        }

        private static int CheckJudge_stored_IDS(OneTouchDown oneTouchDown, ref bool isFoundIds, ref List<JudgestoredIdsRow> judgestoredIdsRows, ref int startIndex, int idsEnd)
        {
            int i;
            for (i = 0; i < oneTouchDown.Lines.Count; i++)
            {
                if (oneTouchDown.Lines[i].Line.Contains("<Judge_stored_IDS"))
                {
                    isFoundIds = true;
                    startIndex = i;
                }
                else if (startIndex != -1)
                {
                    if (idsEnd != -1)
                    {
                        if (oneTouchDown.Lines[i].Line.Contains("print: Judge_stored_IDS end"))
                        {
                            List<BinCutLineBase> lines = oneTouchDown.Lines.GetRange(startIndex + 1, i - startIndex - 1);
                            var reader = new JudgestoredIdsReader(lines);
                            judgestoredIdsRows = reader.JudgestoredIdsRows;
                            break;
                        }
                    }
                    else
                    {
                        string[] arr = oneTouchDown.Lines[i].Line.Trim().Split(' ');
                        if (oneTouchDown.Lines[i].Line.Contains("******************************") || !long.TryParse(arr.First(), out long _) || oneTouchDown.Lines[i].Line.StartsWithIgnoreCase("Flow "))
                        {
                            List<BinCutLineBase> lines = oneTouchDown.Lines.GetRange(startIndex + 1, i - startIndex - 1);
                            var reader = new JudgestoredIdsReader(lines);
                            judgestoredIdsRows = reader.JudgestoredIdsRows;
                            break;
                        }
                    }
                }
            }

            return i;
        }

        private static void SetWriteDecimalCs(StreamWriter streamWriter, EnumJob enumJob, SiteInfo[] siteInfoArray, JudgestoredIdsRow judgestoredIdsRow, int site, double idsValue, string pinMode)
        {
            string powerName = pinMode;
            string idsName = GetIdsName(powerName, siteInfoArray[site].HarvesFlags);
            IdsData? oneRealIds = siteInfoArray[site].RealIds.Find(x => x.IdsName == idsName);
            if (oneRealIds != null && enumJob == EnumJob.CP1)
            {
                bool fail = idsValue != oneRealIds.EfuseVal && oneRealIds.EfuseVal != 0;
                if (fail || BinCutConfig.IsDebugPrint)
                {
                    if (fail)
                    {
                        BinCutPrint.PrintDifferenceCsharp(streamWriter, siteInfoArray, judgestoredIdsRow.Line, site, idsValue.ToString(), oneRealIds.EfuseVal.ToString(), "Judge_stored_IDS - " + idsValue + " - " + oneRealIds.IdsType);
                    }
                    else
                    {
                        BinCutPrint.PrintSameIsFoundSetWriteDecimalCsharp(streamWriter, judgestoredIdsRow.Line);
                    }
                }
            }
            else if (enumJob == EnumJob.CP1)
            {
                BinCutPrint.PrintCommomError(streamWriter, judgestoredIdsRow.Line, " The data of IDS_" + powerName + " can not be found !!!");
            }
        }

        private static void ChekHighLimitCs(StreamWriter streamWriter, SiteInfo[] siteInfoArray, List<List<IdsMax>> allIdsSpec, JudgestoredIdsRow judgestoredIdsRow, int site, string binCutBin, string pinMode, string domain)
        {
            #region check high limit
            string limitHigh = judgestoredIdsRow.High;
            _ = double.TryParse(limitHigh, out double limitHighvalue);
            int binIndex = GetBinIndexCsharp(binCutBin);
            if (allIdsSpec[binIndex].Exists(x => x.Domain.EqualsIgnoreCase(domain)))
            {
                double expectlimitHigh = allIdsSpec[binIndex].Find(x => x.Domain.EqualsIgnoreCase(domain))!.IdsMaxValue;
                if (Math.Abs(expectlimitHigh - limitHighvalue) > 0.002)
                {
                    BinCutPrint.PrintDifferenceCsharp(streamWriter, siteInfoArray, judgestoredIdsRow.Line, site, limitHighvalue.ToString(), expectlimitHigh.ToString(), "Judge_stored_IDS - Limit high Mismatch");
                }
            }
            else
            {
                BinCutPrint.PrintCommomError(streamWriter, judgestoredIdsRow.Line, $"Can not find high limit for {pinMode}!!!");
            }
            #endregion
        }

        private static int GetBinIndex(string binCutBin)
        {
            int binIndex = 0;
            if (binCutBin == "BinCut1")
            {
                binIndex = 0;
            }
            else if (binCutBin == "BinCut2")
            {
                binIndex = 1;
            }
            else if (binCutBin == "BinCut3")
            {
                binIndex = 2;
            }

            return binIndex;
        }

        private static int GetBinIndexCsharp(string binCutBin)
        {
            int binIndex = 0;
            if (binCutBin.EqualsIgnoreCase("Bincut1"))
            {
                binIndex = 0;
            }
            else if (binCutBin.EqualsIgnoreCase("Bincut2"))
            {
                binIndex = 1;
            }
            else if (binCutBin.EqualsIgnoreCase("Bincut3"))
            {
                binIndex = 2;
            }

            return binIndex;
        }

        private static string GetIdsName(string powerName, Dictionary<string, string> harvesFlags)
        {
            string idsName = "";
            idsName = GetIdsNameWithHarv(powerName, harvesFlags);
            if (!string.IsNullOrEmpty(idsName))
            {
                return idsName;
            }

            if (BinCutConfig.IdsNames.Exists(x => x.Item2.EqualsIgnoreCase(powerName)))
            {
                return BinCutConfig.IdsNames.Find(x => x.Item2.EqualsIgnoreCase(powerName))!.Item1;
            }

            if (BinCutConfig.PowerType.ContainsKey(powerName))
            {
                return "IDS_" + powerName.ToUpper();
            }

            return idsName;
        }

        private static string GetIdsNameWithHarv(string powerName, Dictionary<string, string> harvesFlags)
        {
            return harvesFlags == null ? "" : BinCutConfig.ProjectConfig.GetIdsNameWithHarv(powerName, harvesFlags);
        }

        private static void CheckIdsMax(ref SiteInfo[] siteInfoArray, List<List<IdsMax>> allIdsSpec)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].AllPowers.Count == 0)
                {
                    continue;
                }

                int bin = -1;
                for (int binIdx = 0; binIdx < allIdsSpec.Count; binIdx++)
                {
                    bool isAllIdsPass = true;
                    for (int pwrIdx = 0; pwrIdx < siteInfoArray[site].AllPowers.Count; pwrIdx++)
                    {
                        for (int idsIdx = 0; idsIdx < allIdsSpec[binIdx].Count; idsIdx++)
                        {
                            if (siteInfoArray[site].AllPowers[pwrIdx].PinMode.Contains(allIdsSpec[binIdx][idsIdx].Name))
                            {
                                string pwrNameTmp = siteInfoArray[site].AllPowers[pwrIdx].PinMode;
                                if (siteInfoArray[site].AllPowers[pwrIdx].IdsValueReal >= allIdsSpec[binIdx][idsIdx].IdsMaxValue)
                                {
                                    isAllIdsPass = false;
                                    if (siteInfoArray[site].AllPowers[pwrIdx].Bin > BinCutData.BinningTables.Count - 1)
                                    {
                                        siteInfoArray[site].AllPowers[pwrIdx].IsBinOut = true;
                                    }
                                    else
                                    {
                                        siteInfoArray[site].AllPowers[pwrIdx].Bin = binIdx + 1 + 1;
                                    }

                                    if (binIdx == allIdsSpec.Count - 1)
                                    {
                                        //string errorMessage = $"The Ids {siteInfos[site].AllPowers[pwrIdx].IdsValueReal} of {pwrNameTmp} was larger than the spec {allIdsSpec[binIdx][idsIdx].IdsMaxValue}";
                                        string stie = $"X:{siteInfoArray[site].XCoor} Y:{siteInfoArray[site].YCoor}";
                                        ErrorReportManager.AddError(BinCutErrorType.E_Business_01, "Datalog", 0, 0, $"The Ids {siteInfoArray[site].AllPowers[pwrIdx].IdsValueReal} of {pwrNameTmp} was larger than the spec {allIdsSpec[binIdx][idsIdx].IdsMaxValue}" + " in " + stie + " @ Bin " + binIdx, [siteInfoArray[site].AllPowers[pwrIdx].IdsValueReal.ToString(), pwrNameTmp, allIdsSpec[binIdx][idsIdx].IdsMaxValue.ToString(), stie, binIdx.ToString()]);
                                    }
                                    break;
                                }
                            }
                        }
                    }

                    if (isAllIdsPass && bin == -1)
                    {
                        bin = binIdx + 1;
                    }
                }

                //2017/01/30: 原先int bin預設為0, 當沒有找到IDS時, 結果竟然是Bin1, 完全不合Logic. 將int bin預設為-1, 沒找到pass的區段就將其設為fail site
                if (bin == -1) //表示從頭到尾都沒有找到IDS pass的Bin cut num, 
                {
                    if (!BinCutConfig.IsDoAll)
                    {
                        siteInfoArray[site].AllPowers.Clear();
                    }

                    continue;
                }

                // 如果"實際"IDS量測結果為Bin2, 且Fail只有Main Power, 重覆以round down後之IDS重覆進行檢查, 如果Pass就歸頪為Bin1
                //2016/12/01: 原因是 Jerry VBT在Judge_stored_IDS時是以實際量測值進行Bin1/Bin2檢查, 
                //但在排Zone時則是用Round Down後的值進行GradeSearch, 造成VBT與Script不Match
                //這段Code可以說是非常Dirty
                if (bin == 2)
                {
                    bin = HandleBin2(siteInfoArray, allIdsSpec, site, bin);
                }
                else if (bin == 3)
                {
                    bin = HandleBin3(siteInfoArray, allIdsSpec, site, bin);
                }

                siteInfoArray[site].Bin = bin;
            }
        }

        private static int HandleBin3(SiteInfo[] siteInfoArray, List<List<IdsMax>> allIdsSpec, int site, int bin)
        {
            // Re-check target bin 1 (index 1). If all passes, drop bin down to 2.
            return CheckIdsPassForBinIndex(siteInfoArray, allIdsSpec, site, reCheckBinIndex: 1) ? 2 : bin;
        }

        private static int HandleBin2(SiteInfo[] siteInfoArray, List<List<IdsMax>> allIdsSpec, int site, int bin)
        {
            // Re-check target bin 0 (index 0). If all passes, drop bin down to 1.
            return CheckIdsPassForBinIndex(siteInfoArray, allIdsSpec, site, reCheckBinIndex: 0) ? 1 : bin;
        }

        private static bool CheckIdsPassForBinIndex(SiteInfo[] siteInfoArray, List<List<IdsMax>> allIdsSpec, int site, int reCheckBinIndex)
        {
            // Defensive check to avoid out of bounds exceptions on empty spec list structures
            if (allIdsSpec == null || reCheckBinIndex >= allIdsSpec.Count || siteInfoArray == null || site < 0 || site >= siteInfoArray.Length)
            {
                return false;
            }

            List<PowerZone> currentPowers = siteInfoArray[site].AllPowers;
            List<IdsMax> targetSpecs = allIdsSpec[reCheckBinIndex];

            for (int pwrIdx = 0; pwrIdx < currentPowers.Count; pwrIdx++)
            {
                PowerZone powerItem = currentPowers[pwrIdx];
                string pwrNameTmp = powerItem.PinMode;

                for (int idsIdx = 0; idsIdx < targetSpecs.Count; idsIdx++)
                {
                    IdsMax specItem = targetSpecs[idsIdx];

                    if (pwrNameTmp != null && pwrNameTmp.Contains(specItem.Name))
                    {
                        EnumPowerType pwrNameType = BinCutAlgorithmService.GetTypeByPowerName(pwrNameTmp);
                        double actualValue = (pwrNameType == EnumPowerType.Others) ? powerItem.IdsValueReal : powerItem.IdsValue;

                        // If any value hits or exceeds max limit thresholds, immediately fail validation pass
                        if (actualValue >= specItem.IdsMaxValue)
                        {
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        private static List<List<IdsMax>> GetIdsMaxFromTable()
        {
            var allIdsSpec = new List<List<IdsMax>>();
            for (int binTbIdx = 0; binTbIdx < BinCutData.BinningTables.Count; binTbIdx++)
            //run all bin table iteration, eg. bin1->bin2....
            {
                //oneIdsSpec for one binning ids value
                var oneIdsSpec = new List<IdsMax>();

                //STEP1. Get IDS from "VddBinTable" (main power: SOC/CPU/GPU)
                BinningTable binningTable = BinCutData.BinningTables[binTbIdx];
                for (int tbRowIdx = 0; tbRowIdx < binningTable.Rows.Count; tbRowIdx++)
                //all rows in binTable, eg. MC601/MC602/MC603...
                {
                    //ex: MC601
                    string tbModeName = binningTable.Rows[tbRowIdx].RowData[binningTable.ModeIdx];
                    string modeType = tbModeName[..4];
                    double idsMaxValue = double.Parse(binningTable.Rows[tbRowIdx].RowData[binningTable.BinXIdsMaxIdx != -1 ? binningTable.BinXIdsMaxIdx : binningTable.IdsMaxIdx]);

                    //STEP1a. 先找後面有無相同P.Power, 有即代表同Power不同EQN, 以所有EQN最大的值填入allIdsVal
                    if (tbRowIdx != binningTable.Rows.Count - 1)
                    {
                        for (int searchEqnIdx = tbRowIdx + 1; searchEqnIdx < binningTable.Rows.Count; searchEqnIdx++)
                        {
                            //ex: MC601
                            string nextModeName = binningTable.Rows[searchEqnIdx].RowData[binningTable.ModeIdx];
                            double nextIdsMaxValue = double.Parse(binningTable.Rows[searchEqnIdx].RowData[binningTable.BinXIdsMaxIdx != -1 ? binningTable.BinXIdsMaxIdx : binningTable.IdsMaxIdx]);
                            if (nextModeName == tbModeName)
                            {
                                if (nextIdsMaxValue > idsMaxValue)
                                {
                                    idsMaxValue = nextIdsMaxValue;
                                }

                                tbRowIdx = searchEqnIdx;
                            }
                            else
                            {
                                break;
                            }
                        }
                    }

                    //STEP1b. 將找到的Performance type在allIdsVal裡搜尋, 若有且比記錄的小, 則以小的值覆蓋
                    bool isFoundSamePMode = false;
                    foreach (IdsMax onePower in oneIdsSpec)
                    {
                        if (modeType == onePower.Name) //MC/MS/MG
                        {
                            isFoundSamePMode = true;
                            if (onePower.IdsMaxValue > idsMaxValue) //取其小
                            {
                                onePower.IdsMaxValue = idsMaxValue;
                            }
                        }
                    }

                    if (!isFoundSamePMode) //該Power第一次被加入, 直接加入List
                    {
                        var tmp = new IdsMax { Name = modeType, IdsMaxValue = idsMaxValue };
                        oneIdsSpec.Add(tmp);
                    }
                }

                //STEP2. 2 (ohter power: CPUSRAM/GPUSRAM/FIX/LOW)
                BinningTable othTbRef = BinCutData.OtherRailTables[binTbIdx];
                for (int tbRowIdx = 0; tbRowIdx < othTbRef.Rows.Count; tbRowIdx++)
                //all rows in binTable, eg. MC601/MC602/MC603...
                {
                    //ex: CPUSRAM
                    string tbDomainName = othTbRef.Rows[tbRowIdx].RowData[othTbRef.ModeIdx];
                    string tbDomainType = "";
                    if (BinCutConfig.PowerType.ContainsKey(tbDomainName.ToUpper()))
                    {
                        tbDomainType = tbDomainName;
                    }
                    else
                    {
                        tbDomainName = othTbRef.Rows[tbRowIdx].RowData[othTbRef.DomainIdx];
                        if (BinCutConfig.PowerType.ContainsKey(tbDomainName.ToUpper()))
                        {
                            tbDomainType = tbDomainName;
                        }
                        else
                        {
                            if (BinCutConfig.DomainInOtherRail2Power.TryGetValue(tbDomainName, out string? value))
                            {
                                tbDomainType = value;
                            }
                            else
                            {
                                if (BinCutConfig.PowerType.ContainsKey("VDD_" + tbDomainName.ToUpper()))
                                {
                                    tbDomainType = "VDD_" + tbDomainName;
                                }
                                else
                                {
                                    string errorMessage =
                                        string.Format("The domain : " + tbDomainName + " in Other Rails can't be found.");
                                    if (!ErrorReportManager.GetErrorList().Select(x => x.Message).Contains(errorMessage))
                                    {
                                        ErrorReportManager.AddError(BinCutErrorType.E_Business_02, "Datalog", 0, 0, string.Format("The domain : " + tbDomainName + " in Other Rails can't be found."), [tbDomainName]);
                                    }
                                }
                            }
                        }
                    }

                    if (tbDomainType.Length == 0)
                    {
                        string ignoreMessage =
                            string.Format("The domain : " + tbDomainName + " is IGNORE COLUMN in Non_binning_rail sheet.");
                        BinCutController.Controller.RichTextBoxAppend(ignoreMessage, Color.Orange);
                        continue;
                    }

                    double idsValue = double.Parse(othTbRef.Rows[tbRowIdx].RowData[othTbRef.BinXIdsMaxIdx != -1 ? othTbRef.BinXIdsMaxIdx : othTbRef.IdsMaxIdx]);

                    //STEP1b. 將找到的Performance type在allIdsVal裡搜尋, 若有且比記錄的小, 則以小的值覆蓋
                    bool isFoundSamePMode = false;
                    foreach (IdsMax onePower in oneIdsSpec)
                    {
                        if (tbDomainType == onePower.Name) //MC/MS/MG
                        {
                            isFoundSamePMode = true;
                            if (onePower.IdsMaxValue > idsValue) //取其小
                            {
                                onePower.IdsMaxValue = idsValue;
                            }
                        }
                    }

                    if (!isFoundSamePMode) //該Power第一次被加入, 直接加入List
                    {
                        var tmp = new IdsMax { Name = tbDomainType, IdsMaxValue = idsValue };
                        oneIdsSpec.Add(tmp);
                    }
                }

                allIdsSpec.Add(oneIdsSpec);
            }
            return allIdsSpec;
        }

        private static List<List<IdsMax>> GetIdsMaxFromTableCs()
        {
            var allIdsMax = new List<List<IdsMax>>();
            for (int i = 0; i < BinCutData.BinningTables.Count; i++)
            {
                var idsMaxs = new List<IdsMax>();
                BinningTable binningTable = BinCutData.BinningTables[i];
                IEnumerable<IGrouping<string, BinningRow>> groups = binningTable.Rows.GroupBy(x => x.RowData[binningTable.ModeIdx][..4]);
                foreach (IGrouping<string, BinningRow> group in groups)
                {
                    BinningRow firstRow = group.First();
                    // ex: PCPU
                    string domain = firstRow.RowData[binningTable.DomainIdx];
                    double idsMaxValue = group.Min(x => x.GetIdsMax(binningTable));
                    var idsMax = new IdsMax { Name = group.Key, IdsMaxValue = idsMaxValue, Domain = domain };
                    idsMaxs.Add(idsMax);
                }

                BinningTable othTbRef = BinCutData.OtherRailTables[i];
                IEnumerable<IGrouping<string, BinningRow>> groupOthres = othTbRef.Rows.GroupBy(x => x.GetOtherRailDomainType(othTbRef, BinCutConfig.PowerType, BinCutConfig.DomainInOtherRail2Power));
                foreach (IGrouping<string, BinningRow> group in groupOthres)
                {
                    BinningRow firstRow = group.First();
                    // ex: PCPU
                    string domain = firstRow.RowData[binningTable.DomainIdx];
                    double idsMaxValue = group.Min(x => x.GetIdsMax(binningTable));
                    var idsMax = new IdsMax { Name = group.Key, IdsMaxValue = idsMaxValue, Domain = domain };
                    idsMaxs.Add(idsMax);
                }

                allIdsMax.Add(idsMaxs);
            }
            return allIdsMax;
        }
    }
}
