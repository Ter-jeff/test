using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;

using CommonLib.Datalog;

using IgxlLib.Enums;

namespace BinCutScriptLib.Comparer
{
    public class LvResultsComparer(string curInstName)
    {
        private readonly string _curInstName = curInstName;

        public void CompareLvResults(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<LimitLine> limitLines, BvName bvName, EnumJob enumJob, bool isHarvAssignedTrue)
        {
            var failSites = new List<int>();

            Dictionary<string, List<LimitLine>> printOutDic = SplitPrintOutLinesByTestName(limitLines);
            List<LimitLine> groupLines;
            //STEP1. Compare IDS
            if (printOutDic.TryGetValue("IDS", out List<LimitLine>? value))
            {
                groupLines = value;
                for (int lnIdx = 0; lnIdx < groupLines.Count; lnIdx++)
                {
                    int failSite = CompareLvResultIds(streamWriter, ref siteInfoArray, groupLines[lnIdx], bvName);
                    if (failSite != -1 && !failSites.Contains(failSite))
                    {
                        failSites.Add(failSite);
                    }
                }
            }

            //STEP2. Compare LVCC
            if (printOutDic.ContainsKey(enumJob.ToString()) || printOutDic.ContainsKey("CP"))
            {
                //for old format
                groupLines = printOutDic.ContainsKey(enumJob.ToString()) ? printOutDic[enumJob.ToString()] : printOutDic["CP"];

                for (int lnIdx = 0; lnIdx < groupLines.Count; lnIdx++)
                {
                    int failSite = CompareLvResultLvcc(streamWriter, ref siteInfoArray, groupLines[lnIdx], bvName, isHarvAssignedTrue);
                    if (failSite != -1 && !failSites.Contains(failSite))
                    {
                        failSites.Add(failSite);
                    }
                }
            }

            //STEP3. Compare EQN
            if (printOutDic.TryGetValue("EQN", out List<LimitLine>? value1))
            {
                groupLines = value1;
                for (int lnIdx = 0; lnIdx < groupLines.Count; lnIdx++)
                {
                    int failSite = CompareLvResultEqn(streamWriter, ref siteInfoArray, groupLines[lnIdx], bvName, isHarvAssignedTrue);
                    if (failSite != -1 && !failSites.Contains(failSite))
                    {
                        failSites.Add(failSite);
                    }
                }
            }

            //STEP4. Compare BIN
            if (printOutDic.TryGetValue("PASSBIN", out List<LimitLine>? value2))
            {
                groupLines = value2;
                for (int lnIdx = 0; lnIdx < groupLines.Count; lnIdx++)
                {
                    int failSite = CompareLvResultBin(streamWriter, ref siteInfoArray, groupLines[lnIdx], bvName, isHarvAssignedTrue);
                    if (failSite != -1 && !failSites.Contains(failSite))
                    {
                        failSites.Add(failSite);
                    }
                }
            }

            //STEP5. Compare OFFSET
            if (printOutDic.TryGetValue("OFFSET", out List<LimitLine>? value3))
            {
                groupLines = value3;
                for (int lnIdx = 0; lnIdx < groupLines.Count; lnIdx++)
                {
                    int failSite = CompareLvResultDynamicOffset(streamWriter, ref siteInfoArray, groupLines[lnIdx], bvName);
                    if (failSite != -1 && !failSites.Contains(failSite))
                    {
                        failSites.Add(failSite);
                    }
                }
            }

            //!!turn off fail site and don't let next gradeSearch catch it
            foreach (int failSiteNo in failSites)
            {
                siteInfoArray[failSiteNo].AllPowers[bvName.Index].IsFail = true;
                siteInfoArray[failSiteNo].AllPowers[bvName.Index].IsBinOut = true;
            }
        }

        private static Dictionary<string, List<LimitLine>> SplitPrintOutLinesByTestName(List<LimitLine> limitLines)
        {
            #region Split printOutLines by testName
            var printOutDic = new Dictionary<string, List<LimitLine>>();
            foreach (LimitLine line in limitLines)
            {
                string testName = line.GetTestName();
                line.GetSiteOnly(out _);

                if (printOutDic.TryGetValue(testName, out List<LimitLine>? value))
                {
                    value.Add(line);
                }
                else
                {
                    printOutDic.Add(testName, [line]);
                }
            }
            #endregion
            return printOutDic;
        }

        public int CompareLvResultIds(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, LimitLine limitLine, BvName bvName)
        {
            // 1598226  0     GpuTd_Soc_MG007 VDD_SOC                22.t415hc 0.0000 A       24.0000 mA         60.0000 mA     0.0000         0       
            return CompareLimit(streamWriter, ref siteInfoArray, limitLine, bvName, false, CheckLvResultIds);
        }

        private void CheckLvResultIds(StreamWriter streamWriter, SiteInfo[] siteInfoArray, LimitLine limitLine, BvName bvName, bool _, int site, double log)
        {
            //Init for Binout
            if (siteInfoArray[site].AllPowers.Count > bvName.Index)
            {
                siteInfoArray[site].AllPowers[bvName.Index].IsBinOut = false;
            }

            double searchIds = SearchIds(siteInfoArray, bvName.Index, site);
            if (log != searchIds)
            {
                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, limitLine.LineNo.ToString(), site, log.ToString(CultureInfo.InvariantCulture), searchIds.ToString(CultureInfo.InvariantCulture), "IDS", _curInstName);
                siteInfoArray[site].CheckResult.IsLvResultPass = false;
            }
        }

        private static double SearchIds(SiteInfo[] siteInfoArray, int powerIdx, int site)
        {
            PowerZone objPPower = siteInfoArray[site].AllPowers[powerIdx];
            double searchIds = objPPower.IdsValue;
            if (objPPower.SearchStatus == EnumSearchStatus.Search)
            {
                searchIds = objPPower.IdsValue;
            }
            return searchIds;
        }

        public int CompareLvResultLvcc(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, LimitLine limitLine, BvName bvName, bool isHarvAssignedTrue)
        {
            // 1633685  0     SPI_Cpu_MC603_P VDD_CPU                20.x304  559.3750 mV    634.3750 mV        668.7500 mV    0.0000         0       
            return CompareLimit(streamWriter, ref siteInfoArray, limitLine, bvName, isHarvAssignedTrue, CheckLvResultLvcc);
        }

        private void CheckLvResultLvcc(StreamWriter streamWriter, SiteInfo[] siteInfoArray, LimitLine limitLine, BvName bvName, bool isHarvAssignedTrue, int site, double log)
        {
            double searchLvcc = SearchLvcc(siteInfoArray, bvName.Index, site, isHarvAssignedTrue);

            if (log > 1000.0)
            {
                log = (int)log;
                searchLvcc = (int)searchLvcc;
            }
            bool isError = !IsEqualByFormat(log, searchLvcc);
            if (isError)
            {
                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, limitLine.LineNo.ToString(), site, log.ToString(CultureInfo.InvariantCulture), searchLvcc.ToString(CultureInfo.InvariantCulture), "LVCC", _curInstName);
                siteInfoArray[site].CheckResult.IsLvResultPass = false;
            }
        }

        private static double SearchLvcc(SiteInfo[] siteInfoArray, int powerIdx, int site, bool isHarvAssignedTrue)
        {
            PowerZone objPPower = siteInfoArray[site].AllPowers[powerIdx];
            double searchLvcc = 0.0;
            if (isHarvAssignedTrue && siteInfoArray[site].IsPreVddSearch)
            {
                searchLvcc = siteInfoArray[site].AllPowers[powerIdx].PossibleSteps.Find(x => x.Bin == 1 && x.EqName == 1)!.Lvcc;
            }
            if (objPPower.SearchStatus == EnumSearchStatus.Search)
            {
                searchLvcc = objPPower.GetFinalLvcc();
            }
            return searchLvcc;
        }

        public int CompareLvResultEqn(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, LimitLine limitLine, BvName bvName, bool isHarvAssignedTrue)
        {
            // 1590318  0     EQN             VDD_SOC                22.t415hc 1.0000         2.0000             4.0000         0.0000         0       
            return CompareLimit(streamWriter, ref siteInfoArray, limitLine, bvName, isHarvAssignedTrue, CheckLvResultEqn);
        }

        private void CheckLvResultEqn(StreamWriter streamWriter, SiteInfo[] siteInfoArray, LimitLine limitLine, BvName bvName, bool isHarvAssignedTrue, int site, double log)
        {
            int searchEqn = SearchEqn(siteInfoArray, bvName.Index, site, isHarvAssignedTrue);
            ValidateAndLogDifference(streamWriter, siteInfoArray, limitLine, site, log, searchEqn, "EQN");
        }

        private void CheckLvResultBin(StreamWriter streamWriter, SiteInfo[] siteInfoArray, LimitLine limitLine, BvName bvName, bool isHarvAssignedTrue, int site, double log)
        {
            int searchBin = SearchBin(siteInfoArray, bvName.Index, site, isHarvAssignedTrue);
            ValidateAndLogDifference(streamWriter, siteInfoArray, limitLine, site, log, searchBin, "BIN");
        }

        private void ValidateAndLogDifference(StreamWriter streamWriter, SiteInfo[] siteInfoArray, LimitLine limitLine, int site, double log, int calculatedValue, string valueType)
        {
            // Check if the expected log value mismatches the calculated value
            if (log != calculatedValue)
            {
                string lineNoStr = limitLine?.LineNo.ToString() ?? "Unknown";
                string logStr = log.ToString(CultureInfo.InvariantCulture);
                string calcStr = calculatedValue.ToString(CultureInfo.InvariantCulture);

                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, lineNoStr, site, logStr, calcStr, valueType, _curInstName);

                if (siteInfoArray != null && site >= 0 && site < siteInfoArray.Length && siteInfoArray[site]?.CheckResult != null)
                {
                    siteInfoArray[site].CheckResult.IsLvResultPass = false;
                }
            }
        }

        private static int SearchEqn(SiteInfo[] siteInfoArray, int powerIdx, int site, bool isHarvAssignedTrue)
        {
            PowerZone power = siteInfoArray[site].AllPowers[powerIdx];
            int searchEqn = 0;
            if (isHarvAssignedTrue && siteInfoArray[site].IsPreVddSearch)
            {
                searchEqn = siteInfoArray[site].AllPowers[powerIdx].PossibleSteps.Find(x => x.Bin == 1 && x.EqName == 1)!.EqName;
            }
            if (power.SearchStatus == EnumSearchStatus.Search)
            {
                searchEqn = power.GetFinalEqName();
            }
            else if (power.SearchStatus == EnumSearchStatus.BinOut && BinCutConfig.IsDoAll)
            {
                return 1;
            }
            return searchEqn;
        }

        public static int CompareLimit(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, LimitLine limitLine, BvName bvName, bool isHarvAssignedTrue, Action<StreamWriter, SiteInfo[], LimitLine, BvName, bool, int, double> checkCompareAction)
        {
            limitLine.GetSiteData(out int site, out double log);
            if (siteInfoArray[site].AllPowers.Count == 0 || site == -1 || !siteInfoArray[site].IsActiveSite)
            {
                return -1;
            }

            checkCompareAction(streamWriter, siteInfoArray, limitLine, bvName, isHarvAssignedTrue, site, log);

            int failSite = -1;
            bool isFailLine = limitLine.IsFail();
            if (isFailLine)
            {
                failSite = site;
            }

            return failSite;
        }

        public int CompareLvResultBin(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, LimitLine limitLine, BvName bvName, bool isHarvAssignedTrue)
        {
            // 1590320  4     PASSBIN         VDD_SOC                22.t215hc 1.0000         1.0000             2.0000         0.0000         0  
            return CompareLimit(streamWriter, ref siteInfoArray, limitLine, bvName, isHarvAssignedTrue, CheckLvResultBin);
        }

        private static int SearchBin(SiteInfo[] siteInfoArray, int powerIdx, int site, bool isHarvAssignedTrue)
        {
            PowerZone objPPower = siteInfoArray[site].AllPowers[powerIdx];
            int searchBin = 0;
            if (isHarvAssignedTrue && siteInfoArray[site].IsPreVddSearch)
            {
                searchBin = siteInfoArray[site].AllPowers[powerIdx].PossibleSteps.Find(x => x.Bin == 1 && x.EqName == 1)!.Bin;
            }
            if (objPPower.SearchStatus == EnumSearchStatus.Search)
            {
                searchBin = objPPower.GetFinalBin();
            }
            return searchBin;
        }

        public int CompareLvResultDynamicOffset(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, LimitLine limitLine, BvName bvName)
        {
            // 1000031  4     OFFSET                         VDD_SOC                   0.x315   -100.0000 mV   0.0000 V           100.0000 mV    0.0000         0                   
            bool isFailLine = limitLine.IsFail();
            limitLine.GetSiteData(out int site, out double log);

            if (siteInfoArray[site].AllPowers.Count == 0 || site == -1 || !siteInfoArray[site].IsActiveSite)
            {
                return -1;
            }

            double dynamicOffset = siteInfoArray[site].CurrentdynamicOffset;
            if (log != dynamicOffset)
            {
                //if (BinCutConfig.ProjectConfig is Cebu && _curInstName.IndexOf("MG005", StringComparison.OrdinalIgnoreCase) >= 0 && _curInstName.IndexOf("GFXTD", StringComparison.OrdinalIgnoreCase) >= 0)
                //{
                //    sw.WriteLine("=> Cebu Special Case :  DynamicOffset is -6.25, and bypass this error !!!");
                //}
                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, limitLine.LineNo.ToString(), site, log.ToString(CultureInfo.InvariantCulture), dynamicOffset.ToString(CultureInfo.InvariantCulture), "DynamicOffset", _curInstName);
                siteInfoArray[site].CheckResult.IsLvResultPass = false;
            }

            int retSite = -1;
            if (isFailLine)
            {
                retSite = site;
            }

            return retSite;
        }

        public void ComparePlResults(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, List<LimitLine> limitLines)
        {
            var testNameList = new List<string>();
            #region Split printOutLines by testName
            var printOutDic = new Dictionary<string, List<LimitLine>>();
            foreach (LimitLine line in limitLines)
            {
                if (!(line.Line.Contains("_EQN") || line.Line.Contains("_CP") || line.Line.Contains("_PASSBIN")))
                {
                    continue;
                }

                string testName = line.GetTestName();
                line.GetSiteOnly(out _);

                if (printOutDic.TryGetValue(testName, out List<LimitLine>? value))
                {
                    value.Add(line);
                }
                else
                {
                    printOutDic.Add(testName, [line]);
                    testNameList.Add(testName);
                }
            }
            #endregion

            foreach (string testName in testNameList)
            {
                string testItem = testName.Split('_').Last();

                List<LimitLine> groupLines = printOutDic[testName];
                for (int lnIdx = 0; lnIdx < groupLines.Count; lnIdx++)
                {
                    CompareItemResult(streamWriter, ref siteInfoArray, groupLines[lnIdx], testItem);
                }
            }
        }

        public void CompareItemResult(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, LimitLine limitLine, string item)
        {
            // 19488838 0  PP_*_item( CP, PASSBIN, EQN )                           VDD_SOC            11.x410  1000.0000 mV   570.0000 mV          690.0000 mV    0.0000 mV      0
            limitLine.GetSiteData(out int site, out double log);
            string testName = limitLine.GetTestName();

            if (siteInfoArray[site].AllPowers.Count == 0 || site == -1 || !siteInfoArray[site].IsActiveSite)
            {
                return;
            }

            List<PatternInfo> search = siteInfoArray[site].PatternList.FindAll(x => (x.PatternName + "_" + item) == testName);
            int searchIdx = search.FindLastIndex(x => x.IsFail) + 1;
            double itemVal = -1.0;
            if (search.Count == 0)
            {
                streamWriter.WriteLine("Fail line:{0}         X:{1} Y:{2}         Site:{3}", limitLine.LineNo, siteInfoArray[site].XCoor, siteInfoArray[site].YCoor, site);
                streamWriter.WriteLine("Cannot find Pattern:      {0}", testName.Replace("_" + testName.Split('_').Last(), ""));
                streamWriter.WriteLine("");
                return;
            }
            switch (item)
            {
                case "CP":
                    {
                        itemVal = searchIdx >= search.Count ? 0 : search[searchIdx].Cp;
                        break;
                    }
                case "EQN":
                    {
                        itemVal = searchIdx >= search.Count ? 0 : search[searchIdx].EqName;
                        break;
                    }
                case "PASSBIN":
                    {
                        itemVal = searchIdx >= search.Count ? 0 : search[searchIdx].Bin;
                        break;
                    }
            }

            string patName = search[0].PatternName;

            if (log != itemVal)
            {
                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, limitLine.LineNo.ToString(), site, log.ToString(CultureInfo.InvariantCulture), itemVal.ToString(CultureInfo.InvariantCulture), "PL_" + item, _curInstName, patName);
                siteInfoArray[site].CheckResult.IsLvResultPass = false;
            }
        }

        public static bool IsEqualByFormat(double value1, double value2)
        {
            int index = value1.ToString().IndexOf('.');
            if (index == -1)
            {
                string string1 = value1.ToString();
                string string2 = $"{value2:F0}";
                return string1 == string2;
            }
            else
            {
                int len = value1.ToString().Length - index - 1;
                string string1 = value1.ToString();
                string string2 = string.Format("{0:F" + len + "}", value2);
                return string1 == string2;
            }
        }
    }
}
