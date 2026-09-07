using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using TestPlanLib.BinCut.BinCutConfig;

namespace BinCutScriptLib.Printer
{
    internal static class BinCutPrint
    {
        private static readonly HashSet<string> _printedIdsFailLines = [];

        public static void PrintHarvestSourceCodeError(StreamWriter streamWriter, BinCutLineBase binCutLineBase)
        {
            streamWriter.WriteLine("=> HarvestSourceCode Flag condition getting error in line : {0} : {1}", binCutLineBase.LineNo, binCutLineBase.Line);
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintHarvFlagNotFind(StreamWriter streamWriter, HarvestSourceCodetLine harvestSourceCodetLine)
        {
            if (harvestSourceCodetLine.LineNo == 0)
            {
                streamWriter.WriteLine("=> Getting error in line : {0}", harvestSourceCodetLine.Line);
            }
            else
            {
                streamWriter.WriteLine("=> Getting error in line : {0}", harvestSourceCodetLine.LineNo);
            }
            streamWriter.WriteLine("Can not find harvest flags in data log !!!");
            streamWriter.Flush();
        }

        public static void PrintFinalStepNegativeOneErrorMessage(StreamWriter streamWriter, string mode)
        {
            streamWriter.WriteLine("=> FinalStep of {0} can not be -1 !!!", mode);
            streamWriter.WriteLine("=> FinalStep set 999999");
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintInterpolationError(StreamWriter streamWriter, string msg, SiteInfo[] siteInfoArray, BinCutLineBase binCutLineBase, int site, string power, double eqNLvcc, double enValue)
        {
            streamWriter.WriteLine(msg);
            msg = $"Fail line:{binCutLineBase.LineNo}         X:{siteInfoArray[site].XCoor} Y:{siteInfoArray[site].YCoor}         Site{site}";
            streamWriter.WriteLine(msg);
            msg = $"     Mode :{power}";
            streamWriter.WriteLine(msg);
            msg = $"     Expected :{eqNLvcc}";
            streamWriter.WriteLine(msg);
            msg = $"     Datalog  :{enValue}";
            streamWriter.WriteLine(msg);
            streamWriter.WriteLine("");
            siteInfoArray[site].CheckResult.IsInterpolationPass = false;
            streamWriter.Flush();
        }

        public static void PrintSiteMismatchErrorMessage(StreamWriter streamWriter, string curInstanceName, int site)
        {
            streamWriter.WriteLine("=> Site mismatch in instance : {0}", curInstanceName);
            streamWriter.WriteLine("Site {0} should be turn down because last instance pattern fail", site);
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintSyncUpErrorMessage(string instanceName, StreamWriter streamWriter)
        {
            streamWriter.WriteLine("=> Voltage never sync up in instance : {0}", instanceName);
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintErrorMessage(StreamWriter streamWriter, string errormessage)
        {
            streamWriter.WriteLine("=> {0}", errormessage);
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintNoJobMessage(StreamWriter streamWriter, string errormessage)
        {
            streamWriter.WriteLine("=> {0}", errormessage);
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintCommomError(StreamWriter streamWriter, BinCutLineBase binCutLineBase, string errormessage)
        {
            streamWriter.WriteLine("=> {0}", errormessage);
            streamWriter.WriteLine("Line {0}: {1} ", binCutLineBase.LineNo, binCutLineBase.Line);
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintCommomErrorForIdsCsharp(StreamWriter streamWriter, BinCutLineBase binCutLineBase, string errormessage)
        {
            _ = binCutLineBase.Line.Split('\t');
            string cleanLine = Reg.RegexcleanLine.Replace(binCutLineBase.Line, "");
            if (!_printedIdsFailLines.Contains(cleanLine))
            {
                streamWriter.WriteLine("=> {0}", errormessage);
                streamWriter.WriteLine("Line {0} ", cleanLine);
                streamWriter.WriteLine("");
                streamWriter.Flush();
            }
        }

        public static void PrintHarvestSourceCodeError(StreamWriter streamWriter, BinCutLineBase binCutLineBase, string expected)
        {
            streamWriter.WriteLine("Line {0}: {1} ", binCutLineBase.LineNo, binCutLineBase.Line);
            streamWriter.WriteLine("=> Expected: {0}", expected);
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintHarvestSourceCodeErrorCs(StreamWriter streamWriter, BinCutLineBase binCutLineBase, string expected)
        {
            streamWriter.WriteLine("HarvestSourceCode Error");
            streamWriter.WriteLine("Line {0}: {1} ", binCutLineBase.LineNo, binCutLineBase.Line);
            streamWriter.WriteLine("=> Datalog:  {0}", binCutLineBase.Line);
            streamWriter.WriteLine("=> Expected: {0}", expected);
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintBinoutError(StreamWriter streamWriter, BinCutLineBase binCutLineBase)
        {
            streamWriter.WriteLine("=> The pattern fail, but final hard bin is still good bin !!!");
            streamWriter.WriteLine("Line {0}: {1} ", binCutLineBase.LineNo, binCutLineBase.Line);
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintMissingLines(StreamWriter streamWriter, List<BinCutLineBase> binCutLineBases)
        {
            foreach (BinCutLineBase line in binCutLineBases)
            {
                streamWriter.WriteLine("Miss to check line {0}: {1}", line.LineNo, line.Line);
            }

            if (binCutLineBases.Count != 0)
            {
                streamWriter.WriteLine("");
            }
        }

        public static void PrintIdsError(StreamWriter streamWriter, SiteInfo[] siteInfoArray, int lineNo, int site, string line)
        {
            string msg = $"Fail line:{lineNo}         X:{siteInfoArray[site].XCoor} Y:{siteInfoArray[site].YCoor}";
            streamWriter.WriteLine(msg);
            var printLines = new List<PrintLine>
            {
                new() { Type = "Datalog", Value = line },
                new() { Type = "Issue", Value = "Ids can not be zero !!!" }
            };

            int max = printLines.Max(x => x.Type.Length);
            foreach (PrintLine printLine in printLines)
            {
                streamWriter.WriteLine(printLine.Type.PadRight(max, ' ') + " = " + printLine.Value);
            }

            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintSafeVoltage(EnumPrintType enumPrintType, StreamWriter streamWriter, SiteInfo[] siteInfoArray, int tchCnt, string curInstName, BvLineInfo bvLineInfo, int site, string lvCmpStr, string nvCmpStr = "", string vrsCmpStr = "", List<BinCutLineBase>? binCutLineBases = null, bool pass = false)
        {
            if (!string.IsNullOrEmpty(lvCmpStr) || !string.IsNullOrEmpty(nvCmpStr))
            {
                lvCmpStr = lvCmpStr.Replace("0,", ", ").ToLower();
                nvCmpStr = nvCmpStr.Replace("0,", ", ").ToLower();
            }
            PrintBvBase(enumPrintType, streamWriter, siteInfoArray, tchCnt, curInstName, bvLineInfo, site, lvCmpStr, nvCmpStr, vrsCmpStr, binCutLineBases, pass);
        }

        private static void PrintBvBase(EnumPrintType enumPrintType, StreamWriter streamWriter, SiteInfo[] siteInfoArray, int tchCnt, string curInstName, BvLineInfo bvLineInfo, int site, string lvCmpStr, string nvCmpStr, string vrsCmpStr, List<BinCutLineBase>? binCutLineBases, bool pass)
        {
            //Datalog        = BV_VDD_CPU_MC603,0,VDD_CPU=0.634,VDD_GPU=0.675,VDD_SOC=0.600,VDD_CPU_SRAM=0.718,VDD_GPU_SRAM=0.718,VDD_FIXED=0.815,VDD_LOW=0.721
            //LV (Test Plan) = BV_VDD_CPU_MC603,0,VDD_CPU=0.681,VDD_GPU=0.675,VDD_SOC=0.000,VDD_CPU_SRAM=0.721,VDD_GPU_SRAM=0.750,VDD_FIXED=0.815,VDD_LOW=0.721
            //NV (Test Plan) = BV_VDD_CPU_MC603,0,VDD_CPU=0.797,VDD_GPU=0.800,VDD_SOC=0.800,VDD_CPU_SRAM=0.800,VDD_GPU_SRAM=0.800,VDD_FIXED=0.900,VDD_LOW=0.800
            string text = pass ? "Pass line" : "Fail line";
            string msg = $"{text}:{bvLineInfo.Line.LineNo}         X:{siteInfoArray[site].XCoor} Y:{siteInfoArray[site].YCoor}          Touch Down Index = {tchCnt}";

            streamWriter.WriteLine(msg);
            var printLines = new List<PrintLine>
            {
                new() { Type = "Instance name", Value = curInstName },
                new() { Type = "Datalog", Value = bvLineInfo.Line.Line }
            };

            if (!string.IsNullOrEmpty(lvCmpStr))
            {
                printLines.Add(new PrintLine { Type = enumPrintType + "(Test Plan)", Value = lvCmpStr });
            }

            if (!string.IsNullOrEmpty(nvCmpStr))
            {
                printLines.Add(new PrintLine { Type = "NV(Test Plan)", Value = nvCmpStr });
            }

            if (!string.IsNullOrEmpty(vrsCmpStr) && binCutLineBases != null && binCutLineBases.Count != 0)
            {
                printLines.Add(new PrintLine { Type = "VRS(Test Plan)", Value = vrsCmpStr });
            }

            int max = printLines.Max(x => x.Type.Length);
            foreach (PrintLine printLine in printLines)
            {
                streamWriter.WriteLine(printLine.Type.PadRight(max, ' ') + " = " + printLine.Value);
            }

            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintBv(EnumPrintType enumPrintType, EnumSearchType enumSearchType, StreamWriter streamWriter, SiteInfo[] siteInfoArray, int tchCnt, string curInstName, BvLineInfo bvLineInfo, int site, string lvCmpStr, int bin, int eqName, string nvCmpStr = "", string vrsCmpStr = "", List<BinCutLineBase>? binCutLineBases = null, bool pass = false)
        {
            PrintBvBase(enumPrintType, streamWriter, siteInfoArray, tchCnt, curInstName, bvLineInfo, site, lvCmpStr, nvCmpStr, vrsCmpStr, binCutLineBases, pass);

            string lineOnly = bvLineInfo.LineOnly;
            string type = bvLineInfo.Type;
            PrintDetailPowerDifference(streamWriter, enumSearchType, lineOnly, type.EqualsIgnoreCase("Safe Voltage") ? vrsCmpStr : lvCmpStr, bin, eqName, bvLineInfo);
        }

        private static void PrintDetailPowerDifference(StreamWriter streamWriter, EnumSearchType enumSearchType, string line, string cmpStr, int bin, int eqName, BvLineInfo bvLineInfo)
        {
            string[] logSpt = line.Split([','], StringSplitOptions.RemoveEmptyEntries);
            string[] cmpSpt = cmpStr.Split([','], StringSplitOptions.RemoveEmptyEntries);

            int minLength = Math.Min(logSpt.Length, cmpSpt.Length);
            string[] targetSpt = cmpSpt;

            string msg;
            int diffCnt = 1;

            for (int sptIdx = 2; sptIdx < minLength; sptIdx++)
            {
                double logvalue = 0;
                double targetvalue = 0;
                if (logSpt[sptIdx].Contains('='))
                {
                    _ = double.TryParse(logSpt[sptIdx].Split('=')[1], out logvalue);
                }

                if (targetSpt[sptIdx].Contains('='))
                {
                    _ = double.TryParse(targetSpt[sptIdx].Split('=')[1], out targetvalue);
                }

                if (Math.Abs(logvalue - targetvalue) >= 0.002)
                {
                    EnumPowerType pwrNameType = GetTypeByBvName(logSpt, sptIdx);
                    if (pwrNameType == EnumPowerType.Others)
                    {
                        msg = $"\t({diffCnt}). Other voltage Fail:";
                        streamWriter.WriteLine(msg);
                        msg = $"           Expected : {targetSpt[sptIdx]}";
                        streamWriter.WriteLine(msg);
                        msg = $"           Datalog  : {logSpt[sptIdx]}";
                        streamWriter.WriteLine(msg);
                        streamWriter.WriteLine("");
                    }
                    else
                    {
                        //	(1). Performance Fail:
                        //		   Expected EQ : BinCut2 EQ1
                        //		   TestPlan : VDD_CPU=0.528
                        //		   Datalog  : VDD_CPU=0.478
                        msg = $"\t({diffCnt}). Performance Fail:";
                        streamWriter.WriteLine(msg);
                        if (enumSearchType == EnumSearchType.GradeSearch)
                        {
                            msg = $"           Expected EQ : Bin{bin} EQ{eqName}";
                            streamWriter.WriteLine(msg);
                            if (bvLineInfo.Eqn != 0 && bvLineInfo.Passbin != 0)
                            {
                                msg = $"           Datalog EQ  : Bin{bvLineInfo.Passbin} EQ{bvLineInfo.Eqn}";
                                streamWriter.WriteLine(msg);
                            }
                        }
                        msg = $"           Expected : {targetSpt[sptIdx]}";
                        streamWriter.WriteLine(msg);
                        msg = $"           Datalog  : {logSpt[sptIdx]}";
                        streamWriter.WriteLine(msg);
                        streamWriter.WriteLine("");
                    }
                    diffCnt++;
                }
            }
        }

        public static EnumPowerType GetTypeByBvName(string[] logSpt, int sptIdx)
        {
            EnumPowerType pwrNameType = EnumPowerType.Others;
            string pin = logSpt[sptIdx].Split('=')[0].ToUpper();
            if (BinCutConfig.PowerType.TryGetValue(pin, out EnumPowerType value))
            {
                pwrNameType = value;
            }
            else
            {
                string errorMessage = string.Format("The power name : " + pin + " can't be found in the config setting.");
                if (!ErrorReportManager.GetErrorList().Select(x => x.Message).Contains(errorMessage))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_12, "Datalog", 0, 0, string.Format("The power name : " + pin + " can't be found in the config setting."), [pin]);
                    BinCutController.Controller.RichTextBoxAppend(errorMessage, Color.Red);
                }
            }
            return pwrNameType;
        }

        public static void PrintDifference(StreamWriter streamWriter, SiteInfo[] siteInfoArray, string lineNo, int site, string datalog, string expect, string item, string curInstName = "", string curPatName = "")
        {
            string type = "Fail";
            Print(streamWriter, siteInfoArray, lineNo, site, datalog, expect, item, curInstName, curPatName, type);
        }

        private static void Print(StreamWriter streamWriter, SiteInfo[] siteInfoArray, string lineNo, int site, string datalog, string expect, string item, string curInstName, string curPatName, string type)
        {
            streamWriter.WriteLine("{0} line:{1}         X:{2} Y:{3}         Site:{4}", type, lineNo, siteInfoArray[site].XCoor, siteInfoArray[site].YCoor, site);
            PrintBase(streamWriter, datalog, expect, item, curInstName, curPatName);
        }

        private static void PrintBase(StreamWriter streamWriter, string datalog, string expect, string item, string curInstName, string curPatName)
        {
            if (!string.IsNullOrEmpty(curInstName))
            {
                streamWriter.WriteLine(curInstName);
            }

            if (!string.IsNullOrEmpty(curPatName))
            {
                streamWriter.WriteLine(curPatName);
            }

            streamWriter.WriteLine("{0}:", item);
            streamWriter.WriteLine("            Expected : {0}", expect);
            streamWriter.WriteLine("            Datalog  : {0}", datalog);
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        #region Print Selsram Check Result
        public static void PrintSelsramDifferenceCs(StreamWriter streamWriter, SiteInfo[] siteInfoArray, string line, string lineNo, int site, string expect, string item, string curInstName = "", string curPatName = "")
        {
            streamWriter.WriteLine("Fail line:{0}         X:{1} Y:{2}         Site:{3}", lineNo, siteInfoArray[site].XCoor, siteInfoArray[site].YCoor, site);
            PrintBaseInfo(streamWriter, line, site, expect, item, curInstName, curPatName);
        }

        private static void PrintBaseInfo(StreamWriter streamWriter, string line, int site, string expect, string item, string curInstName, string curPatName)
        {
            if (!string.IsNullOrEmpty(curInstName))
            {
                streamWriter.WriteLine(curInstName);
            }

            if (!string.IsNullOrEmpty(curPatName))
            {
                streamWriter.WriteLine(curPatName);
            }

            streamWriter.WriteLine("{0}:", item);
            streamWriter.WriteLine("       Expected :[INFO]  [Site {0}] {1}", site, expect);
            streamWriter.WriteLine("       Datalog  :{0}", line);
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }

        public static void PrintSelsramDifferencePassCs(StreamWriter streamWriter, SiteInfo[] siteInfoArray, string line, string lineNo, int site, string expect, string item, string curInstName = "", string curPatName = "")
        {
            streamWriter.WriteLine("Pass line:{0}         X:{1} Y:{2}         Site:{3}", lineNo, siteInfoArray[site].XCoor, siteInfoArray[site].YCoor, site);
            PrintBaseInfo(streamWriter, line, site, expect, item, curInstName, curPatName);
        }

        public static void PrintSame(StreamWriter streamWriter, SiteInfo[] siteInfoArray, string lineNo, int site, string datalog, string expect, string item, string curInstName = "", string curPatName = "")
        {
            string type = "Pass";
            Print(streamWriter, siteInfoArray, lineNo, site, datalog, expect, item, curInstName, curPatName, type);
        }
        #endregion

        public static void PrintDifferencePass(StreamWriter streamWriter, SiteInfo[] siteInfoArray, string lineNo, int site, string datalog, string expect, string item, string curInstName = "", string curPatName = "")
        {
            streamWriter.WriteLine("Pass line:{0}         X:{1} Y:{2}         Site:{3}", lineNo, siteInfoArray[site].XCoor, siteInfoArray[site].YCoor, site);
            PrintBase(streamWriter, datalog, expect, item, curInstName, curPatName);
        }

        public static void PrintDifference(StreamWriter streamWriter, SiteInfo[] siteInfoArray, BinCutLineBase binCutLineBase, int site, string datalog, string expect, string item, string curInstName = "", string curPatName = "")
        {
            int lineNo = binCutLineBase.LineNo;
            streamWriter.WriteLine("Fail line:{0}         X:{1} Y:{2}         Site:{3}", lineNo, siteInfoArray[site].XCoor, siteInfoArray[site].YCoor, site);
            streamWriter.WriteLine(binCutLineBase.Line);
            PrintBase(streamWriter, datalog, expect, item, curInstName, curPatName);
        }

        public static void PrintDifferencePass(StreamWriter streamWriter, SiteInfo[] siteInfoArray, BinCutLineBase binCutLineBase, int site, string datalog, string expect, string item, string curInstName = "", string curPatName = "")
        {
            int lineNo = binCutLineBase.LineNo;
            streamWriter.WriteLine("Pass line:{0}         X:{1} Y:{2}         Site:{3}", lineNo, siteInfoArray[site].XCoor, siteInfoArray[site].YCoor, site);
            streamWriter.WriteLine(binCutLineBase.Line);
            PrintBase(streamWriter, datalog, expect, item, curInstName, curPatName);
        }

        public static void PrintDifferenceCsharp(StreamWriter streamWriter, SiteInfo[] siteInfoArray, BinCutLineBase binCutLineBase, int site, string datalog, string expect, string item, string curInstName = "", string curPatName = "")
        {
            streamWriter.WriteLine("Fail line:{0}         X:{1} Y:{2}         Site:{3}", binCutLineBase.LineNo, siteInfoArray[site].XCoor, siteInfoArray[site].YCoor, site);
            PrintBase(streamWriter, datalog, expect, item, curInstName, curPatName);
        }

        public static void PrintSameIsFoundSetWriteDecimalCsharp(StreamWriter streamWriter, BinCutLineBase binCutLineBase)
        {
            streamWriter.WriteLine("Pass line:{0} , {1}", binCutLineBase.LineNo, binCutLineBase.Line);
            streamWriter.Flush();
        }

        public static void PrintAdjustProductIdentifier(StreamWriter streamWriter, ProductIdentifierLineRow productIdentifierLineRow, List<int> binXList)
        {
            streamWriter.WriteLine("<Product_Identifier>");
            streamWriter.WriteLine("Product_Identifier Fail at line : {0}", productIdentifierLineRow.Line.LineNo);
            streamWriter.WriteLine("           Expected : {0}", binXList.Contains(productIdentifierLineRow.Site) ? 1 : 0);
            streamWriter.WriteLine("           Datalog  : {0}", productIdentifierLineRow.Site);
            streamWriter.WriteLine("");
            streamWriter.Flush();
        }
    }
}
