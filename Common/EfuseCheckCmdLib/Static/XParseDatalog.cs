using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using Automation.Static;

using CommonLib.Extension;

using EfuseCheckCmdLib.EFuse;
using EfuseCheckCmdLib.EFuse.EFuseApp;
using EfuseCheckCmdLib.IgxlLogLib;
using EfuseCheckCmdLib.IgxlLogLib.Base;

using LogLib.Utility;

using OfficeOpenXml;

using TestPlanLib.Efuse;

namespace EfuseCheckCmdLib.AlgorithmCheck
{
    /* [Function List]
     * xBinCutCp1Log():   建構式, 主要功能在載入相關表單
     * getOneTouchDown(): 由Log中取得一次Touch Down的資訊
     * getDiceInfo():     配合getOneTouchDown(), 由一次Touch Down資訊再擷取出一個Site的DataLog, 並切出PrintOutVddBinning, AdjustVdd等之資料 
     *                    -> 由CP2程式開始的想法, 實際上除了getLogBasicInfo()應無其它function引用, 因為CP1 Script檢查時需同時看所有Site之資訊
     * getLogBasicInfo(): 引用getOneTouchDown()及getDiceInfo(), iteration搜尋出一顆pass dice(pass dice才會有全測項)以找出all instance name
     * getIds():          取得所有Site之ids電流值 -> 並將cpm在<Judge_stored_IDS>前的所有列刪除, 以減少記憶體使用量
     * createPowerTable():1. use ids current to calculate ids zone 
     *                    2. calc possible step(equation) and fill
     *                       -> C / M / LVCC / Bin / IDSMax / CPVmax / CPVmin /CPGB Array
     *                      [NOTICE] Caller of this function guarantee the order of allBinTbs list to be from bin1->bin2->...
     */

    public partial class XParseDatalog(EfuseScriptConfig efuseScriptConfig, string inDir, List<string> stdfDir, string outDir)
    {

        public class StageDicesInfo
        {
            public string Stage = "";
            public string ChannelMapStage = "";
            public List<XDiceInfo> AlldicesDiceInfos = [];
            public string Scenario = "";
            public StageDicesInfo(string stage, string channelMapStage, string scenario, List<XDiceInfo> xDiceInfos)
            {
                Stage = stage;
                ChannelMapStage = channelMapStage;
                Scenario = scenario;
                foreach (XDiceInfo info in xDiceInfos)
                {
                    AlldicesDiceInfos.Add(info);
                }
            }
        }

        public class EFuseSyntaxChkItem
        {
            public string Id = "";
            public string HighLimit = "";
            public string LowLimit = "";
        }
        private static readonly Regex _regex8 = Regex8();
        private static readonly Regex _regex7 = Regex7();
        private static readonly Regex _regex6 = Regex6();

        private static readonly Regex _regex5 = Regex5();

        //[CFG] - [Set eFuse by Site] Set Site#0 [te_misc_harv_ed_cp1] = [0x0] at row#761 in sheet EFUSE_BitDef_Table
        private static readonly Regex _regex4 = Regex4();

        // 18110101 0     VDD_SOC_MS001 VDD Define                          -1       606.2500 mV    612.5000 mV          756.2500 mV
        //	Site(0)  CFGFuse IDS_SetWriteDecimal_SetPatTestPass_Flag                                  IDS_VDD_PCPU = 797        (159.262598 mA / 0.200000mA)
        //Site(1) [CFG] - [Set eFuse]  gpio_ioh_4 = 60 at row#623 in sheet EFUSE_BitDef_Table
        private static readonly Regex _regex3 = Regex3();
        private static readonly Regex _regex2 = Regex2();
        private static readonly Regex _regex1 = Regex1();
        private static readonly Regex _regex = Regex0();

        [GeneratedRegex(@"\(\s*([^)]*_NV)\s*\)", RegexOptions.IgnoreCase)]
        private static partial Regex Regex8();

        [GeneratedRegex(@"(?<lot>\w+)_W(?<wafer>\w+)_X(?<XCorr>\w+)_Y(?<YCorr>\w+)_S(?<site>\w+)", RegexOptions.IgnoreCase)]
        private static partial Regex Regex7();

        [GeneratedRegex(@"\((?<site>\d+)\).*Read from DSSC :[\s_]*(?<Name>[\w\-\s]+)\[(?<order>.*)\]\s*=\s*(?<value1>.+)\s*\[(?<value2>\d+)\]", RegexOptions.IgnoreCase)]
        private static partial Regex Regex6();

        [GeneratedRegex(@"Site\s*\((?<site>\d+)\)\s*(?<Bank>\w+)Fuse.*[(SetWriteVariable_SiteAware)(SetWriteDecimal)(IDS_SetWriteDecimal_SetPatTestPass_Flag)]\s+(?<key>\w+)\s*=(?<value>.*)", RegexOptions.IgnoreCase)]
        private static partial Regex Regex5();

        [GeneratedRegex(@"\[(?<Bank>\w+)\].*\[Set eFuse by Site\]\s*Set\s*Site\#(?<site>\w+)\s\[(?<key>\w+)\].*\[(?<value>\w+)]\s*at", RegexOptions.IgnoreCase)]
        private static partial Regex Regex4();

        [GeneratedRegex(@"Site\s*\((?<site>\d+)\)\s*\[(?<Bank>\w+)\].*\[Set eFuse\]\s+(?<key>\w+)\s*=(?<value>.*)", RegexOptions.IgnoreCase)]
        private static partial Regex Regex3();

        [GeneratedRegex(@"(?<site>\d+)\s*(?<key>\w+)\s*VDD Define\s*", RegexOptions.IgnoreCase)]
        private static partial Regex Regex2();

        [GeneratedRegex(@"^Site\s*\(\d+\)\s*, (?<UDRCMP>\w+)[\s,]*pat:", RegexOptions.IgnoreCase)]
        private static partial Regex Regex1();

        [GeneratedRegex(@"^Site\((?<site>\d+)\), IDS", RegexOptions.IgnoreCase)]
        private static partial Regex Regex0();

        [GeneratedRegex("Setting value '.*' for fuse '.*'", RegexOptions.IgnoreCase)]
        private static partial Regex SettingValueForFuseRegex();

        [GeneratedRegex("Site.*Device_Code", RegexOptions.IgnoreCase)]
        private static partial Regex SiteDeviceCodeRegex();

        [GeneratedRegex(@"N\\*A", RegexOptions.IgnoreCase)]
        private static partial Regex NotApplicableRegex();

        [GeneratedRegex(@"Print\s*Out\s*Efuse\s*Bits\s*Content\((?<instance>\w+)\)\s*\(Site(?<sitenum>\d+)\)", RegexOptions.IgnoreCase)]
        private static partial Regex PrintOutEfuseBitsContentRegex();

        [GeneratedRegex(@"Row = (?<rowName>.*)\[(?<order>.*)\]\:\s*(?<rawData>.*)", RegexOptions.IgnoreCase)]
        private static partial Regex RowLineRegex();

        [GeneratedRegex(@"registry.\(\s*(?<type>\w+)\s*\)")]
        private static partial Regex RegistryTypeRegex();

        [GeneratedRegex("eFuseLotNumber", RegexOptions.IgnoreCase)]
        private static partial Regex EFuseLotNumberRegex();

        [GeneratedRegex("eFuseWaferID", RegexOptions.IgnoreCase)]
        private static partial Regex EFuseWaferIdRegex();

        [GeneratedRegex("eFuseDieX", RegexOptions.IgnoreCase)]
        private static partial Regex EFuseDieXRegex();

        [GeneratedRegex("eFuseDieY", RegexOptions.IgnoreCase)]
        private static partial Regex EFuseDieYRegex();

        [GeneratedRegex("Hram_ECID_53bit", RegexOptions.IgnoreCase)]
        private static partial Regex HramEcid53BitRegex();

        [GeneratedRegex("SVM_CFuse_288Bits", RegexOptions.IgnoreCase)]
        private static partial Regex SvmCFuse288BitsRegex();

        [GeneratedRegex("CFG", RegexOptions.IgnoreCase)]
        private static partial Regex CfgRegex();

        [GeneratedRegex("UDR", RegexOptions.IgnoreCase)]
        private static partial Regex UdrRegex();

        [GeneratedRegex(@"\d+", RegexOptions.IgnoreCase)]
        private static partial Regex DigitsRegex();

        [GeneratedRegex(@"(?<num>\d+)", RegexOptions.IgnoreCase)]
        private static partial Regex NumGroupRegex();

        [GeneratedRegex("_V(.*?)_")]
        private static partial Regex VersionSuffixRegex();

        [GeneratedRegex(@"(?<value>\d+(\.\d)*)(?<unit>[a-zA-Z]*)$", RegexOptions.IgnoreCase)]
        private static partial Regex ValueUnitRegex();
        //CFG table scenario e.g.A00, A01....

        /*
         * auto_eFuse_Initialize:: EnableWord CFG_SVM         = true
         * auto_eFuse_Initialize:: EnableWord CFG_SVM_A00_CP1 = true
         * auto_eFuse_Initialize:: EnableWord eFuse_Disable_ChkLMT = False
         */
        public static string ScenarioInDatalog { get; private set; } = "";
        // Use SVM CFG table
        public static readonly bool IsUseSvm;
        // This flag use to Check whether efuse use check LMT mechanism
        public static readonly bool IsDisableChkLmt;
        private readonly EfuseScriptConfig _config = efuseScriptConfig;
        //this is for T-AutoGen Constructor used
        //this is for T-AutoGen Constructor used 
        public string InDir = inDir;
        public string OutDir = string.IsNullOrEmpty(outDir) ? Path.Combine(inDir, "output") + Path.DirectorySeparatorChar : outDir;
        //public string inPath;
        public string OutPath = "";
        //public eJobLevel TEST_CAT = eJobLevel.CP1;  //cp/cp2/ft1/ft2/qa
        public static int TestCat { get; set; }
        //set to true will let getOneTouchDown() split datalog by touchDown
        public bool IsDebugPrint = false;

        //sheet reference
        public static LoaderEfuseBitDef? BitDefRef { get; set; }
        public static EFuseStdfReader1? StdfReader { get; set; }
        //log base info
        public static string CurrentJob { get; private set; } = "";
        public static string CurrentChannelMapJob { get; private set; } = "";
        public int DiceCnt;
        public int ActiveSite;
        public int CurLnNo;
        //get from getLogBasicInfo()->getOneTouchDown()->getDiceInfo(), which one execute once per log
        public List<string> Title = [];
        public List<string> Stdfs = stdfDir;
        //all dices info
        public List<XDiceInfo> AllDiceInfos = [];
        public List<StageDicesInfo> AllDicesByStages = [];

        public ExcelPackage Package = null!;
        public List<EFuseSyntaxChkItem> SyntaxChkItems = [];
        public static readonly List<XLine> SetWriteVariableLines = [];

        public bool Bin1Only = false;

        public bool GetReadFromDssc(Dictionary<int, XDiceInfo> allDices, ref List<XLine> cpm, ref List<XLine> dsscLines)
        {
            bool isFoundDsscLines = false;
            bool isFoundSyntaxChk = false;
            bool isReadWaferData = false;
            bool isIdsDataline = false;
            var format = new TestDataFormat();
            SyntaxChkItems.Clear();
            foreach (XLine cpmLine in cpm)
            {
                #region get instance
                if (cpmLine.Line.Contains('<') && cpmLine.Line.Contains('>'))
                {
                    cpmLine.Type = LineType.Instance;
                    dsscLines.Add(cpmLine);
                    if (cpmLine.Line.Contains("ReadWaferData>"))
                    {
                        isReadWaferData = true;
                        continue;
                    }
                    if (cpmLine.Line.Contains("IDS_"))
                    {
                        isIdsDataline = true;
                        continue;
                    }
                    else
                    {
                        isIdsDataline = false;
                    }
                    continue;
                }
                else if (cpmLine.Line.ContainsIgnoreCase("Print Out EFuse ReadOut Bits"))
                {
                    cpmLine.Type = LineType.Instance;
                    dsscLines.Add(cpmLine);
                    isIdsDataline = false;
                    continue;
                }
                #endregion

                if (ReadFromDssc(dsscLines, cpmLine))
                {
                    continue;
                }
                if (format.IsFormatLine(cpmLine.Line) || string.IsNullOrEmpty(cpmLine.Line))
                {
                    continue;
                }

                #region Fuse Pattern Check
                bool isFusePatternLine = cpmLine.Line.Contains("_EF_");
                #endregion

                bool flowControl = HandleDataRow(allDices, isIdsDataline, format, cpmLine, isFusePatternLine);
                if (!flowControl)
                {
                    continue;
                }

                #region UDR check 
                if (cpmLine.Line.Contains("pat:"))
                {
                    cpmLine.Type = LineType.UdrcmpPat;
                    dsscLines.Add(cpmLine);
                }
                #endregion

                #region setWrite from HIP
                if (SetWrite(dsscLines, cpmLine))
                {
                    continue;
                }
                #endregion 

                #region SetWrite from Bincut
                if (cpmLine.Line.Contains("VDD Define"))
                {
                    cpmLine.Type = LineType.BvReal;
                    dsscLines.Add(cpmLine);
                }
                #endregion

                CheckSumPrrInformation(dsscLines, cpmLine);

                HandleIedaData(dsscLines, cpmLine);

                HandleEccData(dsscLines, cpmLine);

                if (cpmLine.Line.Contains("SetPatTestPass_Flag"))
                {
                    continue;
                }

                if (cpmLine.Line.Contains("ECID Hexadecimal code =") || cpmLine.Line.Contains("DEVICE_CODE") || cpmLine.Line.Contains("Prober Hex code"))
                {
                    cpmLine.Type = LineType.Prr;
                    dsscLines.Add(cpmLine);
                }

                isReadWaferData = HandleIsReadWaferData(dsscLines, isReadWaferData, format, cpmLine);

                if (cpmLine.Line.Contains("Syntax_Chk>"))
                {
                    isFoundSyntaxChk = true;
                    continue;
                }

                isFoundSyntaxChk = HandleFoundSyntaxChk(isFoundSyntaxChk, cpmLine);
            }

            if (dsscLines.Count != 0)
            {
                isFoundDsscLines = true;
            }

            return isFoundDsscLines;
        }

        private static bool HandleDataRow(Dictionary<int, XDiceInfo> allDices, bool isIdsDataline, TestDataFormat testDataFormat, XLine xLine, bool isFusePatternLine)
        {
            DataFormatDataRow? dataline = testDataFormat.GetDataRow(xLine.Line, null, false);
            if (dataline != null)
            {
                dataline.ActuallineNumber = xLine.LineNo;
                if (isIdsDataline && allDices.TryGetValue(dataline.ActiveSite, out XDiceInfo? value))
                {
                    value.IdsMeasLines.Add(dataline);
                }
                else if (isFusePatternLine && allDices.TryGetValue(dataline.ActiveSite, out XDiceInfo? value1))
                {
                    value1.FusePatternLines.Add(dataline);
                }
            }
            else
            {
                if (xLine.Line.Contains("Key_Name:="))
                {
                    var setWriteHip = new SetHipDicItem(xLine.Line) { Number = xLine.LineNo };
                    if (string.IsNullOrEmpty(setWriteHip.Site) || string.IsNullOrEmpty(setWriteHip.Key))
                    {
                        //bypass the null key
                        return false;
                    }
                    if (allDices.ContainsKey(int.Parse(setWriteHip.Site)))
                    {
                        allDices[int.Parse(setWriteHip.Site)].HipMeasLines.Add(setWriteHip);
                    }
                }
            }

            return true;
        }

        private bool HandleFoundSyntaxChk(bool isFoundSyntaxChk, XLine xLine)
        {
            try
            {
                if (isFoundSyntaxChk)
                {
                    string[] strspt = xLine.Line.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);
                    if (strspt.Length >= 9)
                    {
                        EFuseSyntaxChkItem? existItem = SyntaxChkItems.FirstOrDefault(p => p.Id.EqualsIgnoreCase(strspt[2]));
                        if (existItem == null)
                        {
                            var item = new EFuseSyntaxChkItem();
                            List<string> logs = MergeValueAndUnit(strspt);
                            item.Id = strspt[2];
                            item.LowLimit = logs[4];
                            item.HighLimit = logs[6];
                            SyntaxChkItems.Add(item);
                        }
                    }
                    else
                    {
                        isFoundSyntaxChk = false;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }

            return isFoundSyntaxChk;
        }

        private static void CheckSumPrrInformation(List<XLine> xLines, XLine xLine)
        {
            #region CheckSum PRR information
            if (xLine.Line.ContainsIgnoreCase("**EFUSE CHECK SUM FROM BANK"))
            {
                xLine.Type = LineType.CheckSum;
                xLines.Add(xLine);
            }
            #endregion
        }

        private static void HandleIedaData(List<XLine> xLines, XLine xLine)
        {
            #region IEDA Data
            if (xLine.Line.Contains("IEDA registry"))
            {
                xLine.Type = LineType.Ieda;
                xLines.Add(xLine);
            }
            if (xLine.Line.Contains("Set PRR to"))
            {
                xLine.Type = LineType.Ieda;
                xLines.Add(xLine);
            }
            #endregion
        }

        private static void HandleEccData(List<XLine> xLines, XLine xLine)
        {
            #region ECC Data
            if (xLine.Line.Contains("Print Out EFuse Bits Content") || xLine.Line.Contains("Row =") || xLine.Line.Contains("End of printing out all Program Bit"))
            {
                xLine.Type = LineType.Ecc;
                xLines.Add(xLine);
            }
            #endregion
        }

        private static bool HandleIsReadWaferData(List<XLine> xLines, bool isReadWaferData, TestDataFormat testDataFormat, XLine xLine)
        {
            if (isReadWaferData)
            {
                if (testDataFormat.GetDataRow(xLine.Line, null, false) == null)
                {
                    isReadWaferData = false;
                }
                else
                {
                    xLine.Type = LineType.WaferData;
                    xLines.Add(xLine);
                }
            }

            return isReadWaferData;
        }

        private static bool SetWrite(List<XLine> xLines, XLine xLine)
        {
            if (xLine.Line.Contains("SetWriteVariable_SiteAware") ||
                                xLine.Line.Contains("SetWriteDecimal") ||
                                xLine.Line.Contains("[Set eFuse]") ||
                                xLine.Line.Contains("[Set eFuse by Site]") ||
                                SettingValueForFuseRegex().IsMatch(xLine.Line) ||
                                xLine.Line.Contains("Set fuse value in cache") ||
                                xLine.Line.Contains("is set to value") ||
                                xLine.Line.Contains("Fuse value locked for"))
            {
                xLine.Type = LineType.HipReal;
                xLines.Add(xLine);
                SetWriteVariableLines.Add(xLine);
                return true;
            }

            return false;
        }

        private static bool ReadFromDssc(List<XLine> xLines, XLine xLine)
        {
            #region Read from DSSC
            if (xLine.Line.Contains("Double-Bits = ")) //CONFIG,Double-Bits = True
            {
                xLine.Type = LineType.FuseBitMode;
                xLines.Add(xLine);
                return true;
            }

            if (xLine.Line.Contains("FUSE,"))
            {
                if (xLine.Line.Split(',').Length > 3)
                {
                    xLine.Type = LineType.FuseRawValue;
                    xLines.Add(xLine);
                    return true;
                }
            }

            if (xLine.Line.Contains("Read from DSSC"))
            {
                xLine.Type = LineType.ReadFromDssc;
                xLines.Add(xLine);
                return true;
            }
            #endregion
            return false;
        }

        //------------------------------------------------------------------------
        private static List<string> MergeValueAndUnit(string[] data)
        {
            var datalist = new List<string>();
            int i = 0;

            while (i < data.Length)
            {
                if (data[i].Length > 3 || double.TryParse(data[i], out _))
                {
                    datalist.Add(data[i]);
                }
                else
                {
                    datalist[^1] = data[i - 1] + data[i];
                }
                i++;
            }
            return datalist;
        }

        //------------------------------------------------------------------------
        //Get Log Basic infomation, flow as below:
        //0. Get titles
        //1. Get max site
        //2. Loop catch one site test info until find a bin 1 dice
        //3. Get all the test instance
        public void GetLogBasicInfo(Action<string, string> appendRichText, StreamReader streamReader)
        {
            //var regJob = @"_*(?<Job>(\w*(cp)|(ft))\d+)_*";
            //STEP0. Get title
            string line;
            Title.Clear();
            while (!streamReader.EndOfStream)
            {
                //char[] msg = new char[512];
                //line = sr.ReadToEnd();
                line = streamReader.ReadLine()!;
                if (line.Contains("Device#"))
                {
                    break;
                }

                if (!string.IsNullOrEmpty(line))
                {
                    Title.Add(line);
                }
            }
            streamReader.DiscardBufferedData();
            streamReader.BaseStream.Seek(0, SeekOrigin.Begin);

            TestCat = 0;
            foreach (string titleLine in Title)
            {
                if (titleLine.Contains("Channel map:"))
                {
                    string regJob = "_(?<Job>((cp)|(ft)|(wlft)))";
                    CurrentChannelMapJob = Regex.Match(titleLine, regJob, RegexOptions.IgnoreCase).Groups["Job"].Value;
                    break;
                }
                if (titleLine.Contains("Job Name:"))
                {
                    line = titleLine;
                    string[] spt = line.Split([':'], StringSplitOptions.RemoveEmptyEntries);
                    string testCatTmp = spt[1].Trim().ToUpper();
                    foreach (string job in EfuseAlgorithmCheck.JobFlow.Keys)
                    {
                        if (testCatTmp.EqualsIgnoreCase(job))
                        {
                            CurrentJob = job;
                            break;
                        }
                    }
                    if (string.IsNullOrEmpty(CurrentJob))
                    {
                        List<string> sgmts = [.. testCatTmp.Split('_')];
                        foreach (string sgmt in sgmts)
                        {
                            if (sgmt.ContainsIgnoreCase("cp") ||
                                sgmt.ContainsIgnoreCase("ft"))
                            {
                                CurrentJob = sgmt;
                                break;
                            }
                        }
                    }

                    appendRichText("Current Job: " + CurrentJob + Environment.NewLine, "Blue");
                }
                if (titleLine.Contains("Part Type:"))
                {
                    line = titleLine;
                    string[] spt = line.Split([':'], StringSplitOptions.RemoveEmptyEntries);
                    string testCatTmp = spt[1].Trim().ToUpper();
                    foreach (string job in EfuseAlgorithmCheck.JobFlow.Keys)
                    {
                        if (job.StartsWithIgnoreCase(testCatTmp))
                        {
                            CurrentJob = job;
                            break;
                        }
                    }
                }
            }

            //after read basic info, stream back to 0 pos
            streamReader.DiscardBufferedData();
            streamReader.BaseStream.Seek(0, 0);
        }

        //Get 
        //Get Log Basic infomation, flow as below:
        /*
         * auto_eFuse_Initialize:: EnableWord CFG_SVM         = true
         * auto_eFuse_Initialize:: EnableWord CFG_SVM_A00_CP1 = true
         * auto_eFuse_Initialize:: EnableWord eFuse_Disable_ChkLMT = False
         */
        public void GetEfuseInfo(StreamReader streamReader)
        {
            //var regCFGSVM = @"EnableWord\s*(CFG_SVM)+\s*=+\s*(?<flag>\w+)";
            //var regDisable_ChkLMT = @"Disable_ChkLMT\w*\s*=\s*(?<flag>\w+)";
            bool flagCfgsvm = false;
            bool flagCfgsvmScenario = false;
            bool flagDisableChkLmt = false;
            //var S5eSpecial = @"Test Program Name :: .*CFG Condition = CFG_(?<condition>)NONAND";
            //STEP0. Get title
            string line;
            Title.Clear();
            int scanindex = 0;
            while (!streamReader.EndOfStream && (!flagCfgsvm || !flagCfgsvmScenario ||
                !flagDisableChkLmt) && scanindex <= 100)
            {
                //line = sr.ReadToEnd();
                line = streamReader.ReadLine()!;

                if (line.Contains("Test Program Name ::") && line.Contains("CFG Condition ="))
                {
                    string regCfg = "CFG Condition = CFG_" + @"(?<config>\w+)";
                    ScenarioInDatalog = Regex.Match(line, regCfg, RegexOptions.IgnoreCase).Groups["config"].Value;
                }

                if (line.Contains("Config") && line.Contains("is selected!"))
                {
                    string regCfg = @"Config """ + @"(?<config>\w+)" + @""" is selected!";
                    ScenarioInDatalog = Regex.Match(line, regCfg, RegexOptions.IgnoreCase).Groups["config"].Value;
                    break;
                }

                scanindex++;
            }
            streamReader.DiscardBufferedData();
            streamReader.BaseStream.Seek(0, SeekOrigin.Begin);

        }

        public static string RemoveUnit(string inputStr)
        {
            string retVal = inputStr.Trim(' ');

            /* inputStr Type:
             * 1. 0x0000      -> do nothing
             * 2. D325(Hex)   -> remove (Hex)
             * 3. 29.6000mA   -> remove mA
             * 4. N6T303      -> do nothing
             * 5. CFG_ALL_0   -> do nothing
             */
            bool isUnit = retVal.EndsWith("ma") | retVal.EndsWith("mv") | retVal.EndsWith('a') | retVal.EndsWith('v') | retVal.EndsWith("ohm");
            if (retVal.Contains("0x"))
            {
                //retVal = (double)Convert.ToUInt32(inputStr, 16);
            }
            else if (retVal.ContainsIgnoreCase("(HEX)"))
            {
                retVal = inputStr.ToUpper().Replace("(HEX)", "");
            }
            else if (isUnit && ValueUnitRegex().IsMatch(retVal))
            {
                string unitStr = ValueUnitRegex().Match(retVal).Groups["unit"].ToString();
                string valueStr = retVal;
                if (!string.IsNullOrEmpty(unitStr))
                {
                    valueStr = valueStr.Replace(unitStr, "");
                }

                retVal = valueStr;
            }

            return retVal;
        }

        //------------------------------------------------------------------------
        public static void AccessReadFromDssc(Dictionary<int, XDiceInfo> allDices, List<XLine> xLines, EfuseScriptConfig efuseScriptConfig)
        {
            string curBlock = "";
            string pprCode = "";
            string eccInstance = "";
            int eccSite = -1;
            var format = new TestDataFormat();
            Dictionary<int, string> prrSite = [];
            List<string> blocklist = BitDefRef != null ? BitDefRef.BitDefTable.BlockList : efuseScriptConfig.DatalogBlockList;

            try
            {
                foreach (XLine prgLine in xLines)
                {
                    string line = prgLine.Line;
                    List<string> spt = [.. line.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries)];
                    if (ReadBlock(ref curBlock, line, blocklist, prgLine, spt))
                    {
                        continue;
                    }

                    switch (prgLine.Type)
                    {
                        case LineType.HipReal:
                            HandleHipReal(allDices, prgLine);
                            break;
                        case LineType.BvReal:
                            HandleBvReal(allDices, line, spt);
                            break;
                        case LineType.ReadFromDssc:
                            HandleReadFromDssc(allDices, curBlock, line);
                            break;
                        case LineType.PostCheck:
                            HandlePostCheck(allDices, line);
                            break;
                        case LineType.CheckSum:
                            HandleCheckSum(allDices, line);
                            break;
                        case LineType.Prr:
                            pprCode = HandlePrr(allDices, line, pprCode, prrSite);
                            break;
                        case LineType.Instance:
                            break;
                        case LineType.WaferData:
                            HandleWaferData(allDices, line, format);
                            break;
                        case LineType.FuseRawValue:
                            HandleFuseRawValue(allDices, line);
                            break;
                        case LineType.FuseBitMode:
                            HanldeFuseBitMode(allDices, line);
                            break;
                        case LineType.Ieda:
                            HandleIeda(allDices, prgLine);
                            break;
                        case LineType.Ecc:
                            HandleEcc(allDices, line, ref eccInstance, ref eccSite);
                            break;
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void HandleHipReal(Dictionary<int, XDiceInfo> allDices, XLine xLine)
        {
            Match? match = IsMatch(xLine.Line);
            if (match != null && int.TryParse(match.Groups["site"].Value, out int site))
            {
                if (allDices.ContainsKey(site))
                {
                    ProcessHipReal(allDices, xLine, site, match);
                }
            }
        }

        private static void HandleReadFromDssc(Dictionary<int, XDiceInfo> allDices, string curBlock, string line)
        {
            SetScenarioInDatalog(line);
            Match match = _regex6.Match(line);
            if (match.Success)
            {
                int site = ParseSite(match.Groups["site"].ToString());
                if (allDices.ContainsKey(site))
                {
                    ProcessReadFromDssc(allDices, curBlock, line, site);
                }
            }
        }

        private static void HandlePostCheck(Dictionary<int, XDiceInfo> allDices, string line)
        {
            int site = int.Parse(_regex.Match(line).Groups["site"].Value);
            if (!allDices.TryGetValue(site, out XDiceInfo? dice))
            {
                return;
            }

            dice.PostCheck = true;
        }

        private static string HandlePrr(Dictionary<int, XDiceInfo> allDices, string line, string pprCode, Dictionary<int, string> prrSite)
        {
            if (SiteDeviceCodeRegex().IsMatch(line))
            {
                return pprCode;
            }

            string newPprCode = GetPprCode(allDices, line, pprCode, prrSite);

            if (line.Contains("DEVICE_CODE"))
            {
                string[] parts = line.Split(':');
                if (parts.Length > 1)
                {
                    Match deviceMatch = _regex7.Match(parts[1].Trim());
                    if (deviceMatch.Success && int.TryParse(deviceMatch.Groups["site"].Value, out int site))
                    {
                        if (allDices.ContainsKey(site))
                        {
                            if (prrSite.TryGetValue(site, out string? siteCode))
                            {
                                allDices[site].PrrCode = siteCode;
                            }
                            else if (!string.IsNullOrEmpty(newPprCode))
                            {
                                allDices[site].PrrCode = newPprCode.Split('=')[1].Trim();
                            }
                        }
                    }
                }
            }
            return newPprCode;
        }

        private static int ParseSite(string siteStr)
        {
            int site = 0;
            try
            {
                site = int.Parse(siteStr);
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }

            return site;
        }

        private static bool ReadBlock(ref string curBlock, string line, List<string> blocklist, XLine xLine, List<string> spt)
        {
            if (xLine.Type != LineType.FuseRawValue)
            {
                if (xLine.Type == LineType.Instance || xLine.Type == LineType.UdrcmpPat)
                {
                    curBlock = GetCurBlock(curBlock, line, blocklist, xLine);
                }
                else
                {
                    if (blocklist.Contains(spt[1]))
                    {
                        curBlock = spt[1];
                        return true;
                    }
                }
            }

            return false;
        }

        private static string GetPprCode(Dictionary<int, XDiceInfo> allDices, string line, string pprCode, Dictionary<int, string> prrSite)
        {
            if (line.Contains("ECID Hexadecimal code") || line.Contains("Prober Hex code"))
            {
                pprCode = line;
                string regPrr = @"Site\((?<site>\d+)\) Prober Hex code =\s*(?<PRR>\w+)";
                if (Regex.IsMatch(line, regPrr, RegexOptions.IgnoreCase))
                {
                    int sitePrr = int.Parse(Regex.Match(line, regPrr, RegexOptions.IgnoreCase).Groups["site"].Value);
                    string codePrr = Regex.Match(line, regPrr, RegexOptions.IgnoreCase).Groups["PRR"].Value;
                    prrSite[sitePrr] = codePrr;
                    allDices[sitePrr].PrrCode = codePrr;
                }
            }

            return pprCode;
        }

        private static void ProcessReadFromDssc(Dictionary<int, XDiceInfo> allDices, string curBlock, string line, int site)
        {
            Match match = _regex6.Match(line);
            string newName = match.Groups["Name"].ToString().Trim();
            EfuseDatalogItem? datalogItem = allDices[site].AllReadFromDssc.FirstOrDefault(p => p.Block.EqualsIgnoreCase(curBlock) && p.Id.EqualsIgnoreCase(newName));
            if (datalogItem == null)
            {
                datalogItem = new EfuseDatalogItem();
                allDices[site].AllReadFromDssc.Add(datalogItem);
            }

            string decimalVal = match.Groups["value1"].ToString().Trim().Replace(" ", "");
            string rawVal = match.Groups["value2"].ToString().Trim();
            string order = match.Groups["order"].ToString().Trim();
            //add # information to isolate block name and item name

            datalogItem.Id = newName;
            datalogItem.Block = curBlock;
            datalogItem.RawData = rawVal;
            datalogItem.Value = RemoveUnit(decimalVal);
            datalogItem.Order = order;
        }

        private static void SetScenarioInDatalog(string line)
        {
            if (string.IsNullOrEmpty(ScenarioInDatalog) && line.Split(':')[1].ContainsIgnoreCase("cfg_condition"))
            {
                string regKey = ":(?<key>.*)=";
                string cfgKey =
                    Regex.Match(line, regKey, RegexOptions.IgnoreCase).Groups["key"].Value.Trim();
                string info =
                    line.Split('=')[1].Split([' '], StringSplitOptions.RemoveEmptyEntries)[0];
                if (cfgKey.EqualsIgnoreCase("cfg_condition") &&
                    !NotApplicableRegex().IsMatch(info))
                {
                    ScenarioInDatalog = info;
                }
            }
        }

        private static void HandleEcc(Dictionary<int, XDiceInfo> allDices, string line, ref string eccInstance, ref int eccSite)
        {
            try
            {
                //====== Print Out EFuse Bits Content(Config_Read_Early_LV) (Site1) **SingleBit** ============
                //Row = 000[00031:00000]: 10100101010110101100001100111100
                bool isWrite = line.ContainsIgnoreCase("writedssc") && !line.ContainsIgnoreCase("early");
                if (eccSite == -1 && string.IsNullOrEmpty(eccInstance) && line.ContainsIgnoreCase("print out efuse bits content") && isWrite && !line.ContainsIgnoreCase("3timeswritewave"))
                {
                    Match printOutLines = PrintOutEfuseBitsContentRegex().Match(line);
                    string instanceName = printOutLines.Groups["instance"].ToString();
                    int siteNum = int.Parse(printOutLines.Groups["sitenum"].ToString());
                    if (!allDices[siteNum].EccInfo.ContainsKey(instanceName))
                    {
                        allDices[siteNum].EccInfo.Add(instanceName, new EccData());
                        eccSite = siteNum;
                        eccInstance = instanceName;
                    }

                }
                else if (eccSite != -1 && !string.IsNullOrEmpty(eccInstance) && line.ContainsIgnoreCase("row = "))
                {
                    Match rowLines = RowLineRegex().Match(line);
                    string msb = rowLines.Groups["order"].ToString().Split(':')[0].Trim();
                    string lsb = rowLines.Groups["order"].ToString().Split(':')[1].Trim();
                    string rawData = rowLines.Groups["rawData"].ToString();

                    allDices[eccSite].EccInfo[eccInstance].IsMsbFirst = int.Parse(msb) >= int.Parse(lsb);
                    if (allDices[eccSite].EccInfo[eccInstance].RawDataList == null)
                    {
                        allDices[eccSite].EccInfo[eccInstance].RawDataList = [];
                    }

                    allDices[eccSite].EccInfo[eccInstance].RawDataList.Add(rawData);
                }
                else if (line.ContainsIgnoreCase("end of printing out all program bit"))
                {
                    eccSite = -1;
                    eccInstance = "";
                }

            }
            catch (Exception)
            {
            }
        }

        private static void HandleCheckSum(Dictionary<int, XDiceInfo> allDices, string line)
        {
            //**EFUSE CHECk SUM FROM BANK[hsc_arf], FIELD NAME[blob_checksum], CALCULATION RANGE[[9503:64],[47:0] ]**
            CheckSum checksum = new CheckSum();
            string regEfuseCheckSum = @"BANK\s*\[(?<bank>\s*\w+\s*)\].*FIELD\s*NAME\s*\[(?<field>\s*\w+\s*)\].*CALCULATION\s*RANGE\s*\[(?<range>.*)\]\*\*";
            checksum.Bank = Regex.Match(line, regEfuseCheckSum, RegexOptions.IgnoreCase).Groups["bank"].Value.Trim();
            checksum.FieldName = Regex.Match(line, regEfuseCheckSum, RegexOptions.IgnoreCase).Groups["field"].Value.Trim();
            checksum.Range = Regex.Match(line, regEfuseCheckSum, RegexOptions.IgnoreCase).Groups["range"].Value.Trim();
            allDices[allDices.Keys.First()].ChecksumList.Add(checksum);
        }

        private static void HandleBvReal(Dictionary<int, XDiceInfo> allDices, string line, List<string> spt)
        {
            Match match = _regex2.Match(line);
            int site = int.Parse(match.Groups["site"].Value);
            spt.RemoveAll(p => p.EqualsIgnoreCase("mV"));
            string key = spt[2];
            string value = spt[7];
            allDices[site].BvFuseInfo[key] = value;
        }

        private static void HandleIeda(Dictionary<int, XDiceInfo> allDices, XLine xLine)
        {
            try
            {
                string[] iedaLine = xLine.Line.Split('=');
                string[] iedaVal = iedaLine[^1].Trim().Split(',');
                string regType = RegistryTypeRegex().Match(xLine.Line).Groups["type"].ToString();
                Console.WriteLine(xLine.Line);
                for (int i = 0; i < iedaVal.Length; ++i)
                {
                    if (!string.IsNullOrEmpty(iedaVal[i]))
                    {
                        if (EFuseLotNumberRegex().IsMatch(regType))
                        {
                            allDices[i].EFuseLotNumber = iedaVal[i];
                        }

                        if (EFuseWaferIdRegex().IsMatch(regType))
                        {
                            allDices[i].EFuseWaferId = iedaVal[i];
                        }
                        if (EFuseDieXRegex().IsMatch(regType))
                        {
                            allDices[i].EFuseDieX = iedaVal[i];
                        }
                        if (EFuseDieYRegex().IsMatch(regType))
                        {
                            allDices[i].EFuseDieY = iedaVal[i];
                        }
                        if (HramEcid53BitRegex().IsMatch(regType))
                        {
                            allDices[i].HramEcid53Bit = iedaVal[i];
                        }

                        if (SvmCFuse288BitsRegex().IsMatch(regType))
                        {
                            allDices[i].SvmCFuse288Bits = iedaVal[i];
                        }

                        if (xLine.Line.Contains("Set PRR to"))
                        {
                            string[] arr = [.. xLine.Line.Split(' ', '.').Select(s => s.Trim())];
                            allDices[i].HramEcid53Bit = arr[6];
                        }
                    }
                }

            }
            catch (Exception)
            {
            }
        }

        private static void HanldeFuseBitMode(Dictionary<int, XDiceInfo> allDices, string line)
        {
            try
            {
                if (line.Split(',').Length > 2)
                {
                    int siteNum = int.Parse(line.Split(',')[1]);
                    if (allDices[siteNum].FuseMpDataSet == null)
                    {
                        allDices[siteNum].FuseMpDataSet = new FuseMpData(siteNum.ToString());
                    }
                    allDices[siteNum].FuseMpDataSet!.SetDataMode(line);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void HandleFuseRawValue(Dictionary<int, XDiceInfo> allDices, string line)
        {
            try
            {
                if (line.Contains("FUSE,") && line.Split(',').Length > 3)
                {
                    int siteNum = int.Parse(line.Split(',')[3]);
                    if (allDices[siteNum].FuseMpDataSet == null)
                    {
                        allDices[siteNum].FuseMpDataSet = new FuseMpData(line.Split(',')[3]);
                    }
                    allDices[siteNum].FuseMpDataSet!.SetData(line);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void HandleWaferData(Dictionary<int, XDiceInfo> allDices, string line, TestDataFormat testDataFormat)
        {
            DataFormatDataRow row = testDataFormat.GetDataRow(line, null, false)!;
            string regDec = @"(?<dec>\d+)";
            switch (row.TestName)
            {
                case "Prober_LotID":
                    allDices[row.ActiveSite].Prober.Lot = row.Pattern;
                    break;
                case "Prober_WaferID":
                    allDices[row.ActiveSite].Prober.Wafer = row.Pattern;
                    break;
                case "Prober_X":
                    allDices[row.ActiveSite].Prober.XCor =
                        Regex.Match(row.Measured, regDec, RegexOptions.IgnoreCase).Groups["dec"]
                            .Value;
                    break;
                case "Prober_Y":
                    allDices[row.ActiveSite].Prober.YCor =
                        Regex.Match(row.Measured, regDec, RegexOptions.IgnoreCase).Groups["dec"]
                            .Value;
                    break;
            }
        }

        private static void ProcessHipReal(Dictionary<int, XDiceInfo> allDices, XLine xLine, int site, Match match)
        {
            var lSetWriteItem = new SetWriteItem { Number = xLine.LineNo, Site = match.Groups["site"].Value };
            string key = match.Groups["key"].Value.ToUpper();
            string value = match.Groups["value"].Value;
            string bank = CfgRegex().Replace(match.Groups["Bank"].Value, "CONFIG");
            if (bank.ContainsIgnoreCase("UDR"))
            {
                if (!bank.ContainsIgnoreCase("UDR_"))
                {
                    bank = UdrRegex().Replace(bank, "UDR_");
                }

                if (DigitsRegex().IsMatch(bank))
                {
                    string number =
                        NumGroupRegex().Match(bank).Groups["num"]
                            .Value;
                    bank = bank.Replace(number, int.Parse(number).ToString());
                }
            }
            bank = bank.ToUpper();
            lSetWriteItem.EfuseKey = bank + "#" + key;
            if (value.Contains('@'))
            {
                string regRefKey = @"(?<name>[\w,]+)( Reverse.* Bit : (?<bit>\d+))*";
                lSetWriteItem.ReferenceValue = value.Split('@')[0].Trim();
                if (lSetWriteItem.ReferenceValue.Contains('#'))
                {
                    lSetWriteItem.ReferenceValue = lSetWriteItem.ReferenceValue.Split(' ')[0];
                }

                lSetWriteItem.ReferenceKey = Regex.Match(value.Split('@')[1], regRefKey, RegexOptions.IgnoreCase).Groups["name"].Value;
                lSetWriteItem.ReverseBit = Regex.Match(value.Split('@')[1], regRefKey, RegexOptions.IgnoreCase).Groups["bit"].Value;
            }
            else
            {
                string regIdsFmt = @"(.*?\))";
                bool ismatchIdsFmt = Regex.IsMatch(value, regIdsFmt, RegexOptions.IgnoreCase);
                lSetWriteItem.ReferenceValue = ismatchIdsFmt ? Regex.Match(value, regIdsFmt, RegexOptions.IgnoreCase).Groups[0].Value : value.TrimStart().Split(null)[0];
            }
            if (!allDices[site].HardIpFuseInfo.ContainsKey(bank + "#" + key))
            {
                allDices[site].HardIpFuseInfo.Add(bank + "#" + key, lSetWriteItem);
            }
            else
            {
                //allDices[site].HardIP_FuseInfo[bank + "#" + key] = lSetWriteItem;
                if (!string.IsNullOrEmpty(lSetWriteItem.ReferenceKey))
                {
                    allDices[site].HardIpFuseInfo[bank + "#" + key].ReferenceKey = lSetWriteItem.ReferenceKey;
                }
                if (!string.IsNullOrEmpty(lSetWriteItem.ReferenceValue))
                {
                    allDices[site].HardIpFuseInfo[bank + "#" + key].ReferenceValue = lSetWriteItem.ReferenceValue;
                }

                if (!string.IsNullOrEmpty(lSetWriteItem.ReverseBit))
                {
                    allDices[site].HardIpFuseInfo[bank + "#" + key].ReverseBit = lSetWriteItem.ReverseBit;
                }
            }
        }

        private static Match? IsMatch(string line)
        {
            Match? match = null;
            if (_regex5.IsMatch(line))
            {
                match = _regex5.Match(line);
            }
            else if (_regex3.IsMatch(line))
            {
                match = _regex3.Match(line);
            }
            else if (_regex4.IsMatch(line))
            {
                match = _regex4.Match(line);
            }

            return match;
        }

        private static string GetCurBlock(string curBlock, string line, List<string> blocklist, XLine xLine)
        {
            if (xLine.Type == LineType.UdrcmpPat)
            {
                string cmpudr = _regex1.Match(line).Groups["UDRCMP"].ToString();
                foreach (string block in blocklist)
                {
                    if (Regex.IsMatch(cmpudr, block + "|" + block.Replace("_", ""), RegexOptions.IgnoreCase))
                    {
                        curBlock = block;
                        break;
                    }
                }
            }
            else
            {
                foreach (string block in blocklist)
                {
                    if (line.ContainsIgnoreCase("Print Out EFuse ReadOut Bits"))
                    {
                        string instance = _regex8.Match(line).Groups[1].Value;
                        string[] parts = instance.Split('_');
                        if (parts.Length > 0)
                        {
                            string firstPart = parts[0].Trim();
                            if (block.ContainsIgnoreCase(firstPart))
                            {
                                curBlock = block;
                            }
                        }
                    }

                    if (Regex.IsMatch(line, block + "|" + block.Replace("_", ""), RegexOptions.IgnoreCase))
                    {
                        curBlock = block;
                        break;
                    }
                    else if (Regex.IsMatch(line, block.Replace("CONFIG", "^<CFG") + "|" + block.Replace("CONFIG", "^<TRIM_") + "|" + block.Replace("ECID", "^<UID_NonDEID"), RegexOptions.IgnoreCase))
                    //Golay case
                    {
                        curBlock = block;
                        break;
                    }
                }
            }

            return curBlock;
        }

        //------------------------------------------------------------------------

        private void ParseEachPointer(string inPath, Action<string, string> appendRichText, List<TouchDown> touchDowns)
        {
            var allDices = new List<Dictionary<int, XDiceInfo>>();
            for (int i = 0; i < touchDowns.Count; i++)
            {
                //allDices.Add(new xDiceInfo[this.maxSite]);
                //for (int iSite = 0; iSite < this.maxSite; iSite++)
                //{
                //    allDices[i][iSite] = new xDiceInfo();
                //}
                var siteDiecInfo = new Dictionary<int, XDiceInfo>();
                allDices.Add(siteDiecInfo);
                foreach (int site in touchDowns[i].CurrActiveSiteNum)
                {

                    if (!siteDiecInfo.ContainsKey(site))
                    {
                        siteDiecInfo.Add(site, new XDiceInfo());
                    }
                }
            }

            appendRichText("Parsing ...", "Blue");
            Parallel.For(0, touchDowns.Count, i => MParsingEfuseDsscParallel(inPath, touchDowns[i], allDices[i]));

            //for(int i=0; i<_tdPointers.Count();i++)
            //{
            //    this.allDiceInfos.AddRange(allDices[i].Where(a=>a.site!=-1).ToList());
            //}

            AllDiceInfos.AddRange([.. allDices.SelectMany(allDice => allDice.Values)]);

            appendRichText($"Totally {AllDiceInfos.Count} dices parse complete.", "Blue");
        }

        private void MParsingEfuseDsscParallel(string inPath, TouchDown touchDown, Dictionary<int, XDiceInfo> allDices)
        {
            List<XLine> oneTouchLines = [];
            string? line;
            int lineNum = 0;

            using (var sr = new StreamReader(inPath))
            {
                sr.BaseStream.Seek(touchDown.RegionStartPtr, SeekOrigin.Begin);
                while ((line = sr.ReadLine()) != null)
                {
                    # region must put in initial

                    lineNum++;
                    if (lineNum >= touchDown.RegionLines)
                    {
                        break;
                    }

                    #endregion
                    var xLine = new XLine { LineNo = (int)(touchDown.RegionStartPtr + lineNum - 1), Line = line };
                    oneTouchLines.Add(xLine);
                }
            }
            foreach (int site in touchDown.CurrActiveSiteNum)
            {
                allDices[site].Site = site;
                allDices[site].Sort = touchDown.GetDeivices[site].Sort;
                allDices[site].SortBin = touchDown.GetDeivices[site].Bin;
                allDices[site].XCoor = touchDown.GetDeivices[site].X;
                allDices[site].YCoor = touchDown.GetDeivices[site].Y;
            }
            var dsscLines = new List<XLine>();

            if (GetReadFromDssc(allDices, ref oneTouchLines, ref dsscLines))
            {
                AccessReadFromDssc(allDices, dsscLines, _config);
            }

            //Read
            List<EfuseReadLine> readLines = GetEfuseReadtLines(oneTouchLines);
            var readRows = readLines.Select(x => x.ConvertEfuseReadRow()).ToList();
            foreach (EfuseReadRow readRow in readRows)
            {
                int site = readRow.Site;
                string name = readRow.Name;
                if (name.EqualsIgnoreCase("lot_id"))
                {
                    allDices[site].EFuseLotNumber = readRow.Value;
                }
                else if (name.EqualsIgnoreCase("wafer_id"))
                {
                    allDices[site].EFuseWaferId = readRow.Value;
                }
                else if (name.EqualsIgnoreCase("x_coordinate"))
                {
                    allDices[site].EFuseDieX = readRow.Value;
                }
                else if (name.EqualsIgnoreCase("y_coordinate"))
                {
                    allDices[site].EFuseDieY = readRow.Value;
                }
            }

            //Write
            List<SetFuseValueLine> setFuseValueLines = GetSetFuseValueLines(oneTouchLines);
            var setFuseValueRows = setFuseValueLines.Select(x => x.ConvertSetFuseValueRow()).ToList();
            foreach (SetFuseValueRow setFuseValueRow in setFuseValueRows)
            {
                var datalogItem = new EfuseDatalogItem();
                int site = setFuseValueRow.Site;
                datalogItem.Block = "XXXXX";
                datalogItem.Id = setFuseValueRow.Name;
                datalogItem.RawData = "XXXXX";
                datalogItem.Value = setFuseValueRow.Value;
                datalogItem.Order = "XXXXX";
                allDices[site].AllReadFromDssc.Add(datalogItem);
            }

        }

        private static List<EfuseReadLine> GetEfuseReadtLines(List<XLine> xLines)
        {
            return [.. xLines.Where(x => x.Line.StartsWithIgnoreCase("Site") && x.Line.Contains("EFUSE Read")).Select(x => new EfuseReadLine() { Line = x.Line, LineNo = x.LineNo })];
        }

        private static List<SetFuseValueLine> GetSetFuseValueLines(List<XLine> xLines)
        {
            return [.. xLines.Where(x => x.Line.Contains("Set fuse value in cache")).Select(x => new SetFuseValueLine() { Line = x.Line, LineNo = x.LineNo })];
        }

        //parse() is the main function to user
        //all power sequence / LVCC / flow is judge here
        public void Parse(string inPath, Action<string, string> appendRichText)
        {
            string regJob = @"_(?<Job>((cp)|(ft))\d+)_";
            //init
            CurLnNo = 0;
            DiceCnt = 0;
            AllDiceInfos.Clear();
            CurrentJob = "";
            CurrentChannelMapJob = "";

            using (StreamReader sr = new StreamReader(inPath))
            {
                appendRichText($"File Name: {Path.GetFileName(inPath)}{Environment.NewLine}", "Blue");
                //STEP1. Get Log Basic infomation, including this.powerNames(all performance), this.instName(all test items).
                //---------------------------------
                GetLogBasicInfo(appendRichText, sr);
                if (string.IsNullOrEmpty(CurrentJob))
                {
                    foreach (string job in EfuseAlgorithmCheck.JobFlow.Keys)
                    {
                        if (Regex.IsMatch(Path.GetFileName(inPath), job, RegexOptions.IgnoreCase))
                        {
                            CurrentJob = job;
                        }
                    }
                    if (string.IsNullOrEmpty(CurrentJob))
                    {
                        CurrentJob = Regex.Match(Path.GetFileName(inPath), regJob, RegexOptions.IgnoreCase).Groups["Job"].Value;
                    }
                }
                CurLnNo = 0;
                GetEfuseInfo(sr);
            }
            var dataLogFileInfo = new DataLogFileInfo(inPath);
            List<TouchDown> tdPointers = dataLogFileInfo.MGetDevicePointer();
            appendRichText($"Total touch down = {tdPointers.Count}", "Blue");
            ParseEachPointer(inPath, appendRichText, tdPointers);
        }

        protected static string GetJobName(string dataLogFile)
        {
            string jobName = "";
            if (!File.Exists(dataLogFile))
            {
                return jobName;
            }
            using (var sr = new StreamReader(dataLogFile))
            {
                string? line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line.Contains("Job Name:"))
                    {
                        string[] arr = line.Split(':');
                        jobName = arr[1].Trim(' ');
                        return jobName;
                    }

                }
            }
            return jobName;
        }

        protected static string GetProgramVersion(string dataLogFile)
        {
            string programVersionString = "";
            if (!File.Exists(dataLogFile))
            {
                return programVersionString;
            }
            using (var sr = new StreamReader(dataLogFile))
            {
                string? line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line.Contains("Prog Name:"))
                    {
                        string[] arr = line.Split(':');
                        Match match = VersionSuffixRegex().Match(arr[1]);
                        if (match.Success)
                        {
                            programVersionString = "V" + match.Groups[1].Value;
                        }
                        return programVersionString;
                    }

                }
            }
            return programVersionString;
        }

        public void Execute(bool isMergeResult, EfuseScriptConfig efuseScriptConfig, Action<string, string> appendRichText)
        {

            List<string> bufFiles = ScanFiles(InDir, OutDir, appendRichText);
            for (int idx = 0; idx < bufFiles.Count; idx++)
            {
                string inPath = bufFiles[idx];
                //sample_log.txt
                string fromFName = Path.GetFileNameWithoutExtension(inPath);
                //CP1
                string jobForFileName = GetJobName(inPath);
                //V04A
                string programVersionString = GetProgramVersion(inPath);
                //252303
                string dateStr = DateTime.Now.ToString("yyMMdd");
                string outPutFileName = $"{fromFName}_{programVersionString}_{jobForFileName}_{dateStr}";
                //string outXlsFile = Path.ChangeExtension(outPutFileName, $".xlsx");
                OutPath = Path.Combine(OutDir, outPutFileName + ".xlsx");
                appendRichText($"Parsing datalog File: {inPath}", "Black");
                StdfReader = null;
                try
                {
                    Parse(inPath, appendRichText);
                    string? stdfFile = GetRelatedstdf(inPath, Stdfs);
                    if (!string.IsNullOrEmpty(stdfFile))
                    {
                        appendRichText("Load Stdf File...", "Blue");
                        StdfReader = new EFuseStdfReader1(stdfFile, [.. BitDefRef!.BitDefTable.HipList.Keys]);
                        StdfReader.WorkFlow();
                    }

                    if (Bin1Only)
                    {
                        AllDiceInfos = [.. AllDiceInfos.Where(a => EfuseScriptUtility.IsPassBin(a.SortBin, efuseScriptConfig))];
                    }
                    else if (EfuseStatic.ShowType == 2)
                    {
                        AllDiceInfos = [.. AllDiceInfos.Where(a => !EfuseScriptUtility.IsPassBin(a.SortBin, efuseScriptConfig))];
                    }
                    else if (EfuseStatic.ShowType == 1)
                    {
                        AllDiceInfos = [.. AllDiceInfos.Where(a => EfuseScriptUtility.IsPassBin(a.SortBin, efuseScriptConfig))];
                    }

                    var singleStage = new List<StageDicesInfo> { new(CurrentJob, CurrentChannelMapJob, ScenarioInDatalog, AllDiceInfos) };
                    AllDicesByStages.AddRange(singleStage);
                    appendRichText($"Print BDF Result: {OutPath}", "Blue");
                    new EFuseAppExcelWriterMain(_config).WriteMainFile(appendRichText, OutPath, BitDefRef!, singleStage, SyntaxChkItems);
                    new EFuseAppExcelWriter(_config).WriteMpMainFile(appendRichText, OutPath, BitDefRef!, singleStage, SyntaxChkItems);
                    appendRichText("Print BDF Result done", "Blue");

                    if (new FileInfo(OutPath).Exists && string.IsNullOrEmpty(OutPath))
                    {
                        Process.Start(OutPath);
                    }
                }
                catch (XParseDatalogException logExcption)
                {
                    string msg = logExcption.ErrMsg;
                    appendRichText(msg, "Red");

                }
            }
            if (isMergeResult)
            {
                string mergeOut = Path.Combine(OutDir, "MergeResult.xlsx");
                new EFuseAppExcelWriterMain(_config).WriteMainFile(appendRichText, mergeOut, BitDefRef!, AllDicesByStages, SyntaxChkItems);
            }
        }

        private static string? GetRelatedstdf(string datalog, List<string> stdfs)
        {
            return stdfs.FirstOrDefault(p => Path.GetFileNameWithoutExtension(p).EqualsIgnoreCase(Path.GetFileNameWithoutExtension(datalog)));
        }

        public static List<string> ScanFiles(string datalogDir, string outDir, Action<string, string> appendRichText)
        {
            if (!Directory.Exists(datalogDir))
            {
                appendRichText("Can't find datalog directory, create folder." + Environment.NewLine, "Black");
                Directory.CreateDirectory(datalogDir);
            }

            if (!Directory.Exists(outDir))
            {
                appendRichText("Can't find datalog directory, create folder." + Environment.NewLine, "Black");
                Directory.CreateDirectory(outDir);
            }

            //STEP2. Get all dataLog need to parse
            List<string> bufFiles = [.. Directory.GetFiles(datalogDir, "*.txt", SearchOption.TopDirectoryOnly)];
            if (bufFiles.Count == 0)
            {
                appendRichText("Can't find any log to parse, program terminated." + Environment.NewLine, "Black");
            }

            else
            {
                appendRichText($"Found {bufFiles.Count} log(s): {Environment.NewLine}", "Black");
            }

            return bufFiles;
        }

    }

    //Site(0)@@@ Key_Name:= adc_soc_vref_105c Value:= 58
}
