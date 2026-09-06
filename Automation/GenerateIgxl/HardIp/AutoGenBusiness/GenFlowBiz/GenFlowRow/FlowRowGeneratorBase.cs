using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Business.GenNwire.Business;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.SpecialSetting;
using Automation.InputManager.Data;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow
{
    public class EnableWdTblCondition
    {
        public string EnableWd = "";
        public string Block = "";
        public string PatternName = "";

        public EnableWdTblCondition(string enablewd, string block, string patternname)
        {
            EnableWd = enablewd;
            Block = block;
            PatternName = patternname;
        }
    }
    public class EnableWdGenerateTableReader
    {
        public string HeaderEnableWd = "Enable Word";
        public string HeaderBlock = "Block";
        public string HeaderPatternName = "Pattern Name";

        public List<EnableWdTblCondition> ReadSheet(ExcelWorksheet sheet)
        {
            var result = new List<EnableWdTblCondition>();
            int startRow = 1;
            int startCol = 1;
            bool isValid = false;
            for (int i = 1; i <= sheet.Dimension.Rows; i++)
            {
                for (int j = 1; j <= sheet.Dimension.Columns; j++)
                {
                    if (sheet.Cells[i, j].Text == HeaderEnableWd &&
                        sheet.Cells[i, j + 1].Text == HeaderBlock &&
                        sheet.Cells[i, j + 2].Text == HeaderPatternName)
                    {
                        startRow = i;
                        startCol = j;
                        isValid = true;
                        break;
                    }
                }
                if (isValid)
                {
                    break;
                }
            }

            if (isValid)
            {
                for (int i = startRow + 1; i <= sheet.Dimension.Rows; i++)
                {
                    for (int j = startCol; j <= sheet.Dimension.Columns; j++)
                    {
                        if (sheet.Cells[i, j].Value != null && sheet.Cells[i, j + 1].Value != null)
                        {
                            result.Add(new EnableWdTblCondition(
                                sheet.Cells[i, j].Text,
                                sheet.Cells[i, j + 1].Text,
                                sheet.Cells[i, j + 2].Text.TrimStart('*').TrimEnd('*').Replace("*", "\\w+")));
                            break;
                        }
                    }
                }
            }

            return result;
        }
    }

    public abstract class FlowRowGeneratorBase
    {
        public HardIpInputData HardIpInputData { get; }
        protected int ForNumber;
        protected HardIpPattern Pattern;

        #region Properties

        public string LabelVoltage = string.Empty;
        public HashSet<string> Init3XFlags = new HashSet<string>();
        public HashSet<string> Init3XEnableWds = new HashSet<string>();
        public Dictionary<string, string> InitOriginalItems = new Dictionary<string, string>();

        public HardIpPattern Pat
        {
            set
            {
                Pattern = value;
                SetBasicInfoByPattern(value);
            }
            get
            {
                return Pattern;
            }
        }

        protected string SheetName = string.Empty;

        protected internal string BlockName { set; get; } = string.Empty;

        protected string SubBlockName
        {
            get
            {
                return CommonGenerator.GetSubBlockName(Pat.Pattern.GetLastPayload(), Pat.MiscInfo, BlockName, Pat.ForceCondition.IsCz2InstName);
            }
        }

        protected string TimingAc
        {
            get
            {
                return CommonGenerator.GetTimingAc(Pat);
            }

        }

        protected string InstNameSubStr
        {
            get
            {
                return HardIpService.GetInstNameSubStr(Pat.MiscInfo);
            }
        }

        public List<Timing> TimingsNwire
        {
            get
            {
                return CommonGenerator.GetTimingsNwire(Pat);
            }
        }

        protected string EnableWord
        {
            get
            {
                return CommonGenerator.GenEnableWord(Pat.Pattern.GetLastPayload(), Pat.MiscInfo, LabelVoltage, HardIpInputData);
            }
        }

        protected bool NoPattern
        {
            get
            {
                return HardIpService.IsNoPattern(Pat.Pattern.GetLastPayload());
            }
        }

        protected bool HasSweepCode
        {
            get
            {
                return Pat.SweepCodes.Count > 0 ||
                    Pat.BurstPatterns.Any(patbur => patbur.SweepCodes.Count > 0);
            }
        }

        protected bool HasSweepVoltage
        {
            get
            {
                return Pat.SweepVoltage.Count > 0;
            }
        }

        protected bool IsNestSweepVoltage
        {
            get
            {
                return HardIpService.IsNestSweepVoltage(Pat);
            }
        }

        public bool NeedRealtimeBinOut
        {
            get
            {
                return LocalSpecs.Options.BinTableBeforeEachItem ||
                    Pat.MiscInfoDict.ContainsKey(HardIpConstData.RealtimePatBinOut);
            }
        }

        public bool HasShmoo
        {
            get
            {
                return Pat.Shmoo.CharSteps.Count > 0;
            }
        }

        protected string ActualLabelVoltage
        {
            get
            {
                return HardIpService.ActualLabelVoltage(LabelVoltage, Pat);
            }
        }

        public bool NoNeedToGen
        {
            get
            {
                if (!string.IsNullOrEmpty(Pat.BlockType) && LabelVoltage != "NV")
                {
                    return true;
                }

                return HardIpService.PatFlowNoNeedToGen(LabelVoltage, Pat.MiscInfo);
            }
        }

        #endregion

        private readonly List<EnableWdTblCondition> _enableWdsTbl = new List<EnableWdTblCondition>();

        protected FlowRowGeneratorBase(HardIpInputData hardIpInputData, string sheetName)
        {
            HardIpInputData = hardIpInputData;
            SheetName = sheetName;
            if (SettingStatic.BasicConfigWorkbook != null &&
                SettingStatic.BasicConfigWorkbook.Worksheets[NeededSheets.EnableWdGenerateTable] != null)
            {
                _enableWdsTbl = new EnableWdGenerateTableReader().ReadSheet(
                    SettingStatic.BasicConfigWorkbook.Worksheets[NeededSheets.EnableWdGenerateTable]);
            }
        }

        public int FlowControlFlag
        {
            get
            {
                return Pat.FlowControlFlag;
            }
        }

        public bool IsFlowInsRepeat
        {
            get
            {
                return Pat.IsFlowInsRepeat;
            }
        }

        public abstract List<FlowRow> GenTestRows(bool isCz2Only = false);

        public abstract List<FlowRow> GenRunScenarioRows(string rtosBoostinst = "", bool isFirstScenario = false);

        public abstract FlowRow GenBinTableRow(string voltage = "");

        public abstract List<FlowRow> GenShmooRows(string labelVoltage = "");

        public abstract string CreateIfConditionByTestPanDefine();

        public abstract string CreateTestFailActionByTestPlanDefine();

        public abstract List<FlowRow> GetIfFlowRows(string dataRow, FlowRow testFlowRow, List<FlowRow> flowRows);

        protected abstract void SetBasicInfoByPattern(HardIpPattern pattern);

        #region Specical setting flow row

        public List<FlowRow> GenOpcodeRowsBefPat(List<string> opcodeList)
        {
            var flowRows = new List<FlowRow>();
            flowRows.AddRange(OpCodeSettingMain.GenOpcodeSetting(opcodeList, BlockName, LabelVoltage, EnableWord));
            return flowRows;
        }

        public List<FlowRow> GenOpcodeRowsAftPat()
        {
            var flowRows = new List<FlowRow>();
            if (HardIpConstData.RegOpcodeInMisc.IsMatch(Pat.MiscInfo))
            {
                List<string> opcodeList = SearchInfo.GetOpcode(Pat, "A"); //Extract the opcode which should be "Before" the pattern
                flowRows.AddRange(OpCodeSettingMain.GenOpcodeSetting(opcodeList, BlockName, LabelVoltage, EnableWord));
            }
            return flowRows;
        }

        public List<FlowRow> GenRelayRows(bool usednewrelaysetting = false)
        {
            var flowRows = new List<FlowRow>();
            Dictionary<string, string> relaysettingmode = usednewrelaysetting ? Pat.NewRelaySetting : Pat.RelaySetting;
            if (relaysettingmode.Count > 0)
            {
                flowRows.AddRange(RelaySettingMain.GenRelaySetting(null, relaysettingmode, GenEnable(EnableWord), usednewrelaysetting));
            }

            return flowRows;
        }

        public List<FlowRow> GenNwireChangeRows()
        {
            var flowRows = new List<FlowRow>();
            if (TimingsNwire.Count > 0)
            {
                flowRows.AddRange(GenNwireSetting(null, TimingsNwire, EnableWord, CreateTestJob()));
            }
            return flowRows;
        }

        public List<FlowRow> GenNwireDisOrEnableRows()
        {
            var flowRows = new List<FlowRow>();
            if (Pat.MiscInfo.Split(';').ToList().Exists(s => s.Trim().StartsWith(HardIpConstData.PreNwireEnaOrDis, StringComparison.OrdinalIgnoreCase)))
            {
                flowRows.AddRange(GenNwireSettingByMisc(null, Pat.MiscInfo, EnableWord));
            }
            return flowRows;
        }

        public List<FlowRow> GenExtraRowsByMisc()
        {
            var flowRows = new List<FlowRow>();
            if (Pat.MiscInfo.Split(';').ToList().Exists(s => s.Trim().StartsWith(HardIpConstData.CallExtraFlow, StringComparison.OrdinalIgnoreCase)))
            {
                foreach (string misc in Pat.MiscInfo.Split(';'))
                {
                    if (Regex.IsMatch(misc, HardIpConstData.CallExtraFlow, RegexOptions.IgnoreCase) && misc.Contains(":"))
                    {
                        var newRow = new FlowRow { Opcode = OpCode.Call, Parameter = misc.Split(':')[1] };
                        flowRows.Add(newRow);
                    }
                }
            }
            return flowRows;
        }

#nullable enable
        internal List<FlowRow>? GenSweepCodeForRow()
#nullable restore
        {
            if (!HasSweepCode || (FlowControlFlag == 1 && IsFlowInsRepeat))
            {
                return null;
            }

            var flowRows = new List<FlowRow>();
            if (Pat.Pattern.IsMultiple())
            {
                var sweepCodeUsedForMtd = new List<FlowRow>();
                HashSet<string> srcCodeIndxSet = new HashSet<string>();
                foreach (HardIpPattern pat in Pat.BurstPatterns)
                {
                    foreach (KeyValuePair<string, List<SweepCode>> sweepItems in pat.SweepCodes)
                    {
                        string srcCodeIndx = $"SrcCodeIndx{sweepItems.Key}";
                        if (srcCodeIndxSet.Contains(srcCodeIndx))
                        {
                            continue;
                        }
                        else
                        {
                            srcCodeIndxSet.Add(srcCodeIndx);
                        }

                        var flowRow = new FlowRow { Opcode = "for", Env = CreateTestEnv() };
                        flowRow.Enable = GenEnable(EnableWord, flowRow.Env);
                        int steps;
                        if (sweepItems.Value.Any(p => p.Type == SweepCode.SweepType.Custom))
                        {
                            steps = sweepItems.Value.Where(p => p.Type == SweepCode.SweepType.Custom).ToList()[0].Step;
                            flowRow.Parameter = string.Format("{0} = 0; {0} < {1}; {0}++", srcCodeIndx, steps);
                        }
                        else
                        {
                            flowRow.Parameter = string.Format("{0} = {1}; {0} <= {2}; {0}+={3}", srcCodeIndx,
                                sweepItems.Value[0].Start, sweepItems.Value[0].End, sweepItems.Value[0].Step);
                        }
                        if (Pat.Pattern.IsMultiTimeDomain() &&
                            sweepCodeUsedForMtd.Any(x =>
                                x.Opcode.Equals(flowRow.Opcode, StringComparison.OrdinalIgnoreCase) &&
                                x.Env.Equals(flowRow.Env, StringComparison.OrdinalIgnoreCase) &&
                                x.Enable.Equals(flowRow.Enable, StringComparison.OrdinalIgnoreCase) &&
                                x.Parameter.Equals(flowRow.Parameter, StringComparison.OrdinalIgnoreCase)))
                        {
                            continue;
                        }

                        flowRows.Add(flowRow);
                        sweepCodeUsedForMtd.Add(flowRow);
                        if (!HardIpStatic.FlowUsedInteger.Contains(srcCodeIndx))
                        {
                            HardIpStatic.FlowUsedInteger.Add(srcCodeIndx);
                        }

                        ForNumber++;
                    }
                }
            }
            else
            {
                foreach (KeyValuePair<string, List<SweepCode>> sweepItems in Pat.SweepCodes)
                {
                    string srcCodeIndx = "SrcCodeIndx" + sweepItems.Key;
                    var flowRow = new FlowRow { Opcode = "for", Env = CreateTestEnv() };
                    flowRow.Enable = GenEnable(EnableWord, flowRow.Env);
                    int steps;
                    if (sweepItems.Value.Any(p => p.Type == SweepCode.SweepType.Custom))
                    {
                        steps = sweepItems.Value.Where(p => p.Type == SweepCode.SweepType.Custom).ToList()[0].Step;
                        flowRow.Parameter = string.Format("{0} = 0; {0} < {1}; {0}++", srcCodeIndx, steps);
                    }
                    else
                    {
                        flowRow.Parameter = string.Format("{0} = {1}; {0} <= {2}; {0}+={3}", srcCodeIndx,
                            sweepItems.Value[0].Start, sweepItems.Value[0].End, sweepItems.Value[0].Step);
                    }

                    flowRows.Add(flowRow);
                    if (!HardIpStatic.FlowUsedInteger.Contains(srcCodeIndx))
                    {
                        HardIpStatic.FlowUsedInteger.Add(srcCodeIndx);
                    }

                    ForNumber++;
                }
            }

            return flowRows;
        }

#nullable enable
        public List<FlowRow>? GenSweepVoltageForRow()
#nullable restore
        {
            if (!HasSweepVoltage || IsNestSweepVoltage || ((FlowControlFlag == -1 || FlowControlFlag == 1) && IsFlowInsRepeat))
            {
                return null;
            }

            var flowrows = new List<FlowRow>();

            IOrderedEnumerable<KeyValuePair<string, List<SweepVData>>> svDatasXy = Pat.SweepVoltage.Where(x => !x.Key.ContainsIgnoreCase("NESTSWEEP")).OrderByDescending(p => p.Key);

            foreach (KeyValuePair<string, List<SweepVData>> svDatas in svDatasXy)
            {
                var forRow = new FlowRow { Env = CreateTestEnv() };
                forRow.Enable = GenEnable(EnableWord, forRow.Env);
                forRow.Opcode = "for";
                if (svDatas.Value[0].IsNeedConvert)
                {
                    try
                    {
                        string srcCodeIndx = $"SrcCodeIndex{svDatas.Value[0].Axis}";
                        int count = (int)
                            Math.Ceiling(
                                Math.Abs(decimal.Parse(svDatas.Value[0].Stop) - decimal.Parse(svDatas.Value[0].Start)) /
                                decimal.Parse(svDatas.Value[0].Step));
                        forRow.Parameter = string.Format("{0} = 0; {0} <= {1}; {0}++", srcCodeIndx, count);
                    }
                    catch (Exception)
                    {
                        forRow.Parameter = string.Format("{0} = {1}; {0} <= {2}; {0}+={3}",
                            svDatas.Value[0].PinName, svDatas.Value[0].Start, svDatas.Value[0].Stop, svDatas.Value[0].Step);
                    }
                }
                else
                {
                    string srcCodeIndx = $"SrcCodeIndex{svDatas.Value[0].Axis}";
                    forRow.Parameter = string.Format("{0} = {1}; {0} {4}= {2}; {0} += {3}", srcCodeIndx, svDatas.Value[0].Start, svDatas.Value[0].Stop, svDatas.Value[0].Step, svDatas.Value[0].Comparator);
                }
                flowrows.Add(forRow);
                ForNumber++;
            }

            return flowrows;
        }

#nullable enable
        public List<FlowRow>? GenSweepCodeOrVoltageNextRow()
#nullable restore
        {
            if (ForNumber > 0)
            {
                var flowRows = new List<FlowRow>();
                for (int i = 0; i < ForNumber; i++)
                {
                    var flowRow = new FlowRow { Opcode = "next" };
                    flowRows.Add(flowRow);
                }
                ForNumber = 0;
                return flowRows;
            }
            return null;
        }

        public virtual List<FlowRow> GenPreRetestRows()
        {
            return null;
        }

#nullable enable
        public FlowRow? GenRetestIfRow()
#nullable restore
        {
            if (!Pat.MiscInfoDict.ContainsKey(HardIpConstData.ReTest))
            {
                return null;
            }

            var ifRow = new FlowRow { Opcode = "if", Parameter = HardIpConstData.ReTestFlag, Enable = "" };
            return ifRow;
        }

        public FlowRow GenRetestEndIfRow()
        {
            if (!Pat.MiscInfoDict.ContainsKey(HardIpConstData.ReTest))
            {
                return null;
            }

            var ifRow = new FlowRow { Opcode = "endif", Parameter = HardIpConstData.ReTestFlag };
            return ifRow;
        }

        protected FlowRow GenRetestClearFlagRow()
        {
            var flagClearRow = new FlowRow
            {
                Opcode = OpCode.FlagClear,
                Parameter = HardIpConstData.ReTestFlag,
                Enable = EnableWord
            };
            return flagClearRow;
        }

        #endregion

        #region Flow Start row and End row

        public FlowRow GenPrintStartRow(string flowSheetName = "")
        {
            var printStartRow = new FlowRow
            {
                Opcode = OpCode.Print,
                Parameter = string.IsNullOrEmpty(flowSheetName)
                ? '"' + "Flow " + SheetName.Replace(" ", "_") + " Start " + '"'
                : '"' + flowSheetName.Replace(" ", "_") + " Start " + '"'
            };
            return printStartRow;
        }

        public List<FlowRow> GenTtrFlagClearRow(List<FlowRow> flowBodyRows)
        {
            IEnumerable<string> list = flowBodyRows.Where(x => x.Env == HardIpConstData.EnvTtr).Select(y => y.FailAction).Distinct();
            var flagClearRowList = new List<FlowRow>();
            foreach (string item in list)
            {
                if (string.IsNullOrEmpty(item))
                {
                    continue;
                }

                var flagClearRow = new FlowRow { Opcode = OpCode.FlagClear, Parameter = item };
                flagClearRowList.Add(flagClearRow);
            }
            return flagClearRowList;
        }

        public FlowRow GenPrintStopRow(string flowSheetName = "")
        {
            //stop row
            var printStopRow = new FlowRow
            {
                Opcode = OpCode.Print,
                Parameter = string.IsNullOrEmpty(flowSheetName)
                ? '"' + "Flow " + SheetName.Replace(" ", "_") + " Stop " + '"'
                : '"' + flowSheetName.Replace(" ", "_") + " Stop " + '"'
            };
            return printStopRow;
        }

        public static FlowRow GenCallRow(string callName)
        {
            var row = new FlowRow { Opcode = OpCode.Call, Parameter = callName };
            return row;
        }

        public static FlowRow GenReturnRow()
        {
            var returnRow = new FlowRow { Opcode = OpCode.Return };
            return returnRow;
        }

        #endregion

        #region Create Columns(Common)

        protected string CreateTestJob()
        {
            if (!LocalSpecs.Options.ConvertJobToEnableWord)
            {
                string jobs = SearchInfo.GetTtrEnable(Pat.TtrStr, LabelVoltage);
                if (string.IsNullOrEmpty(jobs.Trim()))
                {
                    jobs = string.Join(",", LocalSpecs.AllJobs.Select(x => $"!{x}"));
                }
                return jobs;
            }

            return "";
        }

        protected string CreateTestOpcode()
        {
            string opcode = OpCode.Test;
            Match matchOpcodeInPatt = HardIpConstData.RegOpcodeInPatt.Match(Pat.Pattern.RealPatternName);
            if (matchOpcodeInPatt.Success)
            {
                opcode = matchOpcodeInPatt.Groups["Opcode"].ToString();
                if (!OpCode.AllOpCodes.Contains(opcode))
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.I_UnsupportedOpcode_01,
                        Pat.SheetName,
                        Pat.RowNum,
                        6,
                        [opcode]
                    );
                    opcode = OpCode.Test;
                }
            }
            return opcode;
        }

        protected string CreateCharacterizeOpcode()
        {
            return OpCode.Characterize;
        }

        protected string CreateUseLimitOpcode()
        {
            return OpCode.UseLimit;
        }

        protected string CreateIfOpcode()
        {
            return OpCode.If;
        }

        protected string CreateFlagClearOpcode()
        {
            return OpCode.FlagClear;
        }

        protected string CreateEnableFlowWdOpcode()
        {
            return OpCode.EnableFlowWd;
        }

        protected string CreateEndIfOpcode()
        {
            return OpCode.EndIf;
        }

        protected string CreateTestEnv()
        {
            var envList = new List<string>();
            string env = SearchInfo.GetEnvFromPattern(Pat);
            if (!string.IsNullOrEmpty(env))
            {
                envList.Add(env);
            }

            return string.Join("_", envList);
        }

        protected string CreateTestEnable(bool isCz2Only)
        {
            string value = isCz2Only ? "!ShmooOnly" : EnableWord;
            return value;
        }

        protected string CreateBinTableJob()
        {
            return "";
        }

        protected string CreateBinTableEnv()
        {
            return "";
        }

        protected string CreateBinTableOpcode()
        {
            return /*string.IsNullOrEmpty(CreateBinTableJob()) ? HardIpConstData.OpcodeNop :*/ OpCode.BinTable;
        }

        protected string CreateBinTableEnable()
        {
            return HardIpConstData.HardipBinEnable;
        }

        protected string CreateTestPart()
        {
            return string.Join(",", Pat.PartStr.Split(','));
        }

        #endregion

        public string GenEnable(string enable, string env = "")
        {
            List<string> enableWords = !string.IsNullOrWhiteSpace(enable) ? new List<string> { enable } : new List<string>();
            var miscInfoEnableWord = new List<string>();
            string extraEnableWord = "";
            string ttrenablewd = "";
            if (Regex.IsMatch(Pat.Pattern.GetLastPayload(), HardIpConstData.RegCzPattern, RegexOptions.IgnoreCase))
            {
                return string.Join("||", enableWords);
            }

            if (SearchInfo.IsHardipRtosSheet(Pat.SheetName))
            {
                ttrenablewd = SearchInfo.GetTtrEnable(Pat.TtrStr, LabelVoltage, LocalSpecs.AllJobsRtos);
            }
            else if (SearchInfo.IsHardipIdsSheet(Pat.SheetName))
            {
                ttrenablewd = SearchInfo.GetTtrEnable(Pat.TtrStr, LabelVoltage, LocalSpecs.AllJobsIds);
            }
            else
            {
                ttrenablewd = SearchInfo.GetTtrEnable(Pat.TtrStr, LabelVoltage);
            }

            if (!string.IsNullOrEmpty(ttrenablewd))
            {
                foreach (string job in ttrenablewd.Split(','))
                {
                    enableWords.Add("Prod_" + job);
                }
            }

            foreach (string item in Pat.MiscInfo.Split(';').Where(item =>
                Regex.IsMatch(item.Split(':')[0], HardIpConstData.EnableWord, RegexOptions.IgnoreCase) ||
                Regex.IsMatch(item.Split(':')[0], HardIpConstData.MappingTtrBranch, RegexOptions.IgnoreCase)
                ))
            {
                string enableKey = item.Split(':')[1];
                if (enableKey.Contains('@'))// EnableWord:CZ,DRAM@NV,HV
                {
                    List<string> voltages = enableKey.Split('@')[1].Split(',').ToList();
                    if (voltages.Exists(p => p.Equals(LabelVoltage, StringComparison.OrdinalIgnoreCase)))
                    {
                        miscInfoEnableWord.AddRange(enableKey.Split(','));
                    }
                }
                else
                {
                    miscInfoEnableWord.AddRange(enableKey.Split(','));
                }
            }

            if (_enableWdsTbl.Count > 0)
            {
                foreach (EnableWdTblCondition item in _enableWdsTbl)
                {
                    if (BlockName.Equals(item.Block, StringComparison.OrdinalIgnoreCase) &&
                        Pat.Pattern.PatternSetList.SelectMany(pats => pats).Any(pat => Regex.IsMatch(pat, item.PatternName, RegexOptions.IgnoreCase)))
                    {
                        enableWords.Add(item.EnableWd);
                    }
                }
            }

            extraEnableWord = miscInfoEnableWord.Count > 1 ?
                "(" + string.Join("||", miscInfoEnableWord) + ")" :
                miscInfoEnableWord.FirstOrDefault();
            if (!string.IsNullOrEmpty(extraEnableWord))
            {
                return "(" + string.Join("||", enableWords.Distinct()) + ")&&" + extraEnableWord;
            }
            return string.Join("||", enableWords.Distinct());
        }

        public List<FlowRow> SortFlowRows(List<FlowRow> originFlowRows)
        {
            var sortedFlowRows = new List<FlowRow>();
            var testRowsDefault = new List<FlowRow>();
            var testRowsCalcLimit = new List<FlowRow>();
            var testRowsMeasC = new List<FlowRow>();
            if (LocalSpecs.Options.Device != EnumDevice.LCD && LocalSpecs.Options.Device != EnumDevice.RF)
            {
                if (Pat.Pattern.IsMultiTimeDomain() && Pat.FunctionName.Equals("HardIPMultiTimeDomain", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (string instpat in Pat.Pattern.RealPatternName.Split('#'))
                    {
                        IEnumerable<FlowRow> currFlowRows = originFlowRows.Where(x => x.PatName == instpat);
                        sortedFlowRows.AddRange(SortPatFlowRows(currFlowRows));
                    }
                }
                else if (Pat.Pattern.IsMultiTimeDomain() && Pat.FunctionName.Equals("hardip_mtd_test", StringComparison.OrdinalIgnoreCase))
                {
                    int patIndex = 0;
                    List<List<FlowRow>> allPatFlowRows = new List<List<FlowRow>>();
                    foreach (string pats in Pat.Pattern.RealPatternName.Split('#'))
                    {
                        List<FlowRow> onePatFlowRows = new List<FlowRow>();
                        foreach (string pat in pats.Split('+'))
                        {
                            onePatFlowRows.AddRange(
                                originFlowRows.Where(flowrow =>
                                    flowrow.PatIndex == patIndex &&
                                    flowrow.PatName == pat.Trim()
                                    ));
                            patIndex++;
                        }
                        allPatFlowRows.Add(onePatFlowRows);
                    }

                    foreach (List<FlowRow> onePatRows in allPatFlowRows)
                    {
                        foreach (FlowRow row in onePatRows)
                        {
                            if (Regex.IsMatch(row.Comment, MeasType.MeasCalc, RegexOptions.IgnoreCase) ||
                                Regex.IsMatch(row.Comment, MeasType.MeasLimit, RegexOptions.IgnoreCase))
                            {
                                testRowsCalcLimit.Add(row);
                            }
                            else if (Regex.IsMatch(row.Comment, MeasType.MeasC, RegexOptions.IgnoreCase))
                            {
                                testRowsMeasC.Add(row);
                            }
                            else
                            {
                                testRowsDefault.Add(row);
                            }
                        }
                        sortedFlowRows.AddRange(testRowsDefault);
                        sortedFlowRows.AddRange(testRowsMeasC);
                        testRowsDefault.Clear();
                        testRowsMeasC.Clear();
                    }
                    sortedFlowRows.AddRange(testRowsCalcLimit);
                }
                else
                {
                    sortedFlowRows.AddRange(SortPatFlowRows(originFlowRows));
                }

            }
            else if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                foreach (FlowRow row in originFlowRows)
                {
                    if (Regex.IsMatch(row.Comment, MeasType.MeasCalc, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(row.Comment, MeasType.MeasLimit, RegexOptions.IgnoreCase))
                    {
                        testRowsCalcLimit.Add(row);
                    }
                    else if (Regex.IsMatch(row.Comment, MeasType.MeasC, RegexOptions.IgnoreCase))
                    {
                        testRowsMeasC.Add(row);
                    }
                    else
                    {
                        testRowsDefault.Add(row);
                    }
                }
                sortedFlowRows.AddRange(testRowsDefault);
                sortedFlowRows.AddRange(testRowsMeasC);
                sortedFlowRows.AddRange(testRowsCalcLimit);
            }
            else
            {
                sortedFlowRows.AddRange(originFlowRows);
            }

            return sortedFlowRows;
        }

        public List<FlowRow> SortPatFlowRows(IEnumerable<FlowRow> currFlowRows)
        {
            var sortedFlowRows = new List<FlowRow>();
            var testRowsDefault = new List<FlowRow>();
            var testRowsCalcLimit = new List<FlowRow>();
            var testRowsMeasC = new List<FlowRow>();

            foreach (FlowRow row in currFlowRows)
            {
                if (Regex.IsMatch(row.Comment, MeasType.MeasCalc, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(row.Comment, MeasType.MeasLimit, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(row.Comment, MeasType.WiMeas, RegexOptions.IgnoreCase))
                {
                    testRowsCalcLimit.Add(row);
                }
                else if (Regex.IsMatch(row.Comment, MeasType.MeasC, RegexOptions.IgnoreCase))
                {
                    testRowsMeasC.Add(row);
                }
                else
                {
                    testRowsDefault.Add(row);
                }
            }

            sortedFlowRows.AddRange(testRowsDefault);
            sortedFlowRows.AddRange(testRowsMeasC);
            sortedFlowRows.AddRange(testRowsCalcLimit);

            return sortedFlowRows;
        }

        #region insert Enable/Disable HardIP datalog  to each HardIP subflow add by Kimi
        public FlowRow GenHardIpDatalogFlow(string header, string tn = "")
        {
            var outputRow = new FlowRow
            {
                Opcode = OpCode.Test,
                Parameter = GenHardIpEnableDisableName(header),
                TNum = tn
            };
            return outputRow;
        }

        public string GenHardIpEnableDisableName(string header)
        {
            return header + "_HardIP_Datalog_Format";
        }

        public FlowRow GenBlockDatalogFlow(string header)
        {
            var outputRow = new FlowRow { Opcode = OpCode.Test, Parameter = header + "_Datalog_Format", TNum = "" };
            return outputRow;
        }
        #endregion

        internal List<FlowRow> GenNwireSetting(SubFlowSheet flowSheet, List<Timing> timings, string enable = "", string job = "")
        {
            var flowRows = new List<FlowRow>();
            #region Default Nwire settings
            foreach (ProtocolAwarePin defalutNwire in NwireSingleton.Instance()
                        .SettingInfo.NwirePins)
            {
                var nWireRow = new FlowRow
                {
                    Opcode = OpCode.Test,
                    Parameter = "FreeRunClk_Disable_" + defalutNwire.CreatePinNameWithDiff(),
                    Enable = enable,
                    Job = job
                };
                flowSheet?.AddRow(nWireRow);

                flowRows.Add(nWireRow);
            }
            #endregion

            #region Timings in Misc Info column
            foreach (ProtocolAwarePin defalutNwire in NwireSingleton.Instance()
                        .SettingInfo.NwirePins)
            {
                var nWireRow = new FlowRow { Opcode = OpCode.Test };
                Timing timing = timings.Find(a => a.Name.ToLower().Equals(defalutNwire.CreatePinNameWithDiff(), StringComparison.OrdinalIgnoreCase));
                if (timing == null)
                {
                    nWireRow.Parameter = NwireInstance.FreeRunClkEnable + "_" + defalutNwire.CreatePinNameWithDiff() + "_keepDefault_" + timings[0].Name + "_" + timings[0].SuffixAcSpecName;
                }
                else
                {
                    nWireRow.Parameter = NwireInstance.FreeRunClkEnable + "_" + defalutNwire.CreatePinNameWithDiff() + "_" + timing.SuffixAcSpecName;
                }
                nWireRow.Enable = enable;
                nWireRow.Job = job;
                flowSheet?.AddRow(nWireRow);

                flowRows.Add(nWireRow);
            }
            #endregion

            return flowRows;
        }

        internal List<FlowRow> GenNwireSettingByMisc(SubFlowSheet flowSheet, string miscInfo, string enable = "")
        {
            var flowRows = new List<FlowRow>();
            foreach (string misc in miscInfo.Split(';'))
            {
                if (Regex.IsMatch(misc, "FreeRunClk", RegexOptions.IgnoreCase) && misc.Contains(":"))
                {
                    var newRow = new FlowRow { Opcode = OpCode.Test, Enable = enable };
                    if (misc.ContainsIgnoreCase("disable"))
                    {
                        newRow.Parameter = HardIpConstData.FreeRunClkDisable + "_" + misc.Split(':')[1];
                    }
                    else
                    {
                        newRow.Parameter = HardIpConstData.FreeRunClkEnable + "_" + misc.Split(':')[1];
                    }

                    flowSheet?.AddRow(newRow);

                    flowRows.Add(newRow);
                }
            }
            return flowRows;
        }
    }
}
