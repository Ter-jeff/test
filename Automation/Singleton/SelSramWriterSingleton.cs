using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.PostAction.SelSram;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using TestPlanLib.Basic;
using TestPlanLib.VbtLib;

namespace Automation.Singleton
{
    public class SelSramWriterSingleton
    {
        private readonly List<string> _domainList = new List<string> { ConSoc, ConCpu, ConGfx };
        private readonly List<string> _hcDsscList = new List<string> { ConDssc, ConHardCode };
        private readonly List<string> _blockList = new List<string> { ConScan, ConMbist };

        private const string ConSoc = "Soc";
        private const string ConCpu = "Cpu";
        private const string ConGfx = "Gfx";
        private const string ConScan = "Scan";
        private const string ConMbist = "Mbist";
        private const string ConDssc = "DSSC";
        private const string ConHardCode = "HC";
        private const string ConVbt = "VBT";
        private const string ConVbtFunctionalTSrm = "Functional_T_SRM";
        private const string ConCsSelsramFuncTestMain = "SelsramFuncTestMain";
        private const string ConVbtSrmInitLoopCnt = "SRM_InitLoopCnt";
        private const string ConCsSelsramInitLoopCount = "SelsramInitLoopCount";
        private const string ConVbtSrmParseChkList = "SRM_ParseChkList";
        private const string ConCsSelsramParseTable = "SelsramParseTable";
        private const string ConVbtSrmInitDatalogSetup = "SRM_InitDatalogSetup";
        private const string ConVbtSrmGetChnType = "SRM_GetChnType";
        private const string ConLevels = "Levels";
        private const string ConDatalogSetupSram = "Datalog_setup_SRAM";
        private const string ConInitSrmParseCkList = "Initial_SRM_ParseChkList";
        private const string ConSrmGetChannelType = "SRM_GetChannelType";
        private const string ConSelsrmChkList = "SelSram_Chklist";

        public DataTable PayloadTypeTable { get; set; }
        private readonly SelSramPatternSingleton _selSramPattern;
        public List<string> UsePatternList;
        private static MultiTestSettingSheetsSingleton _multiTestSettingSheetsSingleton;

        public MultiTestSettingSheetsSingleton MultiTestSettingSheetsSingleton
        {
            set { _multiTestSettingSheetsSingleton = value; }
        }

        public Dictionary<string, PatternData> PatSetTimeSetDictionary { get; set; }

        private static SelSramWriterSingleton _instance;

        private SelSramWriterSingleton()
        {
            PatSetTimeSetDictionary = new Dictionary<string, PatternData>();
            _selSramPattern = SelSramPatternSingleton.GetInstance();
            PayloadTypeTable = new DataTable();
            UsePatternList = new List<string>();
        }

        #region Singleton
        public static SelSramWriterSingleton GetInstance()
        {
            return _instance ?? (_instance = new SelSramWriterSingleton());
        }

        public static void Initialize()
        {
            _instance = null;
        }
        #endregion

        private string GetTimeSetName(string patternName)
        {
            if (PatSetTimeSetDictionary == null)
            {
                return "";
            }

            string timeSet = "";
            if (PatSetTimeSetDictionary.TryGetValue(patternName.ToLower(), out PatternData targetPatternData))
            {
                timeSet = targetPatternData.TimeSetVersion;
            }
            return timeSet;
        }

        public void WriteInstanceSheet(ref InstanceSheet instanceSheet, ref List<string> enableGenList)
        {
            foreach (string chiplet in SelSramPatternSingleton.GetInstance().Chiplet)
            {
                string text = chiplet;
                if (!string.IsNullOrEmpty(text))
                {
                    text = "_" + text;
                }

                foreach (string domain in _domainList)
                {
                    foreach (string type in _hcDsscList)
                    {
                        foreach (string block in _blockList)
                        {
                            bool isBlockNeedGen = false;
                            if (GetDigCapDsscInstName(domain, block, text, type, "H", "L").Equals("TBD", StringComparison.OrdinalIgnoreCase) ||
                                GetDigCapDsscInstName(domain, block, text, type, "L", "H").Equals("TBD", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            instanceSheet.AddRow(WriteInstContent(domain, block, text, type, ref isBlockNeedGen));
                            instanceSheet.AddRow(WriteInstDigCapContent(domain, block, text, type, "H", "L", ref isBlockNeedGen));
                            instanceSheet.AddRow(WriteInstDigCapContent(domain, block, text, type, "L", "H", ref isBlockNeedGen));
                            instanceSheet.AddRow(WriteInstContentWithLoopCnt(domain, block, text, type));
                            if (isBlockNeedGen)
                            {
                                enableGenList.Add($"{domain}_{block}{text}_{type}");
                            }
                        }
                    }
                }
            }
            instanceSheet.AddRow(WriteInitSrmParseCkList());
            if (string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                instanceSheet.AddRow(WriteDatalogSetupSram());
                instanceSheet.AddRow(WriteSrmGetChannelType());
            }
        }

        private InstanceRow WriteDatalogSetupSram()
        {
            var instanceRow = new InstanceRow { TestName = ConDatalogSetupSram, VbtType = ConVbt, VbtName = ConVbtSrmInitDatalogSetup };

            return instanceRow;
        }

        private InstanceRow WriteInitSrmParseCkList()
        {
            var instanceRow = new InstanceRow { TestName = ConInitSrmParseCkList };

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ConCsSelsramParseTable, "", true);
            if (function.IsFound)
            {
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(ConVbtSrmParseChkList, "");
                function.ArgList[0] = ConSelsrmChkList;
            }

            instanceRow.VbtType = function.Type;
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;

            return instanceRow;
        }

        private InstanceRow WriteSrmGetChannelType()
        {
            var instanceRow = new InstanceRow { TestName = ConSrmGetChannelType, VbtType = ConVbt, VbtName = ConVbtSrmGetChnType };

            return instanceRow;
        }

        private InstanceRow WriteInstContentWithLoopCnt(string domain, string block, string chiplet, string type)
        {
            var instanceRow = new InstanceRow { TestName = $"{domain}{block}{chiplet}_SelSram_SetLoopCNT_{type}_NV" };

            string pattern = GetPatternByCondition(domain, block, chiplet, type);
            var payloadList = new List<string> { pattern };
            string payloadType = GetPayloadType(pattern);

            if (block.Equals(ConScan))
            {
                if (payloadType.Equals(""))
                {
                    payloadType = "Sa";
                }

                if (_multiTestSettingSheetsSingleton == null)
                {
                    instanceRow.DcCategory = $"{payloadType}_{domain}_X_X{chiplet}";
                }
                else
                {
                    instanceRow.DcCategory = _multiTestSettingSheetsSingleton.FindScanCategoryName(payloadType, domain, string.Empty, payloadList, out EnumMessageLevel _, out string _, string.Empty, chiplet: chiplet.TrimEnd('_'));
                }

                if (instanceRow.DcCategory.Equals(""))
                {
                    payloadType = "Sa";
                    instanceRow.DcCategory = _multiTestSettingSheetsSingleton?.FindScanCategoryName(payloadType, domain, string.Empty, payloadList, out EnumMessageLevel _, out string _, string.Empty, chiplet: chiplet.TrimEnd('_'));
                }
            }
            else if (block.Equals(ConMbist))
            {
                if (_multiTestSettingSheetsSingleton == null)
                {
                    instanceRow.DcCategory = $"{block}_{domain}_Init_X{chiplet}";
                }
                else
                {
                    instanceRow.DcCategory = _multiTestSettingSheetsSingleton.FindMbistCatgeoryName(domain, payloadType, string.Empty, payloadList, out EnumMessageLevel _, out string _, chiplet: chiplet.TrimEnd('_'));
                }
            }
            if (instanceRow.DcCategory != null && instanceRow.DcCategory.Equals(""))
            {
                instanceRow.DcCategory = "TBD";
            }

            instanceRow.DcSelector = "Typ";

            instanceRow.TimeSets = GetTimeSetName(pattern);
            string timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(instanceRow.TimeSets, BlockType.Scan);
            if (timeSet2Cat == "TBD")
            {
                timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(instanceRow.TimeSets);
            }

            instanceRow.AcCategory = timeSet2Cat;
            instanceRow.AcSelector = "Typ";

            instanceRow.PinLevels = $"{ConLevels}_{block}";

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ConCsSelsramInitLoopCount, "", true);
            if (function.IsFound)
            {
                function.SetParamValue("srmBlock",
                    $"{GetAbbreviation(domain)}_{GetAbbreviation(block)}{chiplet}_{type}");
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(ConVbtSrmInitLoopCnt, "");
                function.ArgList[0] = $"{GetAbbreviation(domain)}_{GetAbbreviation(block)}{chiplet}_{type}";
            }

            instanceRow.VbtType = function.Type;
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;
            if (pattern.Equals(""))
            {
                instanceRow.IsBackup = true;
            }

            return instanceRow;
        }

        internal string GetAbbreviation(string input)
        {
            if (input.Equals(ConSoc))
            {
                return "S";
            }

            if (input.Equals(ConCpu))
            {
                return "C";
            }

            if (input.Equals(ConGfx))
            {
                return "L";
            }

            if (input.Equals(ConScan))
            {
                return "SC";
            }

            if (input.Equals(ConMbist))
            {
                return "BI";
            }

            return input;
        }

        private string GetPatternByCondition(string domain, string block, string chiplet, string type)
        {
            string cat = GetAbbreviation(domain) + "_";
            string blk = "_" + GetAbbreviation(block) + "_";
            if (_selSramPattern.SelSramDatas.Count == 0)
            {
                _selSramPattern.FillData();
            }

            SelSramData ssData = _selSramPattern.SelSramDatas.FirstOrDefault(p => p.Category.StartsWith(cat) && p.Category.Contains(blk) && p.Type.Equals(type) && p.Category.EndsWith(chiplet, StringComparison.OrdinalIgnoreCase));
            return ssData != null ? ssData.Pattern : "";
        }

        internal string GetPayloadType(string pattern)
        {
            string lStrResult = "";
            if (PayloadTypeTable == null || pattern.Equals(""))
            {
                return "";
            }
            for (int i = 0; i < PayloadTypeTable.Rows.Count; i++)
            {
                bool lBIsMatch = true;
                for (int j = 1; j < PayloadTypeTable.Columns.Count; j++)
                {
                    string lStrMatchPattern = PayloadTypeTable.Rows[i][j].ToString();
                    string subName = GetSubName(pattern, PayloadTypeTable.Columns[j].ColumnName);
                    if (!Regex.IsMatch(subName, lStrMatchPattern, RegexOptions.IgnoreCase))
                    {
                        lBIsMatch = false;
                    }
                }

                if (lBIsMatch)
                {
                    lStrResult = PayloadTypeTable.Rows[i][0].ToString();
                    break;
                }
            }

            return lStrResult;
        }

        internal string GetSubName(string name, string rule)
        {
            var resultList = new List<string>();
            string[] words = name.Split('_');
            string[] numbersStrings = rule.Split(',');
            foreach (string numbers in numbersStrings)
            {
                if (rule.ToLower() == "full")
                {
                    return name;
                }
                if (rule == "")
                {
                    //no operation
                }
                else
                {
                    int getNumber = int.Parse(numbers);
                    if (words.Length > getNumber && getNumber >= 0)
                    {
                        //resultName = resultName + "_" + words[getNumber];
                        resultList.Add(words[getNumber]);

                    }
                }
            }
            string resultName = string.Join("_", resultList);

            return resultName;
        }

        private void GetInstanceInfo(ref InstanceRow instanceRow, string domain, string block, string chiplet, string type)
        {
            string pattern = GetPatternByCondition(domain, block, chiplet, type);
            var payloadList = new List<string> { pattern };
            string payloadType = GetPayloadType(pattern);

            if (block.Equals(ConScan))
            {
                if (payloadType.Equals(""))
                {
                    payloadType = "Sa";
                }

                if (_multiTestSettingSheetsSingleton == null)
                {
                    instanceRow.DcCategory = $"{payloadType}_{domain}_X_X";
                }
                else
                {
                    instanceRow.DcCategory = _multiTestSettingSheetsSingleton.FindScanCategoryName(payloadType, domain, string.Empty, payloadList, out EnumMessageLevel _, out string _, string.Empty);
                }

                if (instanceRow.DcCategory.Equals(""))
                {
                    payloadType = "Sa";
                    instanceRow.DcCategory = _multiTestSettingSheetsSingleton?.FindScanCategoryName(payloadType, domain, string.Empty, payloadList, out EnumMessageLevel _, out string _, string.Empty);
                }
            }
            else if (block.Equals(ConMbist))
            {
                if (_multiTestSettingSheetsSingleton == null)
                {
                    instanceRow.DcCategory = $"{block}_{domain}_Init_X";
                }
                else
                {
                    instanceRow.DcCategory = _multiTestSettingSheetsSingleton.FindMbistCatgeoryName(domain, payloadType, string.Empty, payloadList, out EnumMessageLevel _, out string _);
                }
            }

            if (instanceRow.DcCategory != null && instanceRow.DcCategory.Equals(""))
            {
                instanceRow.DcCategory = "TBD";
            }

            instanceRow.DcSelector = "Typ";
            instanceRow.TimeSets = GetTimeSetName(pattern);

            string timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(instanceRow.TimeSets, BlockType.Scan);
            if (timeSet2Cat == "TBD")
            {
                timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(instanceRow.TimeSets);
            }

            instanceRow.AcCategory = timeSet2Cat;
            instanceRow.AcSelector = "Typ";

            instanceRow.PinLevels = $"{ConLevels}_{block}";
        }

        private InstanceRow WriteInstContent(string domain, string block, string chiplet, string type, ref bool isNeedGen)
        {
            var instanceRow = new InstanceRow { TestName = $"{domain}{block}{chiplet}_SelSram_{type}" };

            GetInstanceInfo(ref instanceRow, domain, block, chiplet, type);

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ConCsSelsramFuncTestMain, "", true);
            if (function.IsFound)
            {
                var patternArgNameList = new List<string>
                {
                    "initPattern01",
                    "initPattern02",
                    "initPattern03",
                    "initPattern04",
                    "initPattern05",
                    "initPattern06",
                    "initPattern07",
                    "initPattern08",
                    "initPattern09",
                    "initPattern10",
                    "payloadPattern01",
                    "payloadPattern02",
                    "payloadPattern03",
                    "payloadPattern04",
                    "payloadPattern05",
                };
                function.SetParamValue("patternTimeout", "30");
                function.SetParamValue("srmType", $"{GetAbbreviation(domain)}_{GetAbbreviation(block)}{chiplet}_{type}");
                string blockKey = $"{GetAbbreviation(domain)}_{GetAbbreviation(block)}_PC{chiplet}_{type}";
                if (_selSramPattern.DicInitPatterns.TryGetValue(blockKey, out List<string> initPatList))
                {
                    isNeedGen = isNeedGen || !LocalSpecs.IsPatternValidate ||
                                initPatList.Any(p => UsePatternList.Any(q => q.Equals(p, StringComparison.OrdinalIgnoreCase)));
                    for (int index = 0; index < initPatList.Count; index++)
                    {
                        string pat = initPatList[index];
                        if (index >= 0 && index < patternArgNameList.Count)
                        {
                            string argName = patternArgNameList[index];
                            function.SetParamValue(argName, pat);
                        }
                    }
                }
                else
                {
                    instanceRow.IsBackup = true;
                }
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(ConVbtFunctionalTSrm, "");

                function.ArgList[0] = "30";
                string blockKey = $"{GetAbbreviation(domain)}_{GetAbbreviation(block)}_PC{chiplet}_{type}";
                if (_selSramPattern.DicInitPatterns.TryGetValue(blockKey, out List<string> initPatList))
                {
                    isNeedGen = isNeedGen || !LocalSpecs.IsPatternValidate ||
                                initPatList.Any(p => UsePatternList.Any(q => q.Equals(p, StringComparison.OrdinalIgnoreCase)));
                    for (int index = 0; index < initPatList.Count; index++)
                    {
                        string pat = initPatList[index];
                        function.ArgList[index + 1] = pat;
                    }
                }
                else
                {
                    instanceRow.IsBackup = true;
                }

                function.ArgList[22] = $"{GetAbbreviation(domain)}_{GetAbbreviation(block)}{chiplet}_{type}";
            }

            instanceRow.VbtType = function.Type;
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;
            return instanceRow;
        }

        private string GetBitsLogicPinsString(string logicHl, string sramHl)
        {
            var retList = new List<string>();
            HashSet<string> logicPins = SelSramPatternSingleton.GetInstance().BitsLogicPins;
            for (int i = 0; i < logicPins.Count; ++i)
            {
                retList.Add(logicPins.ElementAt(i));
            }
            retList.Add("VDD_SRAM_GPU");
            if (logicHl.Equals("H"))
            {
                retList.Add("VDD_SRAM_SOC|logic:+0.2;sram:-0.2");
            }
            else if (sramHl.Equals("H"))
            {
                retList.Add("VDD_SRAM_SOC|logic:-0.2;sram:+0.2");
            }

            return string.Join(",", retList);
        }

        private InstanceRow WriteInstDigCapContent(string domain, string block, string chiplet, string type, string logicHl, string sramHl, ref bool isNeedGen)
        {
            var instanceRow = new InstanceRow { TestName = $"{domain}{block}{chiplet}_SelSram_Digcap_{type}_Logic{logicHl}_Sram{sramHl}_NV" };
            string key = domain.ToUpper().Trim() + block.ToUpper().Trim() + chiplet;
            string pattern = _selSramPattern.DicReadbackPat.TryGetValue(key, out string value) ? value : "";
            var payloadList = new List<string> { pattern };
            string payloadType = GetPayloadType(pattern);
            isNeedGen = isNeedGen || !LocalSpecs.IsPatternValidate || UsePatternList.Exists(p => p.Equals(pattern, StringComparison.OrdinalIgnoreCase));
            if (block.Equals(ConScan))
            {
                if (payloadType.Equals(""))
                {
                    payloadType = "Sa";
                }

                if (_multiTestSettingSheetsSingleton == null)
                {
                    instanceRow.DcCategory = $"{payloadType}_{domain}_X_X{chiplet}";
                }
                else
                {
                    instanceRow.DcCategory = _multiTestSettingSheetsSingleton.FindScanCategoryName(payloadType, domain, string.Empty, payloadList, out EnumMessageLevel _, out string _, string.Empty, chiplet: chiplet.Trim('_'));
                }

                if (instanceRow.DcCategory.Equals(""))
                {
                    payloadType = "Sa";
                    instanceRow.DcCategory = _multiTestSettingSheetsSingleton?.FindScanCategoryName(payloadType, domain, string.Empty, payloadList, out EnumMessageLevel _, out string _, string.Empty, chiplet: chiplet.Trim('_'));
                }
            }
            else if (block.Equals(ConMbist))
            {
                if (_multiTestSettingSheetsSingleton == null)
                {
                    instanceRow.DcCategory = $"{block}_{domain}_Init_X{chiplet}";
                }
                else
                {
                    instanceRow.DcCategory = _multiTestSettingSheetsSingleton.FindMbistCatgeoryName(domain, payloadType, string.Empty, payloadList, out EnumMessageLevel _, out string _, chiplet: chiplet.Trim('_'));
                }
            }
            if (instanceRow.DcCategory != null && instanceRow.DcCategory.Equals(""))
            {
                instanceRow.DcCategory = "TBD";
            }

            instanceRow.DcSelector = "Typ";
            instanceRow.TimeSets = GetTimeSetName(pattern);

            string timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(instanceRow.TimeSets, BlockType.Scan);
            if (timeSet2Cat == "TBD")
            {
                timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(instanceRow.TimeSets);
            }

            instanceRow.AcCategory = timeSet2Cat;
            instanceRow.AcSelector = "Typ";

            instanceRow.PinLevels = $"{ConLevels}_{block}";
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ConCsSelsramFuncTestMain, "", true);
            if (function.IsFound)
            {
                function.SetParamValue("patternTimeout", "30");
                function.SetParamValue("initPattern01", pattern);
                function.SetParamValue("cgVoltage", GetBitsLogicPinsString(logicHl, sramHl));
                function.SetParamValue("srmType", "SRMREAD");
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(ConVbtFunctionalTSrm, "");
                function.ArgList[0] = "30";
                function.ArgList[1] = pattern;
                function.ArgList[20] = GetBitsLogicPinsString(logicHl, sramHl);
                function.ArgList[22] = "SRMREAD";
            }


            instanceRow.VbtType = function.Type;
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;

            if (pattern.Equals(""))
            {
                instanceRow.IsBackup = true;
            }

            return instanceRow;
        }

        public void WriteFlowSheet(ref SubFlowSheet flowSheet, List<string> enableGenList)
        {
            List<string> list;
            if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                list = new List<string> { ConInitSrmParseCkList };
            }
            else
            {
                list = new List<string> { ConDatalogSetupSram, ConInitSrmParseCkList, ConSrmGetChannelType };
            }

            foreach (string setup in list)
            {
                if ((setup.Equals(ConDatalogSetupSram) || setup.Equals(ConSrmGetChannelType)) && !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
                {
                    continue;
                }

                var flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = setup };
                flowSheet.AddRow(flowRow);
            }

            foreach (string chiplet in SelSramPatternSingleton.GetInstance().Chiplet)
            {
                string text = chiplet;
                if (!string.IsNullOrEmpty(text))
                {
                    text = "_" + text;
                }

                foreach (string domain in _domainList)
                {
                    foreach (string type in _hcDsscList)
                    {
                        foreach (string block in _blockList)
                        {
                            if (
                                GetDigCapDsscInstName(domain, block, text, type, "H", "L")
                                    .Equals("TBD", StringComparison.OrdinalIgnoreCase) ||
                                GetDigCapDsscInstName(domain, block, text, type, "L", "H")
                                    .Equals("TBD", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            if (!enableGenList.Contains($"{domain}_{block}{text}_{type}"))
                            {
                                continue;
                            }

                            WriteFlowContent(ref flowSheet, domain, block, text, type);
                        }
                    }
                }
            }

            var flowRowReturn = new FlowRow { Opcode = "Return" };
            flowSheet.AddRow(flowRowReturn);
        }

        private string GetDigCapDsscInstName(string domain, string block, string chiplet, string type, string logicHl, string sramHl)
        {
            var retList = new List<string>();
            var instance = SelSramPatternSingleton.GetInstance();
            string readbackPattern = "TBD";
            if (instance.DicReadbackPat.ContainsKey((domain + block + chiplet).ToUpper()))
            {
                readbackPattern = instance.DicReadbackPat[(domain + block + chiplet).ToUpper()];
            }

            string[] spt = readbackPattern.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (spt.Length < 5)
            {
                return readbackPattern;
            }

            string token = string.Empty;
            switch (spt[2])
            {
                case "L":
                    token = ConGfx;
                    break;
                case "S":
                    token = ConSoc;
                    break;
                case "C":
                    token = ConCpu;
                    break;
            }
            switch (spt[4])
            {
                case "BI":
                    token += ConMbist;
                    break;
                case "SC":
                    token += ConScan;
                    break;
            }
            retList.Add(token);
            retList.Add("SelSram");
            retList.Add("Digcap");
            retList.Add(type.ToUpper());
            retList.Add("Logic" + logicHl);
            retList.Add("Sram" + sramHl);
            retList.Add("NV");
            return string.Join("_", retList);
        }

        private void WriteFlowContent(ref SubFlowSheet flowSheet, string domain, string block, string chiplet, string type)
        {
            var flowRow = new FlowRow { Opcode = "print", Parameter = $"\"======== Start of {domain}{block}{chiplet} {type} SelSram Test ========\"" };
            flowSheet.AddRow(flowRow);

            flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = $"{domain}{block}{chiplet}_SelSram_SetLoopCNT_{type}_NV" };
            flowSheet.AddRow(flowRow);

            flowRow = new FlowRow { Opcode = "For", Parameter = "SEL_CASE_CNT = 0; SEL_CASE_CNT < SEL_CASE_End; SEL_CASE_CNT++" };
            flowSheet.AddRow(flowRow);

            flowRow = new FlowRow
            {
                Opcode = "For",
                Parameter = string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder) ?
                "SelSrm_LP_Var = 0; SelSrm_LP_Var < LPCount_End; SelSrm_LP_Var++" :
                "SEL_PAT_CNT = 0; SEL_PAT_CNT < SEL_PAT_END; SEL_PAT_CNT++"
            };
            flowSheet.AddRow(flowRow);

            flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = $"{domain}{block}{chiplet}_SelSram_{type}" };
            flowSheet.AddRow(flowRow);

            flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = $"{domain}{block}{chiplet}_SelSram_Digcap_{type}_Logic{"H"}_Sram{"L"}_NV" };
            //GetDigcapDsscInstName(domain, block, type, "H", "L");
            flowSheet.AddRow(flowRow);

            flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = $"{domain}{block}{chiplet}_SelSram_{type}" };
            flowSheet.AddRow(flowRow);

            flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = $"{domain}{block}{chiplet}_SelSram_Digcap_{type}_Logic{"L"}_Sram{"H"}_NV" };
            //GetDigcapDsscInstName(domain, block, type, "L", "H");
            flowSheet.AddRow(flowRow);

            flowRow = new FlowRow { Opcode = "Next" };
            flowSheet.AddRow(flowRow);

            flowRow = new FlowRow { Opcode = "Next" };
            flowSheet.AddRow(flowRow);

            flowRow = new FlowRow { Opcode = "print", Parameter = $"\"======== End of {domain}{block}{chiplet} {type} SelSram Test ========\"" };
            flowSheet.AddRow(flowRow);
        }
    }
}
