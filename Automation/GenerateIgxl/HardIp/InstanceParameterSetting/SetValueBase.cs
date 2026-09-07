using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting
{
    public abstract class SetValueBase
    {
        protected HardIpInputData HardIpInputData { get; }
        public HardIpSheet HardIpSheet { get; }

        private const string TestSequence = "TESTSEQUENCE";
        private const string CusStrDigcapdata = "CUS_STR_DIGCAPDATA";
        private const string DigsrcEquation = "DIGSRC_EQUATION";
        private const string DigsrcAssignment = "DIGSRC_ASSIGNMENT";
        private const string CalcEqn = "CALC_EQN";
        private const string MeasiRange = "MEASI_RANGE";
        private const string ForcevVal = "FORCEV_VAL";
        private const string ForceiVal = "FORCEI_VAL";
        private const string CusStrMainprogram = "CUS_STR_MAINPROGRAM";

        private const string MCatename = "M_CATENAME";
        private const string DictStoreCodeName = "DICT_STORE_CODE_NAME";

        public List<string> ReservedMiscInfoKeys { get; set; } = new List<string>();

        protected SetValueBase(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet)
        {
            HardIpInputData = hardIpInputData;
            HardIpSheet = hardIpSheet;
        }

        public abstract void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage);

        public virtual void SetValueByParamMapping(Function function, HardIpPattern pattern, string voltage)
        {
            Dictionary<string, List<string>> miscInfoDic = pattern.MiscInfoDict.ToDictionary(x => x.Key, x => x.Value.Split(';').ToList());
            if (VbtFunctionLibShared.EfusePrewriteFunctionList.Exists(f => f.Equals(pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase)) &&
                (pattern.SheetName.Equals("DCTEST_IDS", StringComparison.CurrentCultureIgnoreCase) || voltage.Equals(HardIpConstData.LabelNv, StringComparison.CurrentCultureIgnoreCase)))
            {
                if (!miscInfoDic.ContainsKey("Flag_Name"))
                {
                    miscInfoDic.Add("Flag_Name", new List<string> { pattern.HipPreWriteFlag });
                }

                if (!miscInfoDic.ContainsKey("flagName"))
                {
                    miscInfoDic.Add("flagName", new List<string> { pattern.HipPreWriteFlag });
                }

                ErrorReportManager.AddError(
                    HardIpErrorType.W_FuseWriteJudgeFlag_01,
                    pattern.SheetName,
                    pattern.RowNum,
                    1,
                    [pattern.HipPreWriteFlag]
                );
            }

            List<string> catenames = new List<string>();
            string dspwavesize = "";
            ResolveEfuseCatenames(pattern, miscInfoDic, ref catenames, ref dspwavesize);

            foreach (KeyValuePair<string, List<string>> param in miscInfoDic)
            {
                ApplyMiscInfoParam(function, pattern, voltage, param, catenames, dspwavesize);
            }
        }

        private void ResolveEfuseCatenames(HardIpPattern pattern, Dictionary<string, List<string>> miscInfoDic, ref List<string> catenames, ref string dspwavesize)
        {
            if (pattern.FunctionName.Equals("HIP_eFuse_Read", StringComparison.OrdinalIgnoreCase) ||
                pattern.FunctionName.Equals("HardIPFuseRead", StringComparison.OrdinalIgnoreCase))
            {
                if (HardIpInputData.EfuseBitDefBw == null)
                {
                    ErrorReportManager.AddError(HardIpErrorType.E_HIPeFuseInput_03, pattern.SheetName, pattern.RowNum, 1, [pattern.FunctionName]);
                }
                else if (!miscInfoDic.Any(misc => misc.Key.ToLower().Equals("m_catename") || misc.Key.ToLower().Equals("fieldName".ToLower())))
                {
                    ErrorReportManager.AddError(HardIpErrorType.E_HIPeFuseInput_04, pattern.SheetName, pattern.RowNum, 1, [pattern.FunctionName]);
                }
                else if (!miscInfoDic.Any(misc => misc.Key.ToUpper().Equals("DSPWAVESIZE") || misc.Key.ToLower().Equals("sampleSize".ToLower())))
                {
                    ErrorReportManager.AddError(HardIpErrorType.E_HIPeFuseInput_05, pattern.SheetName, pattern.RowNum, 1, [pattern.FunctionName]);
                }
                else
                {
                    catenames = miscInfoDic.FirstOrDefault(misc => misc.Key.Equals("m_catename", StringComparison.CurrentCultureIgnoreCase) || misc.Key.Equals("fieldName", StringComparison.CurrentCultureIgnoreCase))
                        .Value[0].Split('+').ToList();
                    dspwavesize = miscInfoDic.FirstOrDefault(misc => misc.Key.Equals("DSPWAVESIZE", StringComparison.CurrentCultureIgnoreCase) || misc.Key.Equals("sampleSize", StringComparison.CurrentCultureIgnoreCase))
                        .Value[0];
                }
            }
            else if (pattern.FunctionName.Equals("HIP_eFuse_Write", StringComparison.OrdinalIgnoreCase) ||
                pattern.FunctionName.Equals("HardIPFuseWrite", StringComparison.OrdinalIgnoreCase))
            {
                if (HardIpInputData.EfuseBitDefBw == null)
                {
                    ErrorReportManager.AddError(HardIpErrorType.E_HIPeFuseInput_06, pattern.SheetName, pattern.RowNum, 1, [pattern.FunctionName]);
                }
                else if (!miscInfoDic.Any(misc => misc.Key.ToLower().Equals("m_catename") || misc.Key.ToLower().Equals("fieldName".ToLower())))
                {
                    ErrorReportManager.AddError(HardIpErrorType.E_HIPeFuseInput_04, pattern.SheetName, pattern.RowNum, 1, [pattern.FunctionName]);
                }
                else
                {
                    catenames = miscInfoDic.FirstOrDefault(misc => misc.Key.Equals("m_catename", StringComparison.CurrentCultureIgnoreCase) || misc.Key.Equals("fieldName", StringComparison.CurrentCultureIgnoreCase)).Value[0].Split('+').ToList();
                }
            }
        }

        private void ApplyMiscInfoParam(Function function, HardIpPattern pattern, string voltage, KeyValuePair<string, List<string>> param, List<string> catenames, string dspwavesize)
        {
            if (param.Key.Equals("m_catename", StringComparison.CurrentCultureIgnoreCase) || param.Key.Equals("fieldName", StringComparison.CurrentCultureIgnoreCase))
            {
                ValidateCatenamesAgainstBitDef(pattern, catenames, dspwavesize);
            }

            bool paramFound = false;
            if (!HardIpConstData.MiscKeyList.Exists(s => Regex.IsMatch(param.Key, "^(" + s + ")$", RegexOptions.IgnoreCase)))
            {
                int index = function.Parameters.Split(',').ToList().FindIndex(s => s.Equals(param.Key, StringComparison.OrdinalIgnoreCase));
                if (index != -1)
                {
                    if (param.Key.Equals("CUS_Str_MainProgram", StringComparison.CurrentCultureIgnoreCase) || param.Key.Equals("mainProgramCustomString", StringComparison.CurrentCultureIgnoreCase))
                    {
                        //If the values are inconsistent, leave blank and hightlight wrong keyword in miscinfo.
                        if (function.ArgList[index].Any())
                        {
                            function.ArgList[index] = function.ArgList[index] + ";" + string.Join(";", param.Value);
                        }
                        else
                        {
                            function.ArgList[index] = string.Join(";", param.Value);
                        }
                        paramFound = true;
                    }
                    else if (param.Key.Equals("MeasName") && LocalSpecs.Options.Device == EnumDevice.RF)
                    {
                        List<string> allTn = new List<string>();
                        foreach (string singleTn in string.Join("|", param.Value).Split('|'))
                        {
                            string[] spiTn = singleTn.Split('_');
                            if (spiTn.Length > 8)
                            {
                                spiTn[8] = voltage;
                            }

                            allTn.Add(string.Join("_", spiTn));
                        }
                        function.ArgList[index] = string.Join("|", allTn);
                    }
                    else if (LocalSpecs.Options.Device == EnumDevice.RF && function.FunctionName.ToLower() == VbtFunctionLibShared.VifName.ToLower())
                    {
                        function.ArgList[index] = function.ArgList[index].Trim(';') + ";" + string.Join(";", param.Value);
                        function.ArgList[index] = function.ArgList[index].Trim().Trim(';');
                        paramFound = true;
                    }
                    else
                    {
                        function.SetParamValue(param.Key, string.Join(";", param.Value));
                        paramFound = true;
                    }
                }
                else
                {
                    IEnumerable<List<string>> aliasList = VbtFunctionLibShared.ParamMappingList.Where(s => s.Find(a => a.Equals(param.Key, StringComparison.OrdinalIgnoreCase)) != null);
                    if (aliasList != null)
                    {
                        foreach (List<string> aliasName in aliasList)
                        {
                            foreach (string item in aliasName)
                            {
                                int newIndex = function.Parameters.Split(',').ToList().FindIndex(s => s.Equals(item, StringComparison.OrdinalIgnoreCase));
                                if (newIndex != -1)
                                {
                                    function.ArgList[newIndex] = string.Join(";", param.Value);
                                    paramFound = true;
                                    break;
                                }
                            }
                        }
                    }
                }
                if (!paramFound && !ReservedMiscInfoKeys.Exists(s => Regex.IsMatch(param.Key, "^(" + s + ")$", RegexOptions.IgnoreCase)))
                {
                    int miscInfoIndex = HardIpSheet.PlanHeaderIdx["miscInfoIndex"];
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_MissingParameter_01,
                        pattern.SheetName,
                        pattern.RowNum,
                        miscInfoIndex,
                        [function.FunctionName, param.Key]
                    );
                }
            }
        }

        private void ValidateCatenamesAgainstBitDef(HardIpPattern pattern, List<string> catenames, string dspwavesize)
        {
            for (int i = 0; i < catenames.Count; i++)
            {
                string targetKey = HardIpInputData.EfuseBitDefBw.Keys.FirstOrDefault(x => x.Equals(catenames[i], StringComparison.OrdinalIgnoreCase));
                if (string.IsNullOrEmpty(targetKey))
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_HIPeFuseCatename_01,
                        pattern.SheetName,
                        pattern.RowNum,
                        1,
                        [catenames[i]]
                    );
                }
                else if (dspwavesize != "")
                {
                    if (!int.TryParse(dspwavesize, out int _))
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_HIPeFuseDSPWAVE_01,
                            pattern.SheetName,
                            pattern.RowNum,
                            1,
                            [dspwavesize]
                        );
                    }
                    else if (!HardIpInputData.EfuseBitDefBw[targetKey].Equals(dspwavesize))
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_HIPeFuseDSPWAVE_02,
                            pattern.SheetName,
                            pattern.RowNum,
                            1,
                            [catenames[i], dspwavesize, HardIpInputData.EfuseBitDefBw[targetKey]]
                        );
                    }
                }
            }
        }

        public void CheckInstArgument(Function function, HardIpPattern pattern)
        {
            if (pattern.Pattern.IsMultiTimeDomain())
            {
                return;
            }

            for (int i = 0; i < function.ArgList.Count; i++)
            {
                if (function.ArgList[i].Length > 8000)
                {
                    string errorMessage = "IG-XL 9.0 need to keep the String length < 8000 characters ";
                    if (!ErrorReportManager.GetErrorList().Any(x => x.SheetName == pattern.SheetName && x.RowNum == pattern.RowNum
                        && x.Message == errorMessage))
                    {
                        string argStr = function.Parameters.Split(',')[i];
                        string valStr = function.ArgList[i];
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_IgxlLimitation_01,
                            pattern.SheetName,
                            pattern.RowNum,
                            0,
                            [argStr, valStr]
                        );
                    }
                    //todo need check whether 8000 is correct
                    if (function.ArgList[i].Length > 8000)
                    {
                        function.ArgList[i] = function.ArgList[i].Substring(0, 8000);
                    }
                }
            }
        }

        public string Filter(string item)
        {
            if (item == "|" || item == "|,")
            {
                return "";
            }

            return item;
        }

        internal void SetInterPoseFunc(HardIpPattern pattern, Function function, string voltage, string keepForcePostPat = "")
        {
            if (string.IsNullOrEmpty(pattern.RfInterPose) && string.IsNullOrEmpty(pattern.MiscInfo) && pattern.MiscInfo.Split(';').ToList().Exists(x => x.StartsWith("interpose", StringComparison.CurrentCultureIgnoreCase)))
            {
                return;
            }

            string blockName = CommonGenerator.GetRegAssignName(pattern);
            if (voltage.Equals("NV", StringComparison.CurrentCultureIgnoreCase)) //temp
            {
                foreach (string interposeFunc in ComposeOrderAssignList(pattern, keepForcePostPat))
                {
                    if (!interposeFunc.Contains(":"))
                    {
                        continue;
                    }

                    //generate assign name 
                    string[] paraInfo = interposeFunc.Split(':');
                    string name = paraInfo[0].Trim();
                    string valuelist = paraInfo[1].Trim();
                    foreach (string value in SplitMultipleInterpose(name, valuelist))
                    {
                        AddInterposeAssign(pattern, function, blockName, name, value);
                    }
                }
            }

            var assign = HardIpInputData.InterposeAssigns.Where(x => x.BlockName.Equals(blockName, StringComparison.CurrentCultureIgnoreCase)).OrderBy(x => (int)x.Type).ToList();

            foreach (InterposeAssign item in assign)
            {
                ApplyInterposeAssignToArgs(pattern, function, item);
            }
        }

        private void AddInterposeAssign(HardIpPattern pattern, Function function, string blockName, string name, string value)
        {
            InterposeAssignType interposeType = InterposeAssignType.None;
            if (name.Equals(nameof(InterposeAssignType.InterposePreInit), StringComparison.OrdinalIgnoreCase) || name.Equals(nameof(InterposeAssignType.InterposePrePat), StringComparison.OrdinalIgnoreCase) || name.Equals(nameof(InterposeAssignType.InterposePostInit), StringComparison.OrdinalIgnoreCase))
            {
                interposeType = function.FunctionName.Equals(VbtFunctionLibShared.FunctionalName, StringComparison.CurrentCultureIgnoreCase) ? InterposeAssignType.StartOfBodyFArgs : InterposeAssignType.InterposePrePat;
            }
            else if (name.Equals(nameof(InterposeAssignType.InterposePreRst), StringComparison.OrdinalIgnoreCase) || name.Equals(nameof(InterposeAssignType.InterposePostPat), StringComparison.OrdinalIgnoreCase) || name.Equals(nameof(InterposeAssignType.InterposePostRst), StringComparison.OrdinalIgnoreCase))
            {
                interposeType = function.FunctionName.Equals(VbtFunctionLibShared.FunctionalName, StringComparison.CurrentCultureIgnoreCase) ? InterposeAssignType.EndOfBodyFArgs : InterposeAssignType.InterposePostPat;
            }
            else if (name.Equals(nameof(InterposeAssignType.InterposePreMeas), StringComparison.OrdinalIgnoreCase))
            {
                interposeType = InterposeAssignType.InterposePreMeas;
            }
            else if (name.Equals(nameof(InterposeAssignType.InterposePostMeas), StringComparison.OrdinalIgnoreCase))
            {
                interposeType = InterposeAssignType.InterposePostMeas;
            }

            string assignName = GenerateAssignName(blockName, interposeType, pattern.PatternIndexFlag, function);
            InterposeAssign interposeAssign = null;
            if (!HardIpInputData.InterposeAssigns.Any(x => x.AssignName.Equals(assignName, StringComparison.CurrentCultureIgnoreCase)))
            {
                interposeAssign = new InterposeAssign
                {
                    BlockName = blockName,
                    Type = interposeType,
                    AssignName = assignName
                };

            }
            else
            {
                interposeAssign = HardIpInputData.InterposeAssigns.FirstOrDefault(x => x.AssignName.Equals(assignName, StringComparison.CurrentCultureIgnoreCase));
                var newBag = new ConcurrentBag<InterposeAssign>(HardIpInputData.InterposeAssigns.Where(x => x != interposeAssign));
                HardIpInputData.InterposeAssigns = newBag;
            }
            interposeAssign.InterposeAssignList.Add(string.IsNullOrEmpty(value) ? "NA" : value);
            HardIpInputData.InterposeAssigns.Add(interposeAssign);
        }

        private void ApplyInterposeAssignToArgs(HardIpPattern pattern, Function function, InterposeAssign item)
        {
            int index = function.Parameters.Split(',').ToList().FindIndex(s => s.Equals(item.Type.ToString(), StringComparison.OrdinalIgnoreCase));
            if (index != -1)
            {
                function.ArgList[index] += item.AssignName;
                if (function.FunctionName.Equals(VbtFunctionLibShared.FunctionalName, StringComparison.CurrentCultureIgnoreCase))
                {
                    //Only for Functional_t_updated args
                    string searchKey = item.Type.ToString().Equals(InterposeAssignType.StartOfBodyFArgs) ? "StartOfBodyF" : "EndOfBodyF";
                    index = function.Parameters.Split(',').ToList().FindIndex(s => s.Equals(searchKey));
                    if (index != -1)
                    {
                        function.ArgList[index] = "LCDInterpose";
                    }
                }
                List<string> miscInfoList = new List<string>();
                foreach (string misc in pattern.MiscInfo.Split(';'))
                {
                    if (!misc.ToLower().StartsWith("interpose"))
                    {
                        miscInfoList.Add(misc);
                    }
                }
                pattern.MiscInfo = string.Join(";", miscInfoList);
            }
            else
            {
                int miscInfoIndex = HardIpSheet.PlanHeaderIdx["miscInfoIndex"];
                string errorMessage = "Missing Parameter in " + function.FunctionName + " : " + item.Type + " Or wrong key word in misc-info";
                ErrorReportManager.AddError(
                    HardIpErrorType.E_MissingParameter_01,
                    pattern.SheetName,
                    pattern.RowNum,
                    miscInfoIndex,
                    [function.FunctionName, $"{item.Type}"]
                );
            }
        }

        private string GenerateAssignName(string blockName, InterposeAssignType interposeType, string patternIdxFlag, Function function)
        {
            if (function.FunctionName.Equals(VbtFunctionLibShared.FunctionalName, StringComparison.CurrentCultureIgnoreCase))
            {
                if (interposeType.Equals(InterposeAssignType.StartOfBodyFArgs))
                {
                    return blockName + Regex.Replace(interposeType.ToString(), "StartOfBodyFArgs*", "_StartOfBody", RegexOptions.IgnoreCase) + patternIdxFlag;
                }

                return blockName + Regex.Replace(interposeType.ToString(), "EndOfBodyFArgs*", "_EndOfBody", RegexOptions.IgnoreCase) + patternIdxFlag;
            }

            return blockName + Regex.Replace(interposeType.ToString(), "Interpose*", "_", RegexOptions.IgnoreCase) + patternIdxFlag;
        }

        private List<string> SplitMultipleInterpose(string key, string para)
        {
            if (key.Equals(nameof(InterposeAssignType.InterposePreMeas), StringComparison.CurrentCultureIgnoreCase) || key.Equals(nameof(InterposeAssignType.InterposePostMeas), StringComparison.CurrentCultureIgnoreCase))
            {
                return para.Split('|').ToList();
            }

            return para.Split('&').ToList();
        }

        private List<string> ComposeOrderAssignList(HardIpPattern pattern, string keepForcePostPat)
        {
            List<string> resultList = new List<string>();

            List<string> interposeList = pattern.RfInterPose.Split(';').ToList();
            IEnumerable<string> interposeFromMisc = pattern.MiscInfo.Split(';').ToList().Where(x => x.ToLower().StartsWith("interpose"));

            resultList.AddRange(interposeList.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePreInit).ToUpper())));
            resultList.AddRange(interposeFromMisc.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePreInit).ToUpper())));

            resultList.AddRange(interposeList.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePrePat).ToUpper())));
            resultList.AddRange(interposeFromMisc.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePrePat).ToUpper())));

            resultList.AddRange(interposeList.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePostInit).ToUpper())));
            resultList.AddRange(interposeFromMisc.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePostInit).ToUpper())));

            resultList.AddRange(interposeList.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePreRst).ToUpper())));
            resultList.AddRange(interposeFromMisc.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePreRst).ToUpper())));

            resultList.AddRange(interposeList.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePostPat).ToUpper())));
            resultList.AddRange(interposeFromMisc.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePostPat).ToUpper())));
            if (!string.IsNullOrEmpty(keepForcePostPat))
            {
                resultList.Add("InterposePostPat:" + keepForcePostPat);
            }

            resultList.AddRange(interposeList.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePostRst).ToUpper())));
            resultList.AddRange(interposeFromMisc.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePostRst).ToUpper())));

            resultList.AddRange(interposeList.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePreMeas).ToUpper())));
            resultList.AddRange(interposeFromMisc.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePreMeas).ToUpper())));

            resultList.AddRange(interposeList.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePostMeas).ToUpper())));
            resultList.AddRange(interposeFromMisc.Where(x => x.ToUpper().StartsWith(nameof(InterposeAssignType.InterposePostMeas).ToUpper())));
            return resultList;
        }

        protected string RemoveUnit(string strIn)
        {
            return Regex.Replace(strIn, "@(Hz)|V|A|ohm", "", RegexOptions.IgnoreCase);
        }

        private void ProcessHipEfuseRegisterAssignments(Function function, HardIpPattern pattern, string blockName, string subBlockName, List<string> parameters)
        {
            List<string> mcatenameInfo = new List<string>();
            List<string> dictStoreCodeNameInfo = new List<string>();
            for (int i = 0; i < parameters.Count; i++)
            {
                string argtype = parameters[i];
                string arginfo = function.ArgList[i];
                if (!string.IsNullOrEmpty(arginfo) && arginfo.StartsWith("+"))
                {
                    function.ArgList[i] = "'" + function.ArgList[i];
                    arginfo = function.ArgList[i];
                }

                if (argtype.Equals(MCatename, StringComparison.CurrentCultureIgnoreCase) || argtype.Equals("fieldName", StringComparison.CurrentCultureIgnoreCase))
                {
                    mcatenameInfo = arginfo.Split('+').ToList();
                    function.ArgList[i] = $"Reg_assign:{subBlockName}";
                }
                else if (argtype.Equals(DictStoreCodeName, StringComparison.CurrentCultureIgnoreCase) || argtype.Equals("dictionaryName", StringComparison.CurrentCultureIgnoreCase))
                {
                    dictStoreCodeNameInfo = arginfo.Split('+').ToList();
                    function.ArgList[i] = $"Reg_assign:{subBlockName}";
                }
            }

            var regAssginList = new List<List<string>>();
            if (mcatenameInfo.Count < dictStoreCodeNameInfo.Count)
            {
                mcatenameInfo.AddRange(Enumerable.Repeat("", dictStoreCodeNameInfo.Count - mcatenameInfo.Count));
            }
            else if (mcatenameInfo.Count > dictStoreCodeNameInfo.Count)
            {
                dictStoreCodeNameInfo.AddRange(Enumerable.Repeat("", mcatenameInfo.Count - dictStoreCodeNameInfo.Count));
            }
            int infocount = mcatenameInfo.Count == dictStoreCodeNameInfo.Count ? mcatenameInfo.Count : 0;

            for (int i = 0; i < infocount; i++)
            {
                List<string> regAssignArgs = new List<string>
                {
                    "'" + mcatenameInfo[i], "'" + dictStoreCodeNameInfo[i]
                };
                regAssginList.Add(regAssignArgs);
            }

            var regAssignItem = new HardIpRegAssign();
            blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
            regAssignItem.SubBlockName = subBlockName;
            regAssignItem.Type =
                function.FunctionName.Equals("hip_efuse_read", StringComparison.CurrentCultureIgnoreCase) ||
                function.FunctionName.Equals("HardIPFuseRead", StringComparison.CurrentCultureIgnoreCase) ?
                RegisterAssignType.HIP_eFuse_Read : RegisterAssignType.HIP_eFuse_Write;

            if (!HardIpInputData.HardIpRegAssigns.Exists(
                x => x.SubBlockName.Equals(regAssignItem.SubBlockName, StringComparison.OrdinalIgnoreCase) &&
                x.Type == regAssignItem.Type))
            {
                regAssignItem.RegAssignList = regAssginList;
                HardIpInputData.HardIpRegAssigns.Add(regAssignItem);
            }
        }

        private List<string> BuildPerBurstDigSrcRegisterAssign(HardIpPattern pattern, string blockName, string subBlockName, string type)
        {
            var burstRegList = new List<string>();
            foreach (HardIpPattern pat in pattern.BurstPatterns)
            {
                blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
                subBlockName = blockName + "_" + CommonGenerator.GetSubBlockName(pat.Pattern.GetLastPayload(),
                               pat.MiscInfo, blockName);
                var singleAssign = new HardIpRegAssign { SubBlockName = subBlockName };
                var singleRegAssginList = new List<List<string>>();

                switch (type.ToUpper())
                {
                    case DigsrcEquation:
                    case "DIGSRCEQUATION"://+
                        if (!string.IsNullOrEmpty(pat.DigSrcEquation))
                        {
                            singleAssign.Type = RegisterAssignType.DigSrc_Equation;
                            foreach (string item in pat.DigSrcEquation.Split('+').ToList())
                            {
                                singleRegAssginList.Add(new List<string> { item });
                            }
                        }
                        break;
                    case DigsrcAssignment:
                    case "DIGSRCASSIGNMENT":
                        if (!string.IsNullOrEmpty(pat.RegisterAssignment))
                        {
                            singleAssign.Type = RegisterAssignType.DigSrc_Assignment;
                            foreach (string item in pat.RegisterAssignment.Split(';').ToList())
                            {
                                List<string> data = new List<string>();
                                List<string> arr = item.Split('=').ToList();
                                data.Add(arr[0]);
                                if (arr.Count > 1)
                                {
                                    data.Add("'" + arr[1]);
                                }

                                singleRegAssginList.Add(data);
                            }
                        }
                        break;
                }

                if (!string.IsNullOrEmpty(pat.DigSrcEquation) || !string.IsNullOrEmpty(pat.RegisterAssignment))
                {
                    if (!HardIpInputData.HardIpRegAssigns.Exists(x => x.SubBlockName.Equals(singleAssign.SubBlockName, StringComparison.OrdinalIgnoreCase) && x.Type == singleAssign.Type))
                    {
                        singleAssign.RegAssignList = singleRegAssginList;
                        HardIpInputData.HardIpRegAssigns.Add(singleAssign);
                    }

                    burstRegList.Add($"Reg_assign:{singleAssign.SubBlockName}");
                }
                else
                {
                    burstRegList.Add("");
                }

            }
            return burstRegList;
        }

        private static List<List<string>> GetSimpleAssigns(HardIpRegAssign regAssignItem, RegisterAssignType type, IEnumerable<string> items, Func<string, List<string>> rowBuilder)
        {
            var regAssignList = new List<List<string>>();
            regAssignItem.Type = type;

            foreach (string item in items)
            {
                regAssignList.Add(rowBuilder(item));
            }

            return regAssignList;
        }

        private static List<List<string>> GetMeasureAssigns(HardIpRegAssign regAssignItem, RegisterAssignType type, string infoParameter, string[] splitSeq, char valueSeparator, HardIpPattern pattern, int registerIndex)
        {
            var regAssignList = new List<List<string>>();
            regAssignItem.Type = type;

            List<string> values = infoParameter.Split(valueSeparator).ToList();

            if (values.Count != splitSeq.Length)
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_RegAssignError_02,
                    pattern.SheetName,
                    pattern.RowNum,
                    registerIndex,
                    [$"{type}"]
                );
                return regAssignList;
            }

            for (int i = 0; i < values.Count; i++)
            {
                regAssignList.Add(new List<string>
                {
                    splitSeq[i],
                    "'" + values[i]
                });
            }

            return regAssignList;
        }

        private static List<List<string>> GetDigSrcAssigns(HardIpRegAssign regAssignItem, string infoParameter)
        {
            var regAssignList = new List<List<string>>();
            regAssignItem.Type = RegisterAssignType.DigSrc_Assignment;
            foreach (string item in infoParameter.Split(';').ToList())
            {
                List<string> data = new List<string>();
                List<string> arr = item.Split('=').ToList();
                data.Add(arr[0]);
                if (arr.Count > 1)
                {
                    data.Add("'" + arr[1]);
                }

                regAssignList.Add(data);
            }
            return regAssignList;
        }

        private static List<List<string>> GetCalcEquationAssigns(HardIpRegAssign regAssignItem, string infoParameter)
        {
            var regAssignList = new List<List<string>>();
            regAssignItem.Type = RegisterAssignType.Calc_Eqn;
            foreach (string item in infoParameter.Split(';').ToList())
            {
                if (item.StartsWith("Alg", StringComparison.OrdinalIgnoreCase))
                {
                    var algArgList = new List<string>
                    {
                        "Alg",
                        "",
                        item.Replace("Alg::","")
                    };
                    regAssignList.Add(algArgList);
                }
                else
                {
                    regAssignList.Add(item.Split(':').ToList());
                }
            }
            return regAssignList;
        }

        private static List<List<string>> GetRegAssignList(string type, HardIpRegAssign regAssignItem, string infoParameter, string[] splitSeq, HardIpPattern pattern, int registerIndex)
        {
            var regAssignList = new List<List<string>>();
            switch (type.ToUpper())
            {
                case TestSequence:
                case "MEASURESEQUENCE":
                    regAssignList = GetSimpleAssigns(regAssignItem, RegisterAssignType.TestSequence, splitSeq, item => new List<string> { "'" + item });
                    break;
                case DigsrcEquation:
                case "DIGSRCEQUATION"://+
                    regAssignList = GetSimpleAssigns(regAssignItem, RegisterAssignType.DigSrc_Equation, infoParameter.Split('+'), item => new List<string> { "'" + item });
                    break;
                case DigsrcAssignment:
                case "DIGSRCASSIGNMENT":
                    regAssignList = GetDigSrcAssigns(regAssignItem, infoParameter);
                    break;
                case CusStrDigcapdata:
                case "DIGCAPDATACUSTOMSTRING":
                    regAssignList = GetSimpleAssigns(regAssignItem, RegisterAssignType.CUS_Str_DigCapData, infoParameter.Split(','), item => item.Split(':').ToList());
                    break;
                case CalcEqn:
                case "CALCEQUATION":
                    regAssignList = GetCalcEquationAssigns(regAssignItem, infoParameter);
                    break;
                case MeasiRange:
                case "MEASURERANGE":
                    regAssignList = GetMeasureAssigns(regAssignItem, RegisterAssignType.MeasI_Range, infoParameter, splitSeq, '+', pattern, registerIndex);
                    break;
                case ForcevVal:
                case "MEASUREFORCEV":
                    regAssignList = GetMeasureAssigns(regAssignItem, RegisterAssignType.ForceV_Val, infoParameter, splitSeq, '|', pattern, registerIndex);
                    break;
                case ForceiVal:
                case "MEASUREFORCEI":
                    regAssignList = GetMeasureAssigns(regAssignItem, RegisterAssignType.ForceI_Val, infoParameter, splitSeq, '|', pattern, registerIndex);
                    break;
                case CusStrMainprogram:
                case "MAINPROGRAMCUSTOMSTRING":
                    regAssignList = GetSimpleAssigns(regAssignItem, RegisterAssignType.CUS_Str_MainProgram, infoParameter.Split(','), item => new List<string> { "'" + item });
                    break;
                case "TRIMDICTIONARYSTORENAME":
                    regAssignList = GetSimpleAssigns(regAssignItem, RegisterAssignType.TrimDictionaryStoreName, infoParameter.Split(','), item => new List<string> { "'" + item });
                    break;
                default:
                    regAssignItem.Type = RegisterAssignType.NoneType;
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_RegAssignError_03,
                        pattern.SheetName,
                        pattern.RowNum,
                        registerIndex,
                        [$"{type}"]
                    );
                    break;
            }
            return regAssignList;
        }

        private void HandleGenericRegisterAssignments(Function function, HardIpPattern pattern, string blockName, string subBlockName, List<string> parameters, int registerIndex)
        {
            for (int i = 0; i < parameters.Count; i++)
            {
                if (function.ArgList[i] == null)
                {
                    function.ArgList[i] = "";
                }

                string infoParameter = function.ArgList[i];
                string type = parameters[i];
                // For all args which start with "+"
                if (!string.IsNullOrEmpty(infoParameter) && infoParameter.StartsWith("+") && infoParameter.Trim().Length > 1)
                {
                    function.ArgList[i] = "'" + function.ArgList[i];
                    infoParameter = function.ArgList[i];
                }

                if (infoParameter.Length < 6000)
                {
                    continue;
                }

                if (pattern.Pattern.IsMultiTimeDomain() && !pattern.Pattern.IsFullMtdPattern())
                {
                    continue;
                }

                if (pattern.Pattern.IsMultiple() && (type.Equals(DigsrcEquation, StringComparison.OrdinalIgnoreCase) || type.Equals(DigsrcAssignment, StringComparison.OrdinalIgnoreCase)))
                {
                    List<string> burstRegList = BuildPerBurstDigSrcRegisterAssign(pattern, blockName, subBlockName, type);
                    function.ArgList[i] = string.Join("|", burstRegList);

                    //Hint in error Report
                    string error = $"The length of {type} content is over 6000, move to reg assign table";
                    ErrorReportManager.AddError(
                        HardIpErrorType.W_RegAssignWarning_01,
                        pattern.SheetName,
                        pattern.RowNum,
                        registerIndex,
                        [type]
                    );
                    continue;
                }

                var regAssignItem = new HardIpRegAssign();
                blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
                regAssignItem.SubBlockName = blockName + "_" +
                    CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, blockName);
                var regAssginList = new List<List<string>>();

                string[] splitSeq = function.GetParamValue("measureSequence").Split(',');

                regAssginList = GetRegAssignList(type, regAssignItem, infoParameter, splitSeq, pattern, registerIndex);

                //set args
                function.ArgList[i] = $"Reg_assign:{regAssignItem.SubBlockName}";
                if (!HardIpInputData.HardIpRegAssigns.Exists(x => x.SubBlockName.Equals(regAssignItem.SubBlockName, StringComparison.OrdinalIgnoreCase) && x.Type == regAssignItem.Type))
                {
                    regAssignItem.RegAssignList = regAssginList;
                    HardIpInputData.HardIpRegAssigns.Add(regAssignItem);
                }

                //Hint in error Report
                string errorMessage = $"The legnth of {type} content is over 6000, move to reg assign table";
                ErrorReportManager.AddError(
                    HardIpErrorType.W_RegAssignWarning_01,
                    pattern.SheetName,
                    pattern.RowNum,
                    registerIndex,
                    [type]
                );
            }
        }

        private static readonly HashSet<string> _functionNames = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase
        )
        {
            "hip_efuse_read",
            "hip_efuse_write",
            "HardIPFuseRead",
            "HardIPFuseWrite"
        };

        private static bool IsHipOrHardIpEfuse(Function function)
        {
            return _functionNames.Contains(function.FunctionName);
        }

        public void CheckArgsExceedLimitation(Function function, HardIpPattern pattern)
        {
            string blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
            string blockNameSuffix = CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, blockName);
            bool isNameEqual = pattern.SheetName.Equals(NeededSheets.ClockPllMeas.Replace("_", ""), StringComparison.OrdinalIgnoreCase);
            if (HardIpInputData.HardIpParaData == null || isNameEqual)
            {
                return;
            }

            if (!HardIpSheet.PlanHeaderIdx.TryGetValue("registerIndex", out int registerIndex))
            {
                return;
            }

            string subBlockName = blockName + "_" + blockNameSuffix;
            if (Regex.IsMatch(pattern.Pattern.GetLastPayload(), "^dd_", RegexOptions.IgnoreCase))
            {
                subBlockName += "_DD";
            }

            List<string> parameters = function.Parameters.Split(',').ToList();
            if (IsHipOrHardIpEfuse(function))
            {
                ProcessHipEfuseRegisterAssignments(function, pattern, blockName, subBlockName, parameters);
            }
            else
            {
                HandleGenericRegisterAssignments(function, pattern, blockName, subBlockName, parameters, registerIndex);
            }
        }


        public void CheckArgsExceedLimitationLcd(Function function, HardIpPattern pattern)
        {
            string blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
            string subBlockName = blockName + "_" + CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(),
                                               pattern.MiscInfo, blockName);
            if (Regex.IsMatch(pattern.Pattern.GetLastPayload(), "^dd_", RegexOptions.IgnoreCase))
            {
                subBlockName += "_DD";
            }

            List<string> parameters = function.Parameters.Split(',').ToList();
            for (int i = 0; i < parameters.Count; i++)
            {
                string infoParameter = function.ArgList[i];
                string type = parameters[i];
                if (infoParameter.Length >= 6000)
                {
                    ProcessLcdArgExceedLimitation(function, pattern, subBlockName, type, infoParameter, i);
                }
            }
        }

        private void ProcessLcdArgExceedLimitation(Function function, HardIpPattern pattern, string subBlockName, string type, string infoParameter, int argIndex)
        {
            HardIpRegAssign regAssignItem = HardIpInputData.HardIpRegAssigns.FirstOrDefault(
                x =>
                    x.SubBlockName.Equals(subBlockName, StringComparison.CurrentCultureIgnoreCase) &&
                    x.Type.ToString().Equals(type, StringComparison.CurrentCultureIgnoreCase));
            if (regAssignItem != null)
            {
                return;
            }

            regAssignItem = new HardIpRegAssign { SubBlockName = subBlockName };

            var regAssginList = new List<List<string>>();
            List<string> testSequenceArr = function.ArgList[1].Split(',').ToList();
            if (testSequenceArr.Exists(x => x.ToLower().StartsWith("reg_assign")))
            {
                var result = new List<string>();
                List<List<string>> regList = HardIpInputData.HardIpRegAssigns.FirstOrDefault(
           x =>
               x.SubBlockName.Equals(subBlockName, StringComparison.CurrentCultureIgnoreCase) &&
               x.Type == RegisterAssignType.TestSequence).RegAssignList;
                foreach (List<string> list in regList)
                {
                    result.Add(list[0]);
                }
                testSequenceArr = result.ToList();
            }

            RegisterAssignType regisAssignType = BuildLcdRegAssignList(type, infoParameter, testSequenceArr, regAssginList);

            regAssignItem.Type = regisAssignType;
            regAssignItem.RegAssignList = regAssginList;
            HardIpInputData.HardIpRegAssigns.Add(regAssignItem);
            function.ArgList[argIndex] = $"Reg_assign:{regAssignItem.SubBlockName}";

            //Hint in error Report
            int registerIndex = HardIpSheet.PlanHeaderIdx["registerIndex"];
            ErrorReportManager.AddError(
                HardIpErrorType.W_RegAssignWarning_01,
                pattern.SheetName,
                pattern.RowNum,
                registerIndex,
                [type]
            );
        }

        private static RegisterAssignType BuildLcdRegAssignList(string type, string infoParameter, List<string> testSequenceArr, List<List<string>> regAssginList)
        {
            RegisterAssignType regisAssignType = RegisterAssignType.DigSrc_Assignment;
            #region prepare
            switch (type.ToUpper())
            {
                case "TESTSEQUENCE":
                    regisAssignType = RegisterAssignType.TestSequence;
                    AddPairedLcdAssigns(testSequenceArr, infoParameter.Split(',').ToList(), regAssginList);
                    break;
                case "MEASSTORENAME":
                    regisAssignType = RegisterAssignType.Meas_StoreName;
                    AddPairedLcdAssigns(testSequenceArr, infoParameter.Split('|').ToList(), regAssginList);
                    break;
                case "CALCEQUNAME":
                    regisAssignType = RegisterAssignType.CalcEquName;
                    AddPairedLcdAssigns(testSequenceArr, infoParameter.Split('|').ToList(), regAssginList);
                    break;
                case "CALCSTORENAME":
                    regisAssignType = RegisterAssignType.CalcStoreName;
                    AddPairedLcdAssigns(testSequenceArr, infoParameter.Split('|').ToList(), regAssginList);
                    break;
                case "MEASNAME":
                    regisAssignType = RegisterAssignType.MeasName;
                    foreach (string item in infoParameter.Split('|').ToList())
                    {
                        regAssginList.Add(new List<string> { item });
                    }

                    break;
                case "MEASPINS":
                    regisAssignType = RegisterAssignType.MeasPins;
                    AddPairedLcdAssigns(testSequenceArr, infoParameter.Split('+').ToList(), regAssginList);
                    break;
                case "CUS_STR_DIGCAPDATA":
                    regisAssignType = RegisterAssignType.CUS_Str_DigCapData;
                    foreach (string item in infoParameter.Split(',').ToList())
                    {
                        regAssginList.Add(item.Split(':').ToList());
                    }

                    break;
                case "DIGSRC_EQUATION": //+
                    regisAssignType = RegisterAssignType.DigSrc_Equation;
                    foreach (string item in infoParameter.Split('+').ToList())
                    {
                        regAssginList.Add(new List<string> { item });
                    }

                    break;
                case "DIGSRC_ASSIGNMENT":
                    regisAssignType = RegisterAssignType.DigSrc_Assignment;
                    foreach (string item in infoParameter.Split(';').ToList())
                    {
                        List<string> data = new List<string>();
                        List<string> arr = item.Split('=').ToList();
                        data.Add(arr[0]);
                        if (arr.Count > 1)
                        {
                            data.Add("'" + arr[1]);
                        }

                        regAssginList.Add(data);
                    }
                    break;
                case "CALC_EQN":
                    regisAssignType = RegisterAssignType.Calc_Eqn;
                    foreach (string item in infoParameter.Split(';').ToList())
                    {
                        regAssginList.Add(item.Split(':').ToList());
                    }

                    break;
            }
            #endregion
            return regisAssignType;
        }

        private static void AddPairedLcdAssigns(List<string> testSequenceArr, List<string> valueList, List<List<string>> regAssginList)
        {
            if (testSequenceArr.Count != valueList.Count)
            {
                return;
            }

            for (int idx = 0; idx < valueList.Count; idx++)
            {
                List<string> data = new List<string> { testSequenceArr.ElementAt(idx) };
                data.AddRange(new List<string> { valueList[idx] });
                regAssginList.Add(data);
            }
        }
    }
}
