using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

using TestPlanLib.HardIpDc.BaseData;
using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common
{
    public class CommonGenerator
    {
        #region Get Basic Information from Pattern and Voltage

        /// <summary>
        /// Get Timing AC base on what specified on forceCondition: "AC:pinName:Value"
        /// </summary>
        /// <returns></returns>
        public static string GetTimingAc(HardIpPattern pattern)
        {
            return pattern.GetTimingsByAc().Aggregate("", (current, timing) => current + timing.Name + "_" + timing.SuffixAcSpecName + "_").Trim('_');
        }

        public static string GetHardipSheetName(string sheetName)
        {
            sheetName = sheetName.Trim().Replace(" ", "_");
            return sheetName;
        }

        /// <summary>
        /// Get BlockName
        /// </summary>
        /// <returns></returns>
        public static string GetBlockNameFromSheetName(string sheetName)
        {
            List<string> arr = sheetName.Split('_').ToList();
            if (arr.Count > 1)
            {
                arr.RemoveAt(0);
            }

            return string.Join("", arr).Replace(" ", "").ToUpper();
        }

        public static string GetBlockName(HardIpPattern pattern)
        {
            foreach (MeasPin meas in pattern.MeasPins)
            {
                foreach (string assign in meas.RfInstrumentSetup.Split('$'))
                {
                    string[] asignArr = assign.Split('=');
                    if (asignArr.Length == 2 && asignArr[0].Equals(HardIpConstData.Block, StringComparison.OrdinalIgnoreCase))
                    {
                        return asignArr[1].Replace("_", "");
                    }
                }
            }
            return "";
        }

        public static List<string> SplitByDelimiter(string input, char delimiter, int chunkSize)
        {
            List<string> chunks = new List<string>();
            int startIndex = 0;

            while (startIndex < input.Length)
            {
                int endIndex = input.IndexOf(delimiter, startIndex);
                if (endIndex == -1)
                {
                    endIndex = input.Length;
                }
                int length = endIndex - startIndex;
                if (length > chunkSize)
                {
                    length = chunkSize;
                }
                string chunk = input.Substring(startIndex, length);
                chunks.Add(chunk);
                startIndex = endIndex + 1;
            }

            return chunks;
        }

        public static string GetSubBlockName(string patternName, string miscInfo, string blockName, bool isShmooInChar = false)
        {
            string subBlockName = "";
            List<string> miscInfoSplit = SplitByDelimiter(miscInfo, ';', 1000);
            foreach (string assign in miscInfoSplit)
            {
                string[] asignArr = SplitByDelimiter(assign, ':', 1000).ToArray();
                if (asignArr.Length == 2 && asignArr[0].Trim().Equals(HardIpConstData.SubBlockName, StringComparison.OrdinalIgnoreCase))
                {
                    subBlockName = asignArr[1].Replace("_", "");
                    break;
                }
            }

            if (string.IsNullOrEmpty(subBlockName))
            {
                subBlockName = GetSubBlockNameByPattern(patternName, blockName);
                if (subBlockName.Equals(blockName.Split('_').Last(), StringComparison.OrdinalIgnoreCase))
                {
                    subBlockName = "";
                }
            }

            if (string.IsNullOrEmpty(subBlockName))
            {
                Match matchInsInPatt = HardIpConstData.RegInsInPatt.Match(patternName);
                if (matchInsInPatt.Success)
                {
                    subBlockName = matchInsInPatt.Groups["InsName"].ToString();
                }
            }

            if (string.IsNullOrEmpty(subBlockName))
            {
                Match matchOpcodeInPatt = HardIpConstData.RegOpcodeInPatt.Match(patternName);
                if (matchOpcodeInPatt.Success)
                {
                    subBlockName = matchOpcodeInPatt.Groups["Parameter"].ToString();
                }
            }

            if (isShmooInChar)
            {
                subBlockName += "CZ";
            }

            return subBlockName;
        }

        public static string GetSubBlockNameByPattern(string patternName, string blockName, bool isCheckScghItem = true)
        {
            List<string> subBlocks = new List<string>();
            List<string> patternSeg = patternName.Split('_').ToList();
            int siDmIndex = patternSeg.FindLastIndex(p => p.Equals("SI", StringComparison.OrdinalIgnoreCase) ||
                                                          p.Equals("DM", StringComparison.OrdinalIgnoreCase));
            if (siDmIndex != -1 && siDmIndex != patternSeg.Count - 1)
            {
                List<string> subBlockSegs = patternSeg.GetRange(siDmIndex + 1, patternSeg.Count - siDmIndex - 1);
                foreach (string subBlockSeg in subBlockSegs)
                {
                    if (!subBlockSeg.Equals(blockName, StringComparison.CurrentCultureIgnoreCase) && isCheckScghItem)
                    {
                        subBlocks.Add(subBlockSeg);
                    }
                    else
                    {
                        subBlocks.Add(subBlockSeg);
                    }
                }
            }
            else if (LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                return patternName;
            }
            return string.Join("_", subBlocks);
        }

        public static string GetInitSubBlockName(List<string> inits)
        {
            var initSegs = new List<string>();
            if (LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                return "";
            }
            foreach (string initPat in inits)
            {
                if (initPat.Split('_').Count() < 10)
                {
                    continue;
                }

                if (!initPat.Split('_')[9].Equals("allfrv", StringComparison.OrdinalIgnoreCase) &&
                    !initPat.Split('_')[9].Equals("x", StringComparison.OrdinalIgnoreCase))
                {
                    initSegs.Add(initPat.Split('_')[9]);
                }
            }
            return string.Format("{0}", string.Join("_", initSegs.Distinct()));
        }

        public static string GetSubBlockNameWithoutMinus(string subBlockName)
        {
            return subBlockName.Replace("-", "");
        }

        /// <summary>
        /// Get HardIp Level base on "DC: XXX" specified in forceCondition column
        /// </summary>
        /// <returns></returns>
        public static HardIpCategoryDef GetHardIpDcSetting(string levelUsed)
        {
            if (!string.IsNullOrEmpty(levelUsed) && TestPlanStatic.HardIpDcSheet != null)
            {
                HardIpCategoryDef levelSetting = TestPlanStatic.HardIpDcSheet.Rows.FirstOrDefault(
                    a => a.CategoryName.Equals(levelUsed, StringComparison.OrdinalIgnoreCase));
                return levelSetting;
            }
            return null;
        }

        /// <summary>
        /// get VbtFunction 
        /// </summary>
        /// <returns></returns>
        public static Function GetVbtFunctionBase(string functionName)
        {
            return TestProgram.VbtFunctionLib.GetFunctionByName(functionName, "hardip");
        }

        public static List<Timing> GetTimingsNwire(HardIpPattern pattern)
        {
            List<Timing> timings = pattern.GetTimingsByAc();
            return timings.FindAll(
                    s =>
                        NwireSingleton.Instance()
                            .SettingInfo.NwirePins.Find(
                                a => a.CreatePinNameWithDiff().Equals(s.Name, StringComparison.InvariantCulture)) !=
                        null);
        }

        #endregion

        #region Reused Igxl sheet columns generator

        /// <summary>
        /// generate TestName
        /// </summary>
        /// <returns></returns>
        public static string GenHardIpInsTestName(string blockName, string subBlockName, string patternName, string prefixPatIndexFlag,
            string timingAc, string prefixForceVFlag, string instNameSubStr, string labelVoltage, bool noPattern, bool isPostBurn, bool isgenbyflow, bool isDoMeas)
        {
            #region pattern by "Instance:"

            Match matchInsInPatt = HardIpConstData.RegInsInPatt.Match(patternName);
            if (matchInsInPatt.Success)
            {
                string insName = matchInsInPatt.Groups["InsName"].ToString();
                if (!insName.StartsWith(blockName + "_", StringComparison.CurrentCultureIgnoreCase))
                {
                    insName = blockName + "_" + insName;
                }

                if (isgenbyflow)
                {
                    return insName + prefixPatIndexFlag;
                }
                patternName = "INSREMOV_" + insName;
                return patternName.ToUpper() + prefixPatIndexFlag;
            }

            #endregion

            #region pattern by "Opcode:"

            Match matchOpcodeInPatt = HardIpConstData.RegOpcodeInPatt.Match(patternName);
            if (matchOpcodeInPatt.Success)
            {
                return matchOpcodeInPatt.Groups["Parameter"].ToString();
            }

            #endregion

            subBlockName = subBlockName.Replace("-", "");
            string prefixSubBlock = string.IsNullOrEmpty(subBlockName) ? string.Empty : "_" + subBlockName;
            string prefixInstNameSubStr = string.IsNullOrEmpty(instNameSubStr) ? string.Empty : "_" + instNameSubStr;
            string prefixLabelVoltage = string.IsNullOrEmpty(labelVoltage) ? "" : "_" + labelVoltage;
            string prefixPatternName = "_" + patternName;
            string prefixPostBurn = "";
            if (noPattern && !string.IsNullOrEmpty(instNameSubStr))
            {
                prefixPatternName = prefixInstNameSubStr;
                prefixInstNameSubStr = "";
            }

            string testName = blockName + prefixSubBlock + prefixPatternName + prefixPatIndexFlag +
                              prefixForceVFlag + prefixInstNameSubStr + prefixPostBurn + prefixLabelVoltage;
            return testName.ToUpper();
        }
        public static string GenRtosInsTestName(string functionName, string blockName, string subBlockName, string patternName)
        {
            const string continueProcess = "__CONTINUE__";

            if (functionName.Equals(VbtFunctionLibShared.RtosBootUp, StringComparison.OrdinalIgnoreCase))
            {
                return blockName + "_" + subBlockName;
            }

            if (functionName.Equals(FuncNameConst.CSharpFuncNameRtosRunScenario, StringComparison.OrdinalIgnoreCase))
            {
                var names = new List<string> { "Rtos", subBlockName, patternName };
                string result = string.Join("_",
                    names.Where(x => !string.IsNullOrWhiteSpace(x)));
                return result;
            }
            return continueProcess;

        }
        public static string GenHardIpFlowFailAction(string sheetName, string blockName, string subBlockName, string patternName, string timingAc, string instNameSubStr, string labelVoltage, string miscInfo, bool noPattern)
        {
            string failAction = GenHardIpFlowFailFlag(sheetName, blockName, subBlockName, patternName, timingAc, instNameSubStr, labelVoltage, noPattern);
            if (Regex.IsMatch(miscInfo, HardIpConstData.NoBin, RegexOptions.IgnoreCase))
            {
                failAction = SearchInfo.GetFlagNoBinStr(miscInfo, labelVoltage) + failAction;
            }

            return failAction;
        }

        public static string GenHipEfuseReadTestFailAction()
        {
            return HardIpConstData.PrefixHardIpFailAction + "_HARDIP_" + HardIpConstData.SubfixHipEfuseReadBinTableName + HardIpConstData.SuffixHardIpFailAction;
        }

        public static string GenHardIpEfuseReadBinParameter(string sheetName)
        {
            string prefixSheetName = sheetName.Split('_')[0];
            if (prefixSheetName.ContainsIgnoreCase("hardip"))
            {
                return HardIpConstData.BinFlowFlag + "_" + "HIP" + "_" + HardIpConstData.SubfixHipEfuseReadBinTableName;
            }

            return HardIpConstData.BinFlowFlag + "_" + prefixSheetName + "_" + HardIpConstData.SubfixHipEfuseReadBinTableName;
        }

        public static string GenHardIpFlowBinParameter(string sheetName, string blockName, string subBlockName)
        {
            var binItemsList = new List<string>
            {
                HardIpConstData.BinFlowFlag,
            };
            if (sheetName.ContainsIgnoreCase("hardip_"))
            {
                binItemsList.Add("HIP");
            }

            binItemsList.Add(blockName);
            binItemsList.Add(subBlockName.Replace("-", "_"));
            return string.Join("_", binItemsList);
        }

        public static string GenEnableWord(string patternName, string miscInfo, string labelVoltage, HardIpInputData hardIpInputData)
        {
            const string czPatEnableW = "_CZ";
            const string mnPatEnableW = "_MN";
            const string prefixEnable = "HardIP_";

            if (string.IsNullOrEmpty(labelVoltage))
            {
                return "";
            }

            bool isCzPattern = Regex.IsMatch(patternName, HardIpConstData.RegCzPattern, RegexOptions.IgnoreCase);
            string enableWord = prefixEnable + labelVoltage;
            if (isCzPattern)
            {
                enableWord += czPatEnableW;
            }

            if (Regex.IsMatch(patternName, "^mn_", RegexOptions.IgnoreCase))
            {
                enableWord += mnPatEnableW;
            }

            if (NeedRemoveEnableWord(miscInfo, labelVoltage))
            {
                enableWord = "";
            }

            switch (labelVoltage)
            {
                case HardIpConstData.LabelNv:
                    if (!hardIpInputData.HardIpParaData.NvEnable && !isCzPattern)
                    {
                        enableWord = "";
                    }

                    if (!hardIpInputData.HardIpParaData.CzNvEnable && isCzPattern)
                    {
                        enableWord = "";
                    }

                    break;
                case HardIpConstData.LabelLv:
                    if (!hardIpInputData.HardIpParaData.LvEnable && !isCzPattern)
                    {
                        enableWord = "";
                    }

                    if (!hardIpInputData.HardIpParaData.CzLvEnable && isCzPattern)
                    {
                        enableWord = "";
                    }

                    break;
                case HardIpConstData.LabelHv:
                    if (!hardIpInputData.HardIpParaData.HvEnable && !isCzPattern)
                    {
                        enableWord = "";
                    }

                    if (!hardIpInputData.HardIpParaData.CzHvEnable && isCzPattern)
                    {
                        enableWord = "";
                    }

                    break;
            }

            return enableWord;
        }

        private static bool NeedRemoveEnableWord(string miscInfo, string labelVoltage)
        {
            switch (labelVoltage)
            {
                case HardIpConstData.LabelNv:
                    if (Regex.IsMatch(miscInfo, HardIpConstData.RemoveNv, RegexOptions.IgnoreCase))
                    {
                        return true;
                    }

                    break;
                case HardIpConstData.LabelHv:
                    if (Regex.IsMatch(miscInfo, HardIpConstData.RemoveHv, RegexOptions.IgnoreCase))
                    {
                        return true;
                    }

                    break;
                case HardIpConstData.LabelLv:
                    if (Regex.IsMatch(miscInfo, HardIpConstData.RemoveLv, RegexOptions.IgnoreCase))
                    {
                        return true;
                    }

                    break;
            }
            return false;
        }

        public static string GenHardIpFlowFailFlag(string sheetName, string blockName, string subBlockName, string patternName, string timingAc, string instNameSubStr, string labelVoltage, bool noPattern)
        {
            string prefixSheetNameWithoutUndercroe = "_" + sheetName.Split('_')[0];
            string prefixBlockName = "_" + blockName;
            string prefixSubBlock = string.IsNullOrEmpty(subBlockName) ? string.Empty : "_" + subBlockName;
            string prefixLabelVoltage = string.IsNullOrEmpty(labelVoltage) ? "" : "_" + labelVoltage[0];

            Match matchInsInPatt = HardIpConstData.RegInsInPatt.Match(patternName);
            if (matchInsInPatt.Success)
            {
                patternName = matchInsInPatt.Groups["InsName"].ToString();
            }

            string prefixPatternName = "_" + patternName;
            string prefixInstNameSubStr = string.IsNullOrEmpty(instNameSubStr) ? string.Empty : "_" + instNameSubStr;//20180824 add
            if (noPattern && !string.IsNullOrEmpty(instNameSubStr))
            {
                prefixPatternName = prefixInstNameSubStr;
                prefixInstNameSubStr = "";
            }
            string prefixTimingAc = ""; //Aaron agrees no need ac condition as the part of instance name, 20250227 Hid C#.
            string failFlag = HardIpConstData.PrefixHardIpFailAction + prefixSheetNameWithoutUndercroe + prefixBlockName + GetSubBlockNameWithoutMinus(prefixSubBlock)
                              + prefixPatternName + prefixTimingAc + prefixInstNameSubStr + prefixLabelVoltage + HardIpConstData.SuffixHardIpFailAction;
            return failFlag;
        }

        public static string GenWirelessFlowFailFlag(string blockName, string subBlockName, string labelVoltage)
        {
            string prefixBlockName = "_" + blockName;
            string prefixSubBlock = string.IsNullOrEmpty(subBlockName) ? "" : "_" + subBlockName;
            string prefixLabelVoltage = string.IsNullOrEmpty(labelVoltage) ? "" : "_" + labelVoltage[0];
            return HardIpConstData.PrefixHardIpFailAction + prefixBlockName + GetSubBlockNameWithoutMinus(prefixSubBlock)
                + prefixLabelVoltage + HardIpConstData.SuffixHardIpFailAction;
        }

        public static string GenHardIpBlockFailFlag(string sheetName, string blockName)
        {
            return HardIpConstData.PrefixHardIpFailAction + "_" + sheetName.Split('_')[0] + "_" + blockName + HardIpConstData.SuffixHardIpFailAction;
        }

        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="opcodeList"></param>
        /// <param name="flowRowsBefore"></param>
        /// <param name="voltage"></param>
        /// <returns></returns>
        public static void ConvertPatNameInOpcode(List<string> opcodeList, List<FlowRow> flowRowsBefore, string voltage)
        {
            if (opcodeList == null || opcodeList.Count == 0)
            {
                return;
            }

            string subfixfailAction = voltage[0] + HardIpConstData.SuffixHardIpFailAction;
            for (int i = 0; i < opcodeList.Count; i++)
            {
                string[] opcode = opcodeList[i].Split(':');
                if (opcode.Length != 2)
                {
                    continue;
                }

                Regex reg = new Regex(@"[\w]+");
                string newOpcode = reg.Replace(opcode[1], delegate (Match m)
                {
                    string patternName = m.Value;
                    if (SearchInfo.IsValidPatName(patternName))
                    {
                        var failFlaglst = flowRowsBefore.Where(s =>
                            Regex.IsMatch(s.Parameter, patternName, RegexOptions.IgnoreCase) &&
                            s.Opcode != OpCode.UseLimit &&
                            s.FailAction.EndsWith(subfixfailAction)).Select(s => s.FailAction).ToList();
                        if (failFlaglst.Count > 0)
                        {
                            return string.Join("||", failFlaglst);
                        }
                    }
                    return patternName;
                });
                opcode[1] = newOpcode;
                opcodeList[i] = string.Join(":", opcode.ToList());
            }
        }

        public static string GetRegAssignName(HardIpPattern pattern)
        {
            string blockName = GetBlockNameFromSheetName(pattern.SheetName);
            string name = blockName + "_" + GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, blockName);
            if (Regex.IsMatch(pattern.Pattern.GetLastPayload(), "^dd_", RegexOptions.IgnoreCase))
            {
                name += "_DD";
            }

            return name;
        }

    }
}
