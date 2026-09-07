using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using LogLib.Utility;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class CalculationEquationChecker : HardIpPrecheckBase
    {
        private readonly Dictionary<string, string> _storeNamedir = new Dictionary<string, string>();

        public CalculationEquationChecker(HardIpSheet hardIpSheet, HardIpPattern pattern, Dictionary<string, string> planDic) : base(hardIpSheet, pattern)
        {
            _storeNamedir = planDic;
        }

        private void JudgeCalcEqnForDspUsage(HardIpPattern pattern, string sheetName, int rowNum, int colNum)
        {
            var dspUsageList = new List<string> { "Calc_ADC_Convert_Average", "Calc_GrayCode", "Calc_GrayCodeToBin", "2SCOMPLEMENT", "Calc_MIPI_Tolerance" };
            List<string> calcEqnList = pattern.CalcEqn.Split(';').ToList();
            var captureList = pattern.MeasPins.Where(p => !string.IsNullOrEmpty(p.CusStr)).Select(p => p.CusStr).ToList();
            var dspUsageGroups = calcEqnList.Where(p => !string.IsNullOrEmpty(p) && dspUsageList.
                Exists(q =>
                    q.Equals(Regex.Match(p.Split(':')[2], @"(?<name>\w+)\(").Groups["name"].Value, StringComparison.OrdinalIgnoreCase))
                ).GroupBy(p => Regex.Match(p.Split(':')[2], @"(?<name>\w+)\(").Groups["name"].Value).ToDictionary(p => p.Key, p => p.ToList());
            var calcLimitPins =
                pattern.MeasPins.Where(p => p.MeasType.Equals(MeasType.MeasLimit, StringComparison.OrdinalIgnoreCase) ||
                                            p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (KeyValuePair<string, List<string>> dspUsageGroup in dspUsageGroups)
            {
                string dspStr = "";
                if (dspUsageGroup.Key.Equals("Calc_ADC_Convert_Average", StringComparison.OrdinalIgnoreCase))
                {
                    bool isAdcMoveToDsp = false;
                    dspStr = ProcessAdcConvertAverage(dspUsageGroup.Value, ref isAdcMoveToDsp);
                    int limitCount = 0;
                    if (dspUsageGroup.Value.Count > 1)
                    {
                        foreach (string dspValue in dspUsageGroup.Value)
                        {
                            limitCount += dspValue.Split(',').Length - 4;
                        }
                    }
                    else
                    {
                        if (dspUsageGroup.Value.First().Split(',').Length == 5)
                        {
                            limitCount = 1;
                        }
                        else if (dspUsageGroup.Value.First().Split(',').Length == 6)
                        {
                            int firstOne = dspUsageGroup.Value.First().Split(',')[4].Split('&').Length;
                            int secondOne = dspUsageGroup.Value.First().Split(',')[5].Split('&').Length;
                            limitCount = firstOne * secondOne;
                        }
                    }

                    if (isAdcMoveToDsp)
                    {
                        if (
                            calcLimitPins.Count(p =>
                                p.MeasType.Equals(MeasType.MeasLimit, StringComparison.OrdinalIgnoreCase)
                            ) < limitCount
                        )
                        {
                            ErrorReportManager
                                .AddError(HardIpErrorType.E_ADC_Convertor_01, sheetName, rowNum, colNum, []);
                        }
                        else if (
                            calcLimitPins
                                .GetRange(0, limitCount)
                                .Count(p =>
                                    p.MeasType.Equals(MeasType.MeasLimit, StringComparison.OrdinalIgnoreCase)
                                ) < limitCount
                        )
                        {
                            ErrorReportManager
                                .AddError(HardIpErrorType.E_ADC_Convertor_02, sheetName, rowNum, colNum, [$"{limitCount}"]);
                        }
                    }
                    else
                    {
                        pattern.MeasPins.RemoveAll(calcLimitPins.Contains);
                        pattern.MeasPins.AddRange(calcLimitPins.Where(p => p.MeasType.Equals(MeasType.MeasCalc)).ToList());
                        pattern.MeasPins.AddRange(calcLimitPins.Where(p => p.MeasType.Equals(MeasType.MeasLimit)).ToList());
                    }
                }

                if (dspUsageGroup.Key.Equals("Calc_GrayCode", StringComparison.OrdinalIgnoreCase) ||
                    dspUsageGroup.Key.Equals("Calc_GrayCodeToBin", StringComparison.OrdinalIgnoreCase))
                {
                    dspStr = ProcessGrayCode(dspUsageGroup.Value, captureList);
                }

                if (dspUsageGroup.Key.Equals("2SCOMPLEMENT", StringComparison.OrdinalIgnoreCase))
                {
                    dspStr = Process2SComplement(dspUsageGroup.Value);
                }

                if (dspUsageGroup.Key.Equals("Calc_MIPI_Tolerance", StringComparison.OrdinalIgnoreCase))
                {
                    dspStr = "TOLERANCE";
                }

                string groupValues = string.Join(";", dspUsageGroup.Value);
                if (!string.IsNullOrEmpty(dspStr))
                {
                    calcEqnList.RemoveAll(dspUsageGroup.Value.Contains);
                    pattern.DspFunction.Add(dspStr);
                    ErrorReportManager.AddError(HardIpErrorType.E_CalcEqnCheck_01, sheetName, rowNum, colNum, [groupValues, dspStr]);
                }
                else
                {
                    ErrorReportManager.AddError(HardIpErrorType.E_CalcEqnCheck_02, sheetName, rowNum, colNum, [groupValues, dspStr]);
                }

            }
            pattern.CalcEqn = string.Join(";", calcEqnList);
            pattern.DspFunction.RemoveAll(string.IsNullOrEmpty);
        }

        internal string ProcessAdcConvertAverage(List<string> calcEqns, ref bool isAdcMoveToDsp)
        {
            string regCalcFunc = @"\((?<args>.*)\)";
            var adcLists = calcEqns.Select(p => Regex.Match(p, regCalcFunc, RegexOptions.IgnoreCase).Groups["args"].Value).ToList();
            #region process input if count == 1
            if (adcLists.Count <= 1 || adcLists.Any(p => p.Split(',').Length != adcLists[0].Split(',').Length))
            {
                if (adcLists.Count == 1)
                {
                    isAdcMoveToDsp = true;
                    List<string> adcSgmts = adcLists[0].Split(',').ToList();
                    var convertStr = new List<string>();
                    List<string> preSgmts = adcSgmts.Count >= 4 ? adcSgmts.GetRange(0, 4) : new List<string>();
                    convertStr.AddRange(preSgmts);
                    convertStr.Add(string.Join("&", adcSgmts.GetRange(preSgmts.Count, adcSgmts.Count - preSgmts.Count)));
                    return $"{"Calc_ADC_Convert_Average"}({string.Join(",", convertStr)})";
                }
                return "";
            }
            #endregion
            var majorSameList = new List<string>();
            var majorDiffList = new List<List<string>>();
            var minorSameList = new List<string>();
            var minorDiffList = new List<string>();
            for (int i = 0; i < adcLists[0].Split(',').Length; i++)
            {
                if (adcLists.All(p => p.Split(',')[i].Equals(adcLists[0].Split(',')[i], StringComparison.OrdinalIgnoreCase)))
                {
                    majorSameList.Add(adcLists[0].Split(',')[i]);
                }
                else
                {
                    majorDiffList.Add(adcLists.Select(p => p.Split(',')[i]).ToList());
                }
            }
            bool firstCheckFlag = true;
            var subMinorSame = new List<string>();
            foreach (List<string> majorDiffItem in majorDiffList)
            {
                minorSameList.Clear();
                int minorLoopCount = majorDiffItem.Min(p => p.Split('_').Length);
                for (int i = 1; i <= minorLoopCount; i++)
                {
                    string firstPattern = majorDiffItem[0];
                    if (majorDiffItem.All(p => p.Split('_')[p.Split('_').Length - i]
                        .Equals(firstPattern.Split('_')[firstPattern.Split('_').Length - i], StringComparison.OrdinalIgnoreCase)))
                    {
                        minorSameList.Add(firstPattern.Split('_')[firstPattern.Split('_').Length - i]);
                    }
                    else
                    {
                        break;
                    }
                }
                minorSameList.Reverse();
                string minorSamePattern = string.Join("_", minorSameList);
                if (firstCheckFlag)
                {
                    majorSameList.Add(minorSamePattern);
                    firstCheckFlag = false;
                }
                else
                {
                    subMinorSame.Add(minorSamePattern);
                }
                List<string> minorDiffItems = majorDiffItem;
                //MinorDiffItems.Select(p => p.Replace("_"+MinorSamePattern, ""));
                minorDiffList.AddRange(minorDiffItems.Select(p => p.Replace("_" + minorSamePattern, "")));
            }
            if (subMinorSame.Count > 0)
            {
                majorSameList.Add(string.Join("&", subMinorSame));
            }

            minorDiffList = minorDiffList.Distinct().ToList();
            majorSameList.Add(string.Join("&", minorDiffList));
            if (adcLists.All(p => p.Split(',').ToList().GetRange(0, 3).SequenceEqual(majorSameList.GetRange(0, 3))))
            {
                isAdcMoveToDsp = true;
                return $"{"Calc_ADC_Convert_Average"}({string.Join(",", majorSameList)})";
            }
            return "";
        }

        internal string ProcessGrayCode(List<string> calcEqns, List<string> captures)
        {
            //False to Singed ;True to Unsigned 
            var result = new List<string>();
            string regCalcFunc = @"\((?<args>.*)\)";
            foreach (string calcEqn in calcEqns)
            {
                List<string> argument = Regex.Match(calcEqn, regCalcFunc, RegexOptions.IgnoreCase).Groups["args"].Value.ToLower().Split(',').ToList();
                string dspFunctionName = argument[0].Equals("False", StringComparison.OrdinalIgnoreCase) ? "SignedGray" : "UnSignedGray";
                var argList = new List<string>();
                foreach (string reg in captures)
                {
                    if (argument.Contains(reg.ToLower()))
                    {
                        argList.Add(reg);
                    }
                }

                result.Add($"{dspFunctionName}:{string.Join(",", argList)}");
            }
            return string.Join(";", result);

        }

        private string Process2SComplement(List<string> calcEqns)
        {
            var result = new List<string>();
            string regCalcFunc = @"\((?<args>.*)\)";
            foreach (string calcEqn in calcEqns)
            {
                string argument = Regex.Match(calcEqn, regCalcFunc, RegexOptions.IgnoreCase).Groups["args"].Value;
                result.Add($"{"2sComplement"}:{argument}");
            }
            return string.Join(";", result);
        }

        private void CheckCalcFunction(HardIpPattern pattern)
        {
            if (string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                return;
            }
            int miscInfoIndex = HardIpSheet.PlanHeaderIdx["miscInfoIndex"];
            if (pattern.MiscInfoDict.TryGetValue("calc", out string value))
            {
                foreach (string calcFunction in value.Split(';'))
                {
                    Function target = TestProgram.VbtFunctionLib.GetFunctionByName(calcFunction, "hardip", true);
                    if (target.IsFound && target.IsInterPose && target.NameSpace.ContainsIgnoreCase("CoreTestLibrary.HardIP".ToLower()))
                    {
                        continue;
                    }
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_MissingCalcFunction_01,
                        pattern.SheetName,
                        pattern.RowNum,
                        miscInfoIndex,
                        [target.FunctionName]
                    );
                }
            }
            foreach (MeasPin measPin in pattern.MeasPins.Where(x => x.MeasType == MeasType.MeasLimit))
            {

                if (measPin.MiscInfo.ContainsIgnoreCase("calc"))
                {
                    string[] miscinfo = measPin.MiscInfo.ToLower().Split(';');
                    foreach (string calcInfo in miscinfo)
                    {
                        if (calcInfo.StartsWith("calc:"))
                        {
                            string calcFunction = calcInfo.Split(':')[1];
                            Function target = TestProgram.VbtFunctionLib.GetFunctionByName(calcFunction, "hardip", true);

                            if (target.IsFound && target.IsInterPose && target.NameSpace.ContainsIgnoreCase("CoreTestLibrary.HardIP".ToLower()))
                            {
                                continue;
                            }
                            ErrorReportManager.AddError(
                                HardIpErrorType.E_MissingCalcFunction_01,
                                pattern.SheetName,
                                measPin.RowNum,
                                miscInfoIndex,
                                [target.FunctionName]
                            );
                        }
                    }
                }
            }
        }

        public override void Check()
        {
            CheckCalcFunction(Pattern);
            var dicPinResult = new Dictionary<string, string>();
            try
            {
                string paraMissing = "";
                #region get Name Dictionary
                List<string> trimNames = SearchInfo.GetTrimStoreNameByMiscInfo(Pattern.MiscInfo);
                string digCapName = SearchInfo.GetDigCapNameByMiscInfo(Pattern.MiscInfo);
                BuildPinResultDic(dicPinResult, trimNames, digCapName);

                #endregion

                if (Pattern.CalcEqn == "")
                {
                    return;
                }

                foreach (string equation in Pattern.CalcEqn.Split(';'))
                {
                    ProcessEquation(equation, dicPinResult, ref paraMissing);
                }

                JudgeCalcEqnForDspUsage(Pattern, Pattern.SheetName, Pattern.RowNum,
                    HardIpSheet.PlanHeaderIdx["miscInfoIndex"]);

                if (paraMissing != "")
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_MisCalculationParaDefine_01,
                        Pattern.SheetName,
                        Pattern.RowNum,
                        0,
                        [paraMissing.Trim(',')]
                    );
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        private void BuildPinResultDic(Dictionary<string, string> dicPinResult, List<string> trimNames, string digCapName)
        {
            foreach (MeasPin pin in Pattern.MeasPins)
            {
                if (pin.CusStr != "")
                {
                    if (!dicPinResult.ContainsKey(pin.CusStr.ToUpper()))
                    {
                        dicPinResult.Add(pin.CusStr.ToUpper(), pin.PinName);
                    }
                    else
                    {
                        dicPinResult[pin.CusStr.ToUpper()] = pin.PinName;
                    }
                }
                if (trimNames.Count > 0)
                {
                    foreach (string trimName in trimNames)
                    {
                        if (!dicPinResult.ContainsKey(trimName.ToUpper()))
                        {
                            dicPinResult.Add(trimName.ToUpper(), pin.PinName);
                        }
                        else
                        {
                            dicPinResult[trimName.ToUpper()] = pin.PinName;
                        }
                    }

                }
                if (digCapName != "")
                {
                    if (!dicPinResult.ContainsKey(digCapName.ToUpper()))
                    {
                        dicPinResult.Add(digCapName.ToUpper(), pin.PinName);
                    }
                    else
                    {
                        dicPinResult[digCapName.ToUpper()] = pin.PinName;
                    }
                }
            }
        }

        private void ProcessEquation(string equation, Dictionary<string, string> dicPinResult, ref string paraMissing)
        {
            if (equation == "")
            {
                return;
            }

            string[] equaArr = equation.Split(':');
            string equaAlg = equaArr.Length >= 3 ? equaArr[2] : equaArr[equaArr.Length - 1];
            if (equaArr[0].Split('@').Last() != "Alg" &&
                Regex.IsMatch(equaAlg.Trim(), @"^[\w]+\s*[\(](?<parameter>([\w]|[,])+)[\)]",
                    RegexOptions.IgnoreCase))
            {
                string parameter =
                    Regex.Match(equaAlg.Trim(), @"^[\w]+\s*[\(](?<parameter>([\w]|[,])+)[\)]",
                        RegexOptions.IgnoreCase).Groups["parameter"].ToString();
                if (!string.IsNullOrEmpty(parameter))
                {
                    foreach (string para in parameter.Split(','))
                    {
                        bool pinExist = IsPinExist(para.ToUpper());
                        bool storeNameExist = IsStoreNameExist(para.ToUpper());
                        if (!pinExist && !storeNameExist)
                        {
                            paraMissing += "," + para;
                        }
                    }
                }
                return;
            }
            if (equaArr[0] == "C")
            {
                MatchCollection matches = Regex.Matches(equaAlg, @"\b([a-zA-Z]|[_])[\w]+\b");
                foreach (Match match in matches)
                {
                    string oriPinName = match.Value.ToUpper();
                    if (oriPinName.StartsWith("_") && oriPinName.EndsWith(DataConvertor.Var))
                    {
                        oriPinName = oriPinName.Substring(1, oriPinName.Length - 1 - DataConvertor.Var.Length);
                    }

                    bool pinExist = IsPinExist(oriPinName, match.Value.ToUpper());
                    bool storeNameExist = IsStoreNameExist(oriPinName, match.Value.ToUpper());
                    if (!pinExist && !storeNameExist)
                    {
                        {
                            paraMissing += "," + match.Value;
                        }
                    }
                }
            }
            else
            {
                //find all the label
                MatchCollection matches = Regex.Matches(equaAlg,
                    @"(?<pinName>[\w]+)[\(](?<label>([a-zA-Z]|[_])[\w]*)[\)]");
                foreach (Match match in matches)
                {
                    string label = match.Groups["label"].ToString();
                    if (!dicPinResult.ContainsKey(label.ToUpper()) && !IsPinExist(match.Value.ToUpper()))
                    {
                        paraMissing += "," + match.Value;
                    }
                }
            }
        }

        private static bool IsPinExist(params string[] pins)
        {
            return pins.Any(TestProgram.IgxlWorkBk.PinMapPair.Value.IsPinExist);
        }

        private bool IsStoreNameExist(params string[] pins)
        {
            return pins.Any(_storeNamedir.ContainsKey);
        }
    }
}
