using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess
{
    internal class MeasPinResolver
    {
        //eg: "name1"
        private const string RegTName = @"(([""])(?<testName>.+)([""]))?";
        //eg: "StoreName"
        private const string RegStoreName = @"(\((?<cusStr>.+)\))?";
        //eg : MeasV Pin = PAD_MTR_ANALOG_TEST_N "Test 0" get pinName ->PAD_MTR_ANALOG_TEST_N
        private const string RegPinName = @"(\(*(?<pinName>[\w,:\s]+)\)*)";
        private const string RegBitNum = @"(?<capBit>\d+)";
        private const string RegPineNameExpression = @"^(?<pinName>[^)]+)[\s]*(([""])(?<testName>([^""])+)([""]))";
        //eg: Calc "Add" sn1+sn2
        private const string RegCalcExpression = @"^Calc[\s]+(([""])(?<testName>.+)([""]))?[\s&,]*(?<expression>.*)$";
        //eg: Limits "Cal_A" , (out1):4 or Limits IO1(out2)
        private const string RegLimitsExpression = @"^Limits[\s]*(([""])(?<testName>.+)([""]))?[\w|\s|,]*([\(](?<cusStr>[^)]+)[\)])?[\s]*([\:](?<capBit>[\d]+))?[\s]*";
        //eg: MeasI pin = pin1,pin2
        private const string RegMeasExpression = @"(?<MeasType>(Wi)*[(Meas)|(Src)]\S+)[\s]*(pin)?[\s]*=[\s]*(?<pin>(.*))";
        //eg: calc (LPDP_TX1P(snV2)/LPDP_TX1P(sn2))(sn4)
        private const string RegCalcCusreg = @"[\)]\s*[\(](?<cusStr>[^)]+)[\)]";
        //TBD,NA,N/A
        private const string RegNoLimit = @"^(TBD|NA|N[\/]A)$";

        private readonly TestPlanSheet _planSheet;
        private readonly ForceConditionResolver _forceConditionResolver;

        internal MeasPinResolver(TestPlanSheet planSheet, ForceConditionResolver forceConditionResolver)
        {
            _planSheet = planSheet;
            _forceConditionResolver = forceConditionResolver;
        }

        internal List<MeasPin> GetMeasPins(PatternRow patternRow, HardIpPattern pattern)
        {
            List<MeasPin> measPins = GetPreMeasPins(patternRow, pattern);

            measPins = GetDivideMeasPins(patternRow, measPins, pattern);

            measPins = GetPostMeasPins(measPins);

            return measPins;
        }

        #region preMeasPins
        private List<MeasPin> GetPreMeasPins(PatternRow patternRow, HardIpPattern pattern)
        {
            var measPins = new List<MeasPin>();
            foreach (PatChildRow measRow in patternRow.PatChildRows)
            {
                foreach (TestPlanRow tpMeasRow in ((PatSubChildRow)measRow).TpRows)
                {
                    measPins.Add(GetMeasPinInfo(tpMeasRow, pattern));
                }
            }
            return measPins;
        }

        private MeasPin GetMeasPinInfo(TestPlanRow tpMeasRow, HardIpPattern pattern)
        {
            var measPin = new MeasPin
            {
                MiscInfo = tpMeasRow.MiscInfo,
                RowNum = tpMeasRow.RowNum
            };

            string measStr = tpMeasRow.Meas.Replace("\n", "").Replace("\t", "").Trim();
            measPin.MeasType = GetMeasType(measStr, tpMeasRow.RowNum, _planSheet.PlanHeaderIdx["measIndex"]);
            if (measPin.MeasType == MeasType.MeasVdiff2)
            {
                measStr = Regex.Match(measStr, RegMeasExpression, RegexOptions.IgnoreCase).Groups["pin"].ToString().Trim(',').Trim().Replace(",", "::");
            }
            else if (measPin.MeasType != MeasType.MeasCalc && measPin.MeasType != MeasType.MeasLimit)
            {
                measStr = Regex.Match(measStr, RegMeasExpression, RegexOptions.IgnoreCase).Groups["pin"].ToString().Trim(',').Trim();
            }

            measPin.PinName = GetMeasPinName(measStr, measPin.MeasType);
            measPin.RfInstrumentSetup = tpMeasRow.RfIntrumentSetup;
            measPin.CusStr = GetCusStr(measStr, measPin.MeasType).Replace(" ", "");
            if (measPin.MeasType == MeasType.MeasCalc && measPin.CusStr != "")
            {
                measStr = measStr.Replace("(" + measPin.CusStr + ")", "");  // if Calc_Eqn contain dictionary info, move to CusStr for further assignment
            }

            measPin.CapBit = GetCapBit(measStr, measPin.MeasType);
            measPin.TestName = tpMeasRow.TestName != "" ? tpMeasRow.TestName : GetTestName(measStr, measPin.MeasType);
            measPin.CalcEqn = GetCalcEqnForPin(measStr, tpMeasRow.MiscInfo, measPin.MeasType).Replace(" ", "");
            measPin.RepeatCount = SearchInfo.GetRepeatLimitCount(tpMeasRow.MiscInfo);
            measPin.SkipUnit = GetSkipUnit(measStr, measPin.MeasType);

            //Convert limits to Hv,Lv,Nv limit
            List<List<MeasLimit>> limits = GetLimits(tpMeasRow);
            measPin.MeasLimitsH = limits[0];
            measPin.MeasLimitsL = limits[1];
            measPin.MeasLimitsN = limits[2];

            #region Get measure pin's force condition and sequence index
            SetMeasPinForceConditionAndSequence(measPin, tpMeasRow, pattern);
            #endregion

            //FW use
            measPin.InterPoseFunc = tpMeasRow.InterposeFunc;
            measPin.RfInterPose = tpMeasRow.RfInterpose;
            measPin.MeasWaitTime = GetMeasWaitTime(tpMeasRow);
            measPin.MeasRange = GetMeasIRange(tpMeasRow);
            measPin.RfInstrumentSetup = GetRfIntrumentSetup(tpMeasRow.RfIntrumentSetup);
            if (string.IsNullOrEmpty(measPin.RfInstrumentSetup))
            {
                if (tpMeasRow.MergeRowNumForMeas != 0 && tpMeasRow.MergeRowNumForMeas != tpMeasRow.RowNum)
                {
                    measPin.RowNumForMergeMeas = tpMeasRow.MergeRowNumForMeas;
                }
            }
            measPin.PatternName = pattern.Pattern.RealPatternName;

            return measPin;
        }

        private void SetMeasPinForceConditionAndSequence(MeasPin measPin, TestPlanRow tpMeasRow, HardIpPattern pattern)
        {
            if (pattern.TestPlanSequencesRf.Count > 0)
            {
                TestPlanSequence sequence = pattern.TestPlanSequencesRf.Find(s => s.StartRow <= tpMeasRow.RowNum && s.EndRow >= tpMeasRow.RowNum);
                if (sequence != null && measPin.MeasType != MeasType.MeasC && measPin.MeasType != MeasType.MeasLimit && measPin.MeasType != MeasType.MeasCalc)
                {
                    measPin.ForceConditions = _forceConditionResolver.DivideForceCondition(string.Join(";", sequence.ForceCondition), tpMeasRow.RowNum);
                    //default sequenceIndex=0, if test plan merged, will set the actual sequence index
                    measPin.SequenceIndex = sequence.SeqIndex;
                }
                else if (LocalSpecs.Options.Device == EnumDevice.LCD && measPin.MeasType == MeasType.MeasCalc)
                {
                    measPin.ForceConditions = _forceConditionResolver.DivideForceCondition(string.Join(";", sequence.ForceCondition), tpMeasRow.RowNum);
                    //default sequenceIndex=0, if test plan merged, will set the actual sequence index
                    measPin.SequenceIndex = sequence.SeqIndex;
                }
            }
            else if (pattern.TestPlanSequences.Count > 0)
            {
                TestPlanSequence sequence = pattern.TestPlanSequences.Find(s => s.StartRow <= tpMeasRow.RowNum && s.EndRow >= tpMeasRow.RowNum);
                if (sequence != null && measPin.MeasType != MeasType.MeasC && measPin.MeasType != MeasType.MeasLimit && measPin.MeasType != MeasType.MeasCalc)
                {
                    measPin.ForceConditions = _forceConditionResolver.DivideForceCondition(string.Join(";", sequence.ForceCondition), tpMeasRow.RowNum);
                    //default sequenceIndex=0, if test plan merged, will set the actual sequence index
                    measPin.SequenceIndex = sequence.SeqIndex;
                }
                else if (sequence != null && LocalSpecs.Options.Device == EnumDevice.RF && measPin.MeasType != MeasType.MeasCalc)
                {
                    measPin.ForceConditions = _forceConditionResolver.DivideForceCondition(string.Join(";", sequence.ForceCondition), tpMeasRow.RowNum);
                    measPin.SequenceIndex = sequence.SeqIndex;
                }
                else if (LocalSpecs.Options.Device == EnumDevice.LCD && measPin.MeasType == MeasType.MeasCalc)
                {
                    measPin.ForceConditions = _forceConditionResolver.DivideForceCondition(string.Join(";", sequence.ForceCondition), tpMeasRow.RowNum);
                    //default sequenceIndex=0, if test plan merged, will set the actual sequence index
                    measPin.SequenceIndex = sequence.SeqIndex;
                }
            }
            else
            {
                measPin.ForceConditions = _forceConditionResolver.DivideForceCondition(tpMeasRow.ForceCondition, tpMeasRow.RowNum);
            }
        }

        private List<List<MeasLimit>> GetLimits(TestPlanRow tpMeasRow)
        {
            var limits = new List<List<MeasLimit>>();
            var limitsH = new List<MeasLimit>();
            var limitsL = new List<MeasLimit>();
            var limitsN = new List<MeasLimit>();
            limits.Add(limitsH);
            limits.Add(limitsL);
            limits.Add(limitsN);

            foreach (MeasLimit measLimit in tpMeasRow.Limits)
            {
                var limitH = new MeasLimit(measLimit.JobName);
                var limitL = new MeasLimit(measLimit.JobName);
                var limitN = new MeasLimit(measLimit.JobName);
                limitsH.Add(limitH);
                limitsL.Add(limitL);
                limitsN.Add(limitN);

                //limit valus contains HV,LV,NV
                //user may assign unit "µ", equals to "u"
                string hi = measLimit.HiLimit.Replace("µ", "u");
                string lo = measLimit.LoLimit.Replace("µ", "u");
                //if assign TBD, NA or N/A, means no limit
                if (Regex.IsMatch(hi.Trim(), RegNoLimit, RegexOptions.IgnoreCase))
                {
                    hi = "";
                }

                if (Regex.IsMatch(lo.Trim(), RegNoLimit, RegexOptions.IgnoreCase))
                {
                    lo = "";
                }

                // Set HeaderIndex
                limitH.LoHeaderIndex = measLimit.LoHeaderIndex;
                limitH.HiHeaderIndex = measLimit.HiHeaderIndex;
                limitL.LoHeaderIndex = measLimit.LoHeaderIndex;
                limitL.HiHeaderIndex = measLimit.HiHeaderIndex;
                limitN.LoHeaderIndex = measLimit.LoHeaderIndex;
                limitN.HiHeaderIndex = measLimit.HiHeaderIndex;

                //Set hi limit for Hv,Lv,Nv
                string[] hiArr = hi.Split(',');
                if (hiArr.Length == 3)
                {
                    limitH.HiLimit = hiArr[0];
                    limitL.HiLimit = hiArr[1];
                    limitN.HiLimit = hiArr[2];
                }
                if (hiArr.Length == 2)
                {
                    limitH.HiLimit = hiArr[0];
                    limitL.HiLimit = hiArr[1];
                    limitN.HiLimit = "";
                }
                if (hiArr.Length == 1)
                {
                    limitH.HiLimit = hiArr[0];
                    limitL.HiLimit = hiArr[0];
                    limitN.HiLimit = hiArr[0];
                }
                //Set lo limit for Hv,Lv,Nv
                string[] loArr = lo.Split(',');
                if (loArr.Length == 3)
                {
                    limitH.LoLimit = loArr[0];
                    limitL.LoLimit = loArr[1];
                    limitN.LoLimit = loArr[2];
                }
                if (loArr.Length == 2)
                {
                    limitH.LoLimit = loArr[0];
                    limitL.LoLimit = loArr[1];
                    limitN.LoLimit = "";
                }
                if (loArr.Length == 1)
                {
                    limitH.LoLimit = loArr[0];
                    limitL.LoLimit = loArr[0];
                    limitN.LoLimit = loArr[0];
                }
                //Flag error if has valid format
                if (hiArr.Length > 3 || loArr.Length > 3)
                {
                    ErrorReportManager.AddError(HardIpErrorType.E_WrongLimitValue_01, _planSheet.SheetName, tpMeasRow.RowNum, 0, []);
                }
            }

            return limits;
        }

        private string GetRfIntrumentSetup(string rfIntrumentSetup)
        {
            var dic = new Dictionary<string, string>();
            foreach (string item in rfIntrumentSetup.Split(';'))
            {
                if (item.Contains('='))
                {
                    string name = item.Split('=')[0];
                    string value = item.Split('=')[1];
                    if (!dic.ContainsKey(name))
                    {
                        dic.Add(name, value);
                    }
                }
            }
            var list = new List<string>();
            foreach (KeyValuePair<string, string> item in dic)
            {
                list.Add(item.Key + "=" + item.Value);
            }

            return string.Join(";", list);
        }

        internal string GetMeasWaitTime(TestPlanRow tpRow)
        {
            foreach (string misc in tpRow.MiscInfo.Split(';'))
            {
                string[] miscitems = misc.Split(':');
                if (miscitems[0].Equals("MeasWait", StringComparison.OrdinalIgnoreCase) && miscitems.Length > 1)
                {
                    return miscitems[1];
                }
            }

            if (Regex.IsMatch(tpRow.Meas, MeasType.MeasWait, RegexOptions.IgnoreCase))
            {
                string regWait = @"MeasWait\s*(?<time>.*)";
                string time = Regex.Match(tpRow.Meas, regWait, RegexOptions.IgnoreCase).Groups["time"].ToString();
                return time;
            }
            return "";
        }

        internal string GetMeasIRange(TestPlanRow tpRow)
        {
            foreach (string misc in tpRow.MiscInfo.Split(';'))
            {
                string[] miscitems = misc.Split(':');
                if (miscitems[0].Equals("MeasRange", StringComparison.OrdinalIgnoreCase) && miscitems.Length > 1)
                {
                    return miscitems[1];
                }
            }

            return "";
        }

        private string GetCalcEqnForPin(string pinValue, string pinMiscInfo, string measType)
        {
            var calEqnList = new List<string>();

            //eg: Calc "Add", sn1+sn2, gets "C:Add:sn1+sn2"..update example by raze with comma 20170630
            //eg: Calc "test1", VDD_A(sn1)+VDD_B(sn2), and measure type is I, gets "I:test1:VDD_A(sn1)+VDD_B(sn2)" ..update example by raze with comma 20170630
            //eg: Calc "test1", sn1+sn2 (newStoreName), and measure type is I, gets "I:test1:VDD_A(sn1)+VDD_B(sn2):newStoreName"
            if (measType == MeasType.MeasCalc)
            {
                string calcExp = Regex.Match(pinValue, RegCalcExpression, RegexOptions.IgnoreCase).Groups["expression"].ToString().Trim();
                string testName = Regex.Match(pinValue, RegCalcExpression, RegexOptions.IgnoreCase).Groups["testName"].ToString().Trim();
                if (testName.Contains(":"))
                {
                    testName = testName.Substring(0, testName.IndexOf(":", StringComparison.Ordinal));
                }
                calEqnList.Add(testName + ":" + calcExp);
            }
            //Eg: "Calc:Algorithm_A;CalcParameter:rd0,rd1" in miscInfo, gets "Alg::Algorithm_A(rd0,rd1)"
            if (!string.IsNullOrEmpty(SearchInfo.GetCalculation(pinMiscInfo)))
            {
                calEqnList.Add(SearchInfo.GetCalculation(pinMiscInfo));
            }

            return string.Join(";", calEqnList);
        }

        private string GetTestName(string pinValue, string measType)
        {
            //eg: Limits "test1", test name is "test1"
            if (measType == MeasType.MeasLimit)
            {
                return Regex.Match(pinValue.Replace(",", " "), RegLimitsExpression, RegexOptions.IgnoreCase).Groups["testName"].ToString().Trim();
            }

            //eg: Calc "add" sn1+sn2, test name is "add"
            if (measType == MeasType.MeasCalc)
            {
                string testName = Regex.Match(pinValue.Replace(",", " "), RegCalcExpression, RegexOptions.IgnoreCase).Groups["testName"].ToString().Trim();
                if (testName.Contains(":"))
                {
                    testName = testName.Substring(0, testName.IndexOf(":", StringComparison.Ordinal));
                }

                return testName;
            }

            //Others => eg: MeasC pin = JTAG(data):4 "testName" or  Limits "Cal_A" , (out1):4 or Limits IO1(out2)

            return Regex.Match(pinValue, @"(([""])(?<testName>.+)([""]))", RegexOptions.IgnoreCase).Groups["testName"].ToString().Trim();
        }

        private string GetSkipUnit(string pinValue, string measType)
        {
            string skipUnit = "";
            if (measType == MeasType.MeasCalc)
            {
                string testName = Regex.Match(pinValue.Replace(",", " "), RegCalcExpression, RegexOptions.IgnoreCase).Groups["testName"].ToString().Trim();
                if (testName.Contains(":"))
                {
                    int idx = testName.IndexOf(":", StringComparison.Ordinal);
                    skipUnit = testName.Substring(idx + 1, testName.Length - idx - 1);
                }
            }
            return skipUnit;
        }

        private string GetCapBit(string pinValue, string measType)
        {
            //eg: MeasC pin = JTAG(data):4 "testName", ccpBit is "4"
            if (measType == MeasType.MeasC)
            {
                return Regex.Match(pinValue, GetMeasCRegExpression(), RegexOptions.IgnoreCase).Groups["capBit"].ToString().Trim();
            }
            //eg: Limits "Cal_A" , (out1):4,
            if (measType == MeasType.MeasLimit)
            {
                return Regex.Match(pinValue, RegLimitsExpression, RegexOptions.IgnoreCase).Groups["capBit"].ToString().Trim();
            }
            return "";
        }

        private string GetMeasCRegExpression()
        {
            var result = new List<string>
            {
                RegPinName,
                RegStoreName,
                ":",
                RegBitNum,
                RegTName
            };
            return string.Join(@"\s*", result);
        }

        private string GetCommomnMeasRegExpression()
        {
            var result = new List<string> { RegPinName, RegStoreName, RegTName };
            return string.Join(@"\s*", result);
        }

        private string GetCusStr(string pinValue, string measType)
        {
            //eg: MeasC pin = JTAG(data):4 "testName", cusStr is "data"
            if (measType == MeasType.MeasC)
            {
                return Regex.Match(pinValue, GetMeasCRegExpression(), RegexOptions.IgnoreCase).Groups["cusStr"].ToString().Trim();
            }

            if (measType == MeasType.MeasCalc)
            {
                if (Regex.IsMatch(pinValue, RegCalcCusreg, RegexOptions.IgnoreCase))
                {
                    return Regex.Match(pinValue, RegCalcCusreg, RegexOptions.IgnoreCase).Groups["cusStr"].ToString().Trim();
                }
            }
            //eg: Limits "Cal_A" , (out1):4, cusStr is "out1" or eg:  Limits IO1(out2), cusStr is "out2"
            else if (measType == MeasType.MeasLimit)
            {
                if (Regex.IsMatch(pinValue, RegLimitsExpression, RegexOptions.IgnoreCase))
                {
                    return Regex.Match(pinValue, RegLimitsExpression, RegexOptions.IgnoreCase).Groups["cusStr"].ToString().Trim();
                }
            }
            //eg:MeasI pin=(VDDA,VDDB)(sn1), cusStr is "sn1",MeasF Pin = (IO1, IO3_P::IO3_N) (sn2)
            else
            {
                if (Regex.IsMatch(pinValue, GetCommomnMeasRegExpression(), RegexOptions.IgnoreCase))
                {
                    return Regex.Match(pinValue, GetCommomnMeasRegExpression(), RegexOptions.IgnoreCase).Groups["cusStr"].ToString().Trim();
                }
            }
            return "";
        }

        private string GetMeasPinName(string pinValue, string measType)
        {
            if (string.IsNullOrEmpty(measType))
            {
                return "";
            }

            //Limits, Calc => eg: Limits "Cal_A" , (out1):4 or Limits IO1(out2)
            if (measType == MeasType.MeasCalc || measType == MeasType.MeasLimit)
            {
                return HardIpConstData.FakePin;
            }

            //eg: MeasC pin = JTAG(data):4 "testName", pin name is "JTAG"
            if (measType == MeasType.MeasC)
            {
                return Regex.Match(pinValue, GetMeasCRegExpression(), RegexOptions.IgnoreCase).Groups["pinName"].ToString().Trim().ToUpper();
            }

            if (Regex.IsMatch(pinValue, GetCommomnMeasRegExpression(), RegexOptions.IgnoreCase))
            {
                return Regex.Match(pinValue, GetCommomnMeasRegExpression(), RegexOptions.IgnoreCase).Groups["pinName"].ToString().Trim();
            }
            //eg: MeasI pin= VDDA(sn1), pin name is "VDDA"
            //eg: MeasVdiff pin = pin_diff_Grp, this differential pin group exist in pin map, convert to "pin_P::pin_N"
            //because user can not write differential pin group name to patInfo, but do it in test plan.
            if (measType == MeasType.MeasVdiff || measType == MeasType.MeasIdiff || measType == MeasType.MeasVocm)
            {
                string name = Regex.Match(pinValue, RegPineNameExpression, RegexOptions.IgnoreCase).Groups["pinName"].ToString().Trim().ToUpper();
                if (name == "")
                {
                    name = pinValue;
                }

                return DataConvertor.ConvertDifferentialPinGroup(name);
            }
            if (Regex.IsMatch(pinValue, RegPineNameExpression, RegexOptions.IgnoreCase))
            {
                return Regex.Match(pinValue, RegPineNameExpression, RegexOptions.IgnoreCase).Groups["pinName"].ToString().Trim().ToUpper();
            }

            return pinValue.ToUpper();
        }

        private string GetMeasType(string measStr, int rowNum, int colNum)
        {
            //Limit
            if (Regex.IsMatch(measStr.Replace(",", " "), RegLimitsExpression, RegexOptions.IgnoreCase))
            {
                return MeasType.MeasLimit;
            }
            //Calc
            if (Regex.IsMatch(measStr, RegCalcExpression, RegexOptions.IgnoreCase))
            {
                return MeasType.MeasCalc;
            }
            //Meas
            if (Regex.IsMatch(measStr, RegMeasExpression, RegexOptions.IgnoreCase))
            {
                string type = Regex.Match(measStr, RegMeasExpression, RegexOptions.IgnoreCase).Groups["MeasType"].ToString();
                string realtype = MeasType.MeasTypes.Find(s => s.Equals(type, StringComparison.OrdinalIgnoreCase));
                return realtype ?? type;
            }

            if (Regex.IsMatch(measStr, MeasType.MeasWait, RegexOptions.IgnoreCase))
            {
                return MeasType.MeasWait;
            }
            if (Regex.IsMatch(measStr, MeasType.WiMeas, RegexOptions.IgnoreCase))
            {
                return MeasType.WiMeas;
            }
            if (Regex.IsMatch(measStr, MeasType.WiSrc, RegexOptions.IgnoreCase))
            {
                return MeasType.WiSrc;
            }
            if (Regex.IsMatch(measStr, MeasType.MeasWait, RegexOptions.IgnoreCase))
            {
                return MeasType.MeasWait;
            }
            if (Regex.IsMatch(measStr, MeasType.MeasN, RegexOptions.IgnoreCase))
            {
                return MeasType.MeasN;
            }
            ErrorReportManager.AddError(
                HardIpErrorType.E_WrongMeasContent_01,
                _planSheet.SheetName,
                rowNum,
                colNum,
                []
            );
            return MeasType.MeasLimit;
        }
        #endregion

        #region divideMeasPins
        private List<MeasPin> GetDivideMeasPins(PatternRow patternRow, List<MeasPin> measPins, HardIpPattern pattern)
        {
            if (measPins == null || measPins.Count == 0)
            {
                return measPins;
            }

            //if pin name is multiply pins, seperate it to single pins
            measPins = GetMeasPinsByName(measPins);

            if (measPins.Count == 0)
            {
                return measPins;
            }

            #region repeatLimit
            int count = SearchInfo.GetRepeatLimitCount(patternRow.MiscInfo);
            if (measPins.Count > 0)
            {
                measPins = GetMeasPinsByRepeat(measPins, count, pattern);
            }

            #endregion

            if (measPins.Count == 0)
            {
                return measPins;
            }

            measPins = GetMeasPinsByPowerMerge(measPins);

            return measPins;
        }

        private List<MeasPin> GetMeasPinsByName(List<MeasPin> measPins)
        {
            var measPinlist = new List<MeasPin>();
            for (int i = 0; i < measPins.Count; i++)
            {
                MeasPin measPin = measPins[i];
                string[] split = (measPin.PinName ?? string.Empty).Split(',');
                for (int s = 0; s < split.Length; s++)
                {
                    string pinName = split[s];
                    string trimmed = pinName.Trim();
                    MeasPin newPin = measPin.Copy();
                    newPin.PinName = trimmed;
                    newPin.PinCount = GetPinCount(trimmed);
                    newPin.VisitedTime = newPin.PinCount;
                    measPinlist.Add(newPin);
                }
            }

            return measPinlist;
        }

        private int GetPinCount(string pinNames)
        {
            if (string.IsNullOrEmpty(pinNames) || pinNames.Length == 0)
            {
                return 0;
            }

            if (pinNames.IndexOf(',') < 0)
            {
                string pin = pinNames;
                if (pin.IndexOf("::", StringComparison.Ordinal) >= 0)
                {
                    return 2;
                }

                if (TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(pin))
                {
                    List<string> pinGroupList = TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(pin);
                    return pinGroupList.Count;
                }

                return 1;
            }

            return 0;
        }

        private List<MeasPin> GetMeasPinsByRepeat(List<MeasPin> measPins, int count, HardIpPattern pattern)
        {
            if (measPins == null || measPins.Count == 0)
            {
                return measPins;
            }

            var testplanSeq = new List<TestPlanSequence>();
            int seqMax = measPins.Count == 0 ? 0 : measPins.Max(p => p.SequenceIndex);
            var newMeasPins = new List<MeasPin>();
            int j = 0;
            if (count > 0)
            {
                for (int i = 0; i <= count; i++)
                {
                    j = 0;
                    foreach (MeasPin measpin in measPins)
                    {
                        if (pattern.TestPlanSequences.Count > 0)
                        {
                            testplanSeq.Add(pattern.TestPlanSequences[j]);
                        }

                        MeasPin newpin = measpin.Copy();
                        newpin.SequenceIndex += seqMax * i;
                        newpin.CusStr += "_" + i;
                        newMeasPins.Add(newpin);
                        j++;
                    }
                }
                pattern.TestPlanSequences = testplanSeq;
                return newMeasPins;
            }

            int subI = 0;
            foreach (MeasPin measpin in measPins)
            {
                if (pattern.TestPlanSequences.Count > 0 && j < pattern.TestPlanSequences.Count)
                {
                    testplanSeq.Add(pattern.TestPlanSequences[j]);
                }

                MeasPin newpin = measpin.Copy();
                newpin.SequenceIndex += subI;
                newMeasPins.Add(newpin);

                if (measpin.RepeatCount > 0)
                {
                    newpin.CusStr += "_" + 0;
                    for (int k = 1; k <= measpin.RepeatCount; k++)
                    {
                        if (pattern.TestPlanSequences.Count > 0)
                        {
                            testplanSeq.Add(pattern.TestPlanSequences[j]);
                        }

                        newpin = measpin.Copy();
                        newpin.SequenceIndex += ++subI;
                        newpin.CusStr += "_" + k;
                        newMeasPins.Add(newpin);
                    }
                }
                j++;
            }
            pattern.TestPlanSequences = testplanSeq;
            return newMeasPins;
        }

        private List<MeasPin> GetMeasPinsByPowerMerge(List<MeasPin> measPins)
        {
            if (measPins == null || measPins.Count == 0)
            {
                return measPins;
            }

            if (TestPlanStatic.PowerMergeSheet != null && TestPlanStatic.PowerMergeSheet.PowerMerge != null && measPins.Count != 0)
            {
                var newMeasPins = new List<MeasPin>();
                foreach (MeasPin pin in measPins)
                {
                    List<string> newPins = _forceConditionResolver.JudgePinByPowerMerge(pin.PinName);
                    foreach (string newPin in newPins)
                    {
                        MeasPin newMeas = pin.Copy();
                        newMeas.PinName = newPin;
                        newMeasPins.Add(newMeas);
                    }
                }
                return newMeasPins;
            }

            return measPins;
        }

        #endregion

        private List<MeasPin> GetPostMeasPins(List<MeasPin> measPins)
        {
            foreach (MeasPin measPin in measPins)
            {
                var currentRangeCalculator = new MeasPinCurrentRangeCalculator(measPin);
                measPin.CurrentRangeList = currentRangeCalculator.GetCurrentRangeList();
                measPin.CurrentRangeListH = currentRangeCalculator.GetCurrentRangeListByVoltage(measPin.MeasLimitsH);
                measPin.CurrentRangeListL = currentRangeCalculator.GetCurrentRangeListByVoltage(measPin.MeasLimitsL);
                measPin.CurrentRangeListN = currentRangeCalculator.GetCurrentRangeListByVoltage(measPin.MeasLimitsN);
            }
            return measPins;
        }
    }
}
