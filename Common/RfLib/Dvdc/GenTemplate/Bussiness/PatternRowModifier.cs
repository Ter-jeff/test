using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Utility;

using RfLib.Dvdc.GenTemplate.TestPlanFormat;

using ScghLib.Reader;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    internal partial class PatternRowModifier(TemplateAutoGen templateAutoGen)
    {
        [GeneratedRegex("V|Hz|A", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex9();
        [GeneratedRegex(@"(?<unit>(A|V|Hz|ohm).*)", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex18();

        private readonly TemplateAutoGen _owner = templateAutoGen;

        public void ModifyPatternRow(WirelessTemplateRow wirelessTemplateRow, ProdCharSheetRow prodCharSheetRow, HardIpInfo hardIpInfo, EnumVbtFuncType enumVbtFuncType, bool isHardIPSheet, bool testNameFromScgh, List<string> dsscSetupNameTestPlan)
        {
            switch (enumVbtFuncType)
            {
                case EnumVbtFuncType.RF:
                case EnumVbtFuncType.RFTrim:
                    ModifyPatternRowForRf(wirelessTemplateRow, prodCharSheetRow, hardIpInfo, enumVbtFuncType, dsscSetupNameTestPlan);
                    break;
                case EnumVbtFuncType.WiTrim:
                    ModifyPatternRowForWiTrim(wirelessTemplateRow, prodCharSheetRow, hardIpInfo, enumVbtFuncType, dsscSetupNameTestPlan);
                    break;
                case EnumVbtFuncType.LCD:
                    ModifyPatternRowForLcd(wirelessTemplateRow, prodCharSheetRow, hardIpInfo, enumVbtFuncType, dsscSetupNameTestPlan);
                    break;
                default:
                    ModifyPatternRowForHardIp(wirelessTemplateRow, prodCharSheetRow, enumVbtFuncType, isHardIPSheet, testNameFromScgh, dsscSetupNameTestPlan);
                    break;
            }
        }

        private static void ModifyPatternRowForRf(WirelessTemplateRow wirelessTemplateRow, ProdCharSheetRow prodCharSheetRow, HardIpInfo hardIpInfo, EnumVbtFuncType enumVbtFuncType, List<string> dsscSetupNameTestPlan)
        {
            #region RF/RFTrim
            if (!string.IsNullOrEmpty(hardIpInfo.TrimTarget))
            {
                wirelessTemplateRow.TrimRegName.Add(hardIpInfo.TrimFuseName);
                wirelessTemplateRow.TrimTarget =
                    (ProcessTrimTarget(hardIpInfo.TrimTarget).FirstOrDefault(p => !string.IsNullOrEmpty(p)) ?? "").Replace(":", "::");
                wirelessTemplateRow.TrimType = string.IsNullOrEmpty(hardIpInfo.TrimType)
                    ? "DOALL"
                    : hardIpInfo.TrimType.Replace("&", "::");
            }
            if (prodCharSheetRow.GetInitList().Count > 0)
            {
                wirelessTemplateRow.Interpose = "InterposePrePat:RunInitPat(" + string.Join(",", prodCharSheetRow.GetInitList()) + ");";
            }

            if (hardIpInfo.DisconnectPrePat.Count > 0)
            {
                wirelessTemplateRow.Interpose += "InterposePreInit:" +
                                      ProcessDisConnectPins(hardIpInfo.DisconnectPrePat) + ";";
            }

            if (!string.IsNullOrEmpty(hardIpInfo.DisConnectPins))
            {
                wirelessTemplateRow.Interpose += "InterposePostMeas:" +
                                      ProcessDisConnectPins([.. hardIpInfo.DisConnectPins.Split('|')]) + ";";
            }

            wirelessTemplateRow.BestCode = hardIpInfo.BestCodeCalcFunc;

            if (prodCharSheetRow.GetPayloadList().Count > 1)
            {
                string[] payloadList2 = new string[prodCharSheetRow.GetPayloadList().Count];
                prodCharSheetRow.GetPayloadList().CopyTo(payloadList2);
                List<string> temp = [.. payloadList2];
                temp.RemoveAt(0);
                wirelessTemplateRow.Interpose += "InterposePostPat:RunPayloadPat(" + string.Join(",", temp) + ");";
            }

            //if ((type == VBTFuncType.RFTrim && (scgh.GetInitList().Any() || scgh.GetPayloadList().Count() > 1))|| type== VBTFuncType.RF)
            {
                string dsscSetupNameRf = TemplateAutoGenHelpers1.CheckDuplicateDSSCName(enumVbtFuncType + "_" + prodCharSheetRow.Block + "_" + CommonGenerator.GetSubBlockNameByPattern(prodCharSheetRow.PayloadValue, prodCharSheetRow.Item, false).Replace("_", "") + "_DSSCSetup", dsscSetupNameTestPlan);
                //template.RegisterAssignment = dsscSetupNameRf;
                wirelessTemplateRow.MiscInfo += string.Format("DsscSetup:{0};", dsscSetupNameRf);
                dsscSetupNameTestPlan.Add(dsscSetupNameRf);
            }
            #endregion
        }

        private void ModifyPatternRowForWiTrim(WirelessTemplateRow wirelessTemplateRow, ProdCharSheetRow prodCharSheetRow, HardIpInfo hardIpInfo, EnumVbtFuncType enumVbtFuncType, List<string> dsscSetupNameTestPlan)
        {
            #region Witrim
            List<string> trimTarget = ProcessTrimTarget(hardIpInfo.TrimTarget);
            wirelessTemplateRow.TrimRegName.Add(hardIpInfo.TrimFuseName);
            wirelessTemplateRow.TrimTarget = (trimTarget.FirstOrDefault(p => !string.IsNullOrEmpty(p)) ?? "").Replace(":", "::");
            wirelessTemplateRow.TrimType = string.IsNullOrEmpty(hardIpInfo.TrimType)
                ? "DOALL"
                : hardIpInfo.TrimType.Replace("&", "::");
            if (hardIpInfo.NewInfo != null)
            {

                wirelessTemplateRow.TrimMeas = GetTrimMeasInfo(hardIpInfo);
                wirelessTemplateRow.ForceCondition = GetTrimForceFondition(hardIpInfo);
            }
            if (string.IsNullOrEmpty(hardIpInfo.BestCodeCalcFunc))
            {
                wirelessTemplateRow.BestCode = wirelessTemplateRow.TrimTarget.Contains(';')
                    ? "algFindMultipleBestCode"
                    : "algFindBestCode";
            }
            else
            {
                wirelessTemplateRow.BestCode = hardIpInfo.BestCodeCalcFunc;
            }

            wirelessTemplateRow.MiscInfo += GenSearchLimit(hardIpInfo);
            //template.MiscInfo += _ProcessTrimPrePatAndPostPat(scgh);

            if (prodCharSheetRow.GetInitList().Count > 0)
            {
                wirelessTemplateRow.Interpose += "InterposePrePat:RunInitPat(" + string.Join(",", prodCharSheetRow.GetInitList()) +
                                      ");";
            }

            if (hardIpInfo.DisconnectPrePat.Count > 0)
            {
                wirelessTemplateRow.Interpose += "InterposePreInit:" +
                                      ProcessDisConnectPins(hardIpInfo.DisconnectPrePat) + ";";
            }

            if (!string.IsNullOrEmpty(hardIpInfo.DisConnectPins))
            {
                wirelessTemplateRow.Interpose += "InterposePostMeas:" +
                                      ProcessDisConnectPins([.. hardIpInfo.DisConnectPins.Split('|')]) + ";";
            }

            if (prodCharSheetRow.GetPayloadList().Count > 1)
            {
                string[] payloadList2 = new string[prodCharSheetRow.GetPayloadList().Count];
                prodCharSheetRow.GetPayloadList().CopyTo(payloadList2);
                List<string> temp = [.. payloadList2];
                temp.RemoveAt(0);
                wirelessTemplateRow.Interpose += "InterposePostPat:RunPayloadPat(" + string.Join(",", temp) + ");";

            }

            if (prodCharSheetRow.GetInitList().Count > 0 || prodCharSheetRow.GetPayloadList().Count > 1)
            {
                string dsscSetupNameWiTrim = TemplateAutoGenHelpers1.CheckDuplicateDSSCName(enumVbtFuncType + "_" + prodCharSheetRow.Block + "_" + CommonGenerator.GetSubBlockNameByPattern(prodCharSheetRow.PayloadValue, prodCharSheetRow.Item, false).Replace("_", "") + "_DSSCSetup", dsscSetupNameTestPlan);

                //template.MiscInfo += string.Format("InitpatDsscSetup:{0};", dsscSetupNameWiTrim);
                wirelessTemplateRow.MiscInfo += string.Format("DsscSetup:{0};", dsscSetupNameWiTrim);
                dsscSetupNameTestPlan.Add(dsscSetupNameWiTrim);
            }
            #endregion
        }

        private static void ModifyPatternRowForLcd(WirelessTemplateRow wirelessTemplateRow, ProdCharSheetRow prodCharSheetRow, HardIpInfo hardIpInfo, EnumVbtFuncType enumVbtFuncType, List<string> dsscSetupNameTestPlan)
        {
            #region LCD
            if (hardIpInfo.DisconnectPrePat.Count > 0)
            {
                wirelessTemplateRow.Interpose += "InterposePreInit:" +
                                      ProcessDisConnectPins(hardIpInfo.DisconnectPrePat) + ";";
            }

            if (!string.IsNullOrEmpty(hardIpInfo.DisConnectPins))
            {
                wirelessTemplateRow.Interpose += "InterposePostMeas:" +
                                      ProcessDisConnectPins([.. hardIpInfo.DisConnectPins.Split('|')]) +
                                      ";";
            }

            if (prodCharSheetRow.GetInitList().Count > 0)
            {
                wirelessTemplateRow.Interpose += "InterposePrePat:RunInitPat(" + string.Join(",", prodCharSheetRow.GetInitList()) +
                                      ");";
            }

            if (prodCharSheetRow.GetPayloadList().Count > 1)
            {
                string[] payloadList2 = new string[prodCharSheetRow.GetPayloadList().Count];
                prodCharSheetRow.GetPayloadList().CopyTo(payloadList2);
                List<string> temp = [.. payloadList2];
                temp.RemoveAt(0);
                wirelessTemplateRow.Interpose += "InterposePostPat:RunPayloadPat(" + string.Join(",", temp) + ");";
            }

            string dsscSetupNameLcd = TemplateAutoGenHelpers1.CheckDuplicateDSSCName(enumVbtFuncType + "_" + prodCharSheetRow.Block + "_" + CommonGenerator.GetSubBlockNameByPattern(prodCharSheetRow.PayloadValue, prodCharSheetRow.Item, false).Replace("_", "") + "_DSSCSetup", dsscSetupNameTestPlan);

            wirelessTemplateRow.MiscInfo += string.Format("DsscSetup:{0};", dsscSetupNameLcd);
            dsscSetupNameTestPlan.Add(dsscSetupNameLcd);
            #endregion
        }

        private static void ModifyPatternRowForHardIp(WirelessTemplateRow wirelessTemplateRow, ProdCharSheetRow prodCharSheetRow, EnumVbtFuncType enumVbtFuncType, bool isHardIPSheet, bool testNameFromScgh, List<string> dsscSetupNameTestPlan)
        {
            #region HardIP
            var hardIPInterPoseList = new List<string>();
            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                HardIpInfo payloadPatInfo = LocalSpecs.HardIpInfos.GetHardIpInfo(prodCharSheetRow.PayloadValue);
                RegisterLimitationGenerator.ProcessDigSrcDictionary(prodCharSheetRow);
                {
                    wirelessTemplateRow.Pattern = prodCharSheetRow.PayloadValue;
                    //if (DigSrcinfo.ContainsKey(scgh.PayloadValue))
                    //    template.RegisterAssignment = GenerateRegsisterAssignment(payloadPatInfo,
                    //            DigSrcinfo[scgh.PayloadValue]);
                }
                if (prodCharSheetRow.GetInitList().Count > 0)
                {
                    string runinitFormat = $"RunInitPat({string.Join(",", prodCharSheetRow.GetInitList())})";
                    if (string.IsNullOrEmpty(payloadPatInfo.MeasSeqStr) &&
                        string.IsNullOrEmpty(payloadPatInfo.SendBitStr) &&
                        string.IsNullOrEmpty(payloadPatInfo.CapBitStr))
                    {
                        wirelessTemplateRow.MiscInfo += $"StartOfBodyF:RFInterPose;StartOfBodyFArgs:{runinitFormat};";
                    }
                    else
                    {
                        hardIPInterPoseList.Add(runinitFormat);
                    }

                    string dsscSetupNameHardIP;
                    if (testNameFromScgh)
                    {
                        dsscSetupNameHardIP = TemplateAutoGenHelpers1.CheckDuplicateDSSCName(prodCharSheetRow.Item + "_DSSCSetup", dsscSetupNameTestPlan);
                    }
                    else
                    {
                        dsscSetupNameHardIP = TemplateAutoGenHelpers1.CheckDuplicateDSSCName(enumVbtFuncType + "_" + prodCharSheetRow.Block + "_" + CommonGenerator.GetSubBlockNameByPattern(prodCharSheetRow.PayloadValue, prodCharSheetRow.Item, false).Replace("_", "") + "_DSSCSetup", dsscSetupNameTestPlan);
                    }

                    //template.MiscInfo += string.Format("InitpatDsscSetup:{0};", dsscSetupNameWiTrim);

                    wirelessTemplateRow.MiscInfo += $"DsscSetup:{dsscSetupNameHardIP};";
                    dsscSetupNameTestPlan.Add(dsscSetupNameHardIP);

                }
                if (prodCharSheetRow.GetPayloadList().Count > 1)
                {
                    string[] payloadList2 = new string[prodCharSheetRow.GetPayloadList().Count];
                    prodCharSheetRow.GetPayloadList().CopyTo(payloadList2);
                    List<string> temp = [.. payloadList2];
                    temp.RemoveAt(0);
                    wirelessTemplateRow.MiscInfo += "Interpose_PostPat:RunInitPat(" + string.Join(",", temp) + ");";
                }
            }
            if (!isHardIPSheet)
            {
                hardIPInterPoseList.Add("RFTName");
            }

            if (hardIPInterPoseList.Count > 0)
            {
                wirelessTemplateRow.MiscInfo += string.Format("Interpose_PrePat:{0};", string.Join(";", hardIPInterPoseList));
            }

            wirelessTemplateRow.TestName = testNameFromScgh ? prodCharSheetRow.Item : wirelessTemplateRow.TestName;
            #endregion
        }

        private string GetTrimForceFondition(HardIpInfo hardIpInfo)
        {
            var forceSets = new List<string>();

            foreach (HardIpSeqInfoNew seq in hardIpInfo.NewInfo.SeqInfo)
            {
                var forceSubSets = new List<string>();
                foreach (ForcePin forcePin in seq.ForceConditions.ForcePins)
                {
                    if ((seq.MeasSeq != "N" &&
                        !seq.MeasPins.Exists(p => p.PinName.EqualsIgnoreCase(forcePin.PinName))) ||
                        TemplateAutoGen.MyRegex11().IsMatch(forcePin.ForceType))
                    {
                        forceSubSets.Add(string.Format("{0}:{1}:{2}", _owner.SearchPinInChannelMap(new MeasPin(forcePin.PinName, "")), forcePin.ForceType, forcePin.ForceValue));
                    }
                }
                forceSets.Add(string.Join(">", forceSubSets));
            }
            if (forceSets.All(p => string.IsNullOrEmpty(p.Replace(",", ""))))
            {
                return "";
            }
            return string.Join("|", forceSets);
        }

        private string GetTrimMeasInfo(HardIpInfo hardIpInfo)
        {
            //MeasVDM@AUX_DC_OUTP:AUX_DC_OUTN,1.8,0.05
            //FIMV @ PADIO_DCTPP,0,0.02,1.8,0.01
            try
            {
                List<string> trimSet = [];
                int trimTargetindex = hardIpInfo.TrimTarget.Split('|').ToList().FindIndex(p => !string.IsNullOrEmpty(p));
                int i = 0;
                string measStr = "";
                bool specialTargetFlag = hardIpInfo.NewInfo.SeqInfo.Exists(p => p.MeasPin.Count > 1);
                foreach (HardIpSeqInfoNew seq in hardIpInfo.NewInfo.SeqInfo)
                {
                    trimSet.Add(BuildTrimMeasSequenceString(seq, i, trimTargetindex, specialTargetFlag, ref measStr));
                    i++;
                }
                return string.Join(";\n", trimSet);
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                return "";
            }
        }

        private string BuildTrimMeasSequenceString(HardIpSeqInfoNew hardIpSeqInfoNew, int i, int trimTargetindex, bool specialTargetFlag, ref string measStr)
        {
            var trimSubSet = new List<string>();
            int j = 0;
            try
            {
                foreach (string pin in hardIpSeqInfoNew.MeasPin)
                {
                    MeasPin measpin = hardIpSeqInfoNew.MeasPins[j];
                    measStr = BuildTrimMeasPinString(hardIpSeqInfoNew, pin, measpin, i, j, trimTargetindex, specialTargetFlag, measStr);
                    trimSubSet.Add(measStr);
                    j++;
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
            foreach (MeasPin calcItem in hardIpSeqInfoNew.Calc)
            {
                trimSubSet.Add(string.Format("Calc@{0}", calcItem.CalcEqn));
            }

            return string.Join("&", trimSubSet);
        }

        private string BuildTrimMeasPinString(HardIpSeqInfoNew hardIpSeqInfoNew, string pin, MeasPin measPin, int i, int j, int trimTargetindex, bool specialTargetFlag, string previousMeasStr)
        {
            string targetPin =
                ((j == trimTargetindex) && specialTargetFlag
                    ? "Target_" + _owner.SearchPinInChannelMap(measPin)
                    : _owner.SearchPinInChannelMap(measPin)).Replace("&", "::");

            DetermineTrimForceInfo(hardIpSeqInfoNew, pin, targetPin, out string force, out string forceType, out string forcePin);

            ApplyTrimMeasPinLimits(hardIpSeqInfoNew, measPin, i);

            string irange = new MeasPinCurrentRangeCalculator(measPin).GetCurrentRange();
            string vrange = measPin.HighLimit;
            vrange = vrange.Length == 0 ? "1.8" : vrange;
            irange = irange.Length == 0 ? "0.02" : irange;
            forcePin = forcePin.Replace("&", "::");
            string measWait = string.IsNullOrEmpty(measPin.MeasWaitTime) ? "0.01" : measPin.MeasWaitTime;

            return BuildTrimMeasStrForCase(hardIpSeqInfoNew, targetPin, force, forceType, forcePin, irange, vrange, measWait, previousMeasStr);
        }

        private static void DetermineTrimForceInfo(HardIpSeqInfoNew hardIpSeqInfoNew, string pin, string targetPin, out string force, out string forceType, out string forcePin)
        {
            force = "";
            forceType = "";
            forcePin = "";
            //if force pin not same to meas pin, move force information to force condtion, not trimmeas. exception Case, MeasType = N would use FV or FI
            if (hardIpSeqInfoNew.ForcePin.Count == 0 || hardIpSeqInfoNew.MeasSeq.EqualsIgnoreCase("N"))
            {
                if (hardIpSeqInfoNew.ForcePin.Count == 0)
                {
                    forcePin = targetPin;
                }
                else
                {
                    forcePin = hardIpSeqInfoNew.ForcePin[0];
                    force = hardIpSeqInfoNew.ForceValue[0];
                }
                if (string.IsNullOrEmpty(hardIpSeqInfoNew.ForceType[0]))
                {
                    if (hardIpSeqInfoNew.MeasSeq.EqualsIgnoreCase("V"))
                    {
                        forceType = "I";
                    }
                    else
                    {
                        forceType = "V";
                    }
                }
                else
                {
                    forceType = hardIpSeqInfoNew.ForceType[0];
                    if (forceType != "V" && forceType != "I")
                    {
                        forceType = "";
                    }
                }
            }
            else if (hardIpSeqInfoNew.MeasSeq.EqualsIgnoreCase("DUTYCYCLE"))
            {
                force = ConvertValueUnit(hardIpSeqInfoNew.ExpectFreq[0]);
            }
            else
            {
                ForcePin? forceInfo =
                    hardIpSeqInfoNew.ForceConditions.ForcePins.FirstOrDefault(
                        p => p.PinName.EqualsIgnoreCase(pin));
                if (forceInfo != null)
                {
                    force = forceInfo.ForceValue;
                    forceType = forceInfo.ForceType;
                }
            }
        }

        private static void ApplyTrimMeasPinLimits(HardIpSeqInfoNew hardIpSeqInfoNew, MeasPin measPin, int i)
        {
            if (i < hardIpSeqInfoNew.Llimit.Count)
            {
                measPin.LowLimit = ConvertValueUnit(hardIpSeqInfoNew.Llimit[i]);
            }
            if (i < hardIpSeqInfoNew.Hlimit.Count)
            {
                measPin.HighLimit = ConvertValueUnit(hardIpSeqInfoNew.Hlimit[i]);
            }
            List<List<MeasLimit>> limits = GetLimits(measPin);
            measPin.MeasType = "Meas" + measPin.MeasType;
            measPin.MeasLimitsH = limits[0];
            measPin.MeasLimitsL = limits[1];
            measPin.MeasLimitsN = limits[2];
            measPin.ForceConditions.Add(hardIpSeqInfoNew.ForceConditions);
            measPin.CurrentRangeList = new MeasPinCurrentRangeCalculator(measPin).GetCurrentRangeList();
            measPin.MeasType = measPin.MeasType.Replace("Meas", "");
        }

        private string BuildTrimMeasStrForCase(HardIpSeqInfoNew hardIpSeqInfoNew, string targetPin, string force, string forceType, string forcePin, string irange, string vrange, string measWait, string previousMeasStr)
        {
            string measStr = previousMeasStr;
            switch (hardIpSeqInfoNew.MeasSeq.ToUpper())
            {
                case "F":
                    measStr = string.Format("MeasF@{0},{1}", targetPin, measWait);
                    break;
                case "VDM":
                    force = !string.IsNullOrEmpty(force) ? force : "0";
                    measStr = string.Format("MeasVDM@{0},{1},{2},{3}", targetPin, force, vrange, measWait);
                    break;
                case "R":
                case "R1":
                    force = !string.IsNullOrEmpty(force) ? force : "0";
                    measStr = string.Format("Meas{0}@{1},{2},{3},{4}", hardIpSeqInfoNew.MeasSeq, targetPin, force, double.Parse(force) / double.Parse(irange), measWait);
                    break;
                case "IDIFF":
                    force = !string.IsNullOrEmpty(force) ? force : "0";
                    measStr = string.Format("MeasIdiff@{0},{1},{2},{3}", targetPin, force, irange, measWait);
                    break;
                case "VDIFF":
                    force = !string.IsNullOrEmpty(force) ? force : "0";
                    measStr = string.Format("MeasVdiff@{0},{1},{2},{3}", targetPin, force, vrange, measWait);
                    break;
                case "N":
                    if (!string.IsNullOrEmpty(forceType) && !string.IsNullOrEmpty(forcePin))
                    {
                        measStr = string.Format("F{0}@{1},{2},{3},{4},{5}", forceType, _owner.SearchPinInChannelMap(new MeasPin(forcePin, "")), force, irange, vrange, measWait);
                    }
                    else
                    {
                        measStr = string.Format("N @,,,,");
                    }
                    break;
                case "I":
                    force = !string.IsNullOrEmpty(force) ? force : "0";
                    measStr =
                        string.Format("FVMI@{0},{1},{2},{3},{4}", targetPin, force, irange, "1.8", measWait);
                    break;
                case "V":
                    force = !string.IsNullOrEmpty(force) ? force : "0";
                    measStr =
                        string.Format("FIMV@{0},{1},{2},{3},{4}", targetPin, force, "0.02", vrange, measWait);
                    break;
                case "DUTYCYCLE":
                    force = !string.IsNullOrEmpty(force) ? force : "0";
                    measStr =
                        string.Format("DUTYCYCLE@{0},{1},,,{2}", targetPin, force, measWait);
                    break;
            }

            return measStr;
        }

        private static List<List<MeasLimit>> GetLimits(MeasPin measPin)
        {
            var limits = new List<List<MeasLimit>>();
            var limitsH = new List<MeasLimit>();
            var limitsL = new List<MeasLimit>();
            var limitsN = new List<MeasLimit>();
            limits.Add(limitsH);
            limits.Add(limitsL);
            limits.Add(limitsN);

            var limitH = new MeasLimit("CP1");
            var limitL = new MeasLimit("CP1");
            var limitN = new MeasLimit("CP1");
            limitsH.Add(limitH);
            limitsL.Add(limitL);
            limitsN.Add(limitN);

            //limit valus contains HV,LV,NV
            //user may assign unit "µ", equals to "u"
            string hi = measPin.HighLimit.Replace("µ", "u");
            string lo = measPin.LowLimit.Replace("µ", "u");

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

            return limits;
        }

        private static string GenSearchLimit(HardIpInfo hardIpInfo)
        {
            var limitsSeq = new List<string>();
            if (hardIpInfo.NewInfo == null)
            {
                return "";
            }

            int countindex = 0;

            foreach (HardIpSeqInfoNew seq in hardIpInfo.NewInfo.SeqInfo)
            {
                var limits = new List<string>();
                foreach (MeasPin meas in seq.MeasPins)
                {
                    string unit = "";
                    string lowLimit = MyRegex9().Replace(meas.LowLimit, "");
                    string highLimit = MyRegex9().Replace(meas.HighLimit, "");
                    if (meas.MeasType.ContainsIgnoreCase("v"))
                    {
                        unit = "V";
                    }
                    else if (meas.MeasType.ContainsIgnoreCase("f"))
                    {
                        unit = "Hz";
                    }
                    else if (meas.MeasType.ContainsIgnoreCase("i"))
                    {
                        unit = "A";
                    }

                    lowLimit = countindex < hardIpInfo.NewInfo.TrimLoopLLimit.Count
                        ? hardIpInfo.NewInfo.TrimLoopLLimit[countindex]
                        : "";
                    highLimit = countindex < hardIpInfo.NewInfo.TrimLoopHLimit.Count
                        ? hardIpInfo.NewInfo.TrimLoopHLimit[countindex]
                        : "";
                    string limit = string.IsNullOrEmpty(lowLimit) && string.IsNullOrEmpty(highLimit)
                        ? ""
                        : string.Format("{0},{1},{2}", lowLimit, highLimit, unit);
                    limits.Add(limit);
                    if (hardIpInfo.NewInfo.TrimLoopHLimit.Count > 0)
                    {
                        ;
                    }
                }
                limitsSeq.Add(string.Join("&", limits));
            }

            if (limitsSeq.Any(p => !string.IsNullOrEmpty(p.Replace("&", ""))))
            {
                return string.Format("SearchLimit:{0};", string.Join("|", limitsSeq));
            }

            return "";
        }

        private static List<string> ProcessTrimTarget(string targetInfo)
        {
            var result = new List<string>();
            string[] targets = targetInfo.Split(['|', '>'], StringSplitOptions.RemoveEmptyEntries);
            foreach (string target in targetInfo.Split('$'))
            {
                string regMeasc = @"\(\w+\)";
                if (Regex.IsMatch(target, regMeasc, RegexOptions.IgnoreCase))
                {
                    return [target];
                }
            }
            if (targets.Any(p => !string.IsNullOrEmpty(p)))
            {
                foreach (string target in targets.ToList())
                {

                    if (string.IsNullOrEmpty(target))
                    {
                        result.Add("");
                    }
                    else
                    {
                        char multiTrimDelimit = ';';
                        var multiTrimList = new List<string>();

                        foreach (string subTarget in target.Split(multiTrimDelimit))
                        {
                            string convertTar = ConvertValueUnit(subTarget);
                            multiTrimList.Add(convertTar);
                        }

                        result.Add(string.Join(";", multiTrimList));
                    }

                }
            }
            return result;
        }

        private static string ConvertValueUnit(string rawValue)
        {
            if (rawValue.Contains('&'))
            {
                List<string> unitvals = [];
                foreach (string rawval in rawValue.Split('&'))
                {
                    unitvals.Add(ConvertValueUnit(rawval));
                }

                return string.Join("&", unitvals);
            }

            string result;
            string scale = "";
            string regValue = @"^(?<value>-*\d+(\.\d+)*([Ee][\+-]\d*)*)";
            string regScale = @"(?<scale>[umkMG].*)";
            string unit = MyRegex18().Match(rawValue).Groups["unit"].ToString();
            //var regScip = @"(?<unit>e-*\d+)";
            //var regScim = @"(?<unit>e-*\d+)";
            //if (!Regex.IsMatch(rawValue, @"^(-*\d+)", RegexOptions.IgnoreCase))
            //    return rawValue;

            result = Regex.Match(rawValue, regValue).Groups["value"].ToString();
            if (rawValue.Contains('&'))
            {
                return rawValue.Replace("&", ":");
            }
            if (string.IsNullOrEmpty(result))
            {
                return rawValue;
            }
            if (!string.IsNullOrEmpty(unit))
            {
                scale = Regex.Match(rawValue.Replace(unit, ""), regScale, RegexOptions.IgnoreCase).Groups["scale"].ToString();
            }
            string unit_after = "";
            if (rawValue.Contains('['))
            {
                return rawValue;
            }
            else if (!string.IsNullOrEmpty(scale))
            {

                switch (scale)
                {
                    case "u":
                        unit_after = "E-6";
                        break;
                    case "m":
                        unit_after = "E-3";
                        break;
                    case "k":
                    case "K":
                        unit_after = "E3";
                        break;
                    case "M":
                        unit_after = "E6";
                        break;
                    case "G":
                        unit_after = "E9";
                        break;
                    default:
                        break;
                }

            }
            //else if (Regex.IsMatch(rawValue, regSci, RegexOptions.IgnoreCase))
            //{
            //    unit_after = Regex.Match(rawValue, regSci, RegexOptions.IgnoreCase).Groups["unit"].ToString().ToUpper();
            //}
            //if (Regex.IsMatch(unit_after, regSci, RegexOptions.IgnoreCase))
            //{
            //    var regPower = Regex.Match(unit_after, @"e(?<Power>-*\d+)", RegexOptions.IgnoreCase).Groups["Power"].Value;
            //    var calcValue = double.Parse(result) * Math.Pow(10, int.Parse(regPower));
            //    return Convert.ToString(calcValue) + unit;
            //}
            return result + unit + unit_after;
        }

        private static string ProcessDisConnectPins(List<string> disConnectPins, string delimiter = "|")
        {
            if (disConnectPins.Count == 0)
            {
                return "";
            }
            var result = new List<string>();
            foreach (string disConnectPin in disConnectPins)
            {
                if (!string.IsNullOrEmpty(disConnectPin))
                {
                    string powerDownStr = string.Format("DisconnectPins({0})", disConnectPin);
                    result.Add(powerDownStr);
                }
                else
                {
                    result.Add(disConnectPin);
                }
            }
            if (result.All(string.IsNullOrEmpty))
            {
                return "";
            }
            return string.Join(delimiter, result);
        }
    }
}
