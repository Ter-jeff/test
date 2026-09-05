using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Static;
using LogLib.Utility;

using RfLib.Dvdc.GenTemplate.TestPlanFormat;

using ScghLib.Reader;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    internal partial class TemplateAutoGenHelpers1
    {
        [GeneratedRegex("V|Hz|A", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex10();
        [GeneratedRegex("sweep", RegexOptions.IgnoreCase, "en-US")]
        internal static partial Regex MyRegex11();
        [GeneratedRegex(@"[\>\,\[\]\|]", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex16();
        [GeneratedRegex("isFW|Func:rffunc_trx_universal", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex17();
        [GeneratedRegex("subblock", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex21();

        public static Dictionary<string, List<string>> CheckDupRegName()
        {
            Response.Report(string.Format("Check Duplicate Reg Name ..."), percentage: Convert.ToInt32(60));

            var duplicateReg = new Dictionary<string, List<string>>();

            foreach (ProdCharSheetRow scgh in ScghStatic.ScghData.ConvertedPatternRowListByHardip)
            {
                var allPatterns = new List<string>();
                allPatterns.AddRange(scgh.GetInitList());
                allPatterns.AddRange(scgh.GetPayloadList());

                foreach (string pattern in allPatterns)
                {
                    HardIpInfo patinfo = LocalSpecs.HardIpInfos.GetHardIpInfo(pattern);
                    if (string.IsNullOrEmpty(patinfo.DigSrcAssignment))
                    {
                        continue;
                    }
                    if (duplicateReg.ContainsKey(pattern))
                    {
                        continue;
                    }
                    List<string> assignment = [.. patinfo.DigSrcAssignment.Split(';')];
                    var assignfunc = new List<string>();
                    var tmpdupfunc = new List<string>();

                    try
                    {
                        foreach (string assign in assignment)
                        {
                            if (assignfunc.Contains(assign.Split(':')[0]))
                            {
                                tmpdupfunc.Add(assign.Split(':')[0]);
                            }
                            else
                            {
                                assignfunc.Add(assign.Split(':')[0]);
                            }
                        }
                        if (tmpdupfunc.Count > 0)
                        {
                            duplicateReg.Add(pattern, [.. tmpdupfunc.Distinct()]);
                        }

                    }
                    catch (Exception ex)
                    {
                        ErrorMessageBox.Show(string.Format(ex.ToString()));
                    }

                }
            }
            return duplicateReg;

        }

        public static Dictionary<string, List<string>> CheckMissingReg()
        {
            Response.Report(string.Format("Check Missing Reg ..."), percentage: Convert.ToInt32(55));

            var notmatchDic = new Dictionary<string, List<string>>();

            foreach (ProdCharSheetRow scgh in ScghStatic.ScghData.ConvertedPatternRowListByHardip)
            {
                var allPatterns = new List<string>();
                allPatterns.AddRange(scgh.GetInitList());
                allPatterns.AddRange(scgh.GetPayloadList());

                foreach (string pattern in allPatterns)
                {
                    HardIpInfo patinfo = LocalSpecs.HardIpInfos.GetHardIpInfo(pattern);
                    if (string.IsNullOrEmpty(patinfo.DigSrcAssignment))
                    {
                        continue;
                    }
                    if (notmatchDic.ContainsKey(pattern))
                    {
                        continue;
                    }
                    List<string> allreg = [.. patinfo.SendBitName.Split('+')];
                    List<string> assignment = [.. patinfo.DigSrcAssignment.Split(';')];
                    List<string> missingReg = [];
                    try
                    {
                        foreach (string reg in assignment)
                        {
                            if (!allreg.Exists(p => p.EqualsIgnoreCase(reg.Split(':')[0])))
                            {
                                missingReg.Add(reg.Split(':')[0]);
                            }
                        }

                        if (missingReg.Count > 0)
                        {
                            notmatchDic.Add(pattern, missingReg);
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorMessageBox.Show(string.Format(ex.ToString()));
                    }

                }
            }

            return notmatchDic;

        }

        internal static string CheckDuplicateDSSCName(string oriDSSCName, List<string> currentSetupList)
        {
            string result = oriDSSCName;
            if (currentSetupList.Contains(oriDSSCName))
            {
                List<string> segs = [.. oriDSSCName.Split('_')];
                if (!int.TryParse(segs.Last(), out int dupValue))
                {
                    result += "_1";
                }
                else
                {
                    List<string> excepLast = segs.GetRange(0, segs.Count - 1);
                    result = string.Join("_", excepLast) + "_" + (dupValue + 1);
                }
                if (currentSetupList.Contains(result))
                {
                    result = CheckDuplicateDSSCName(result, currentSetupList);
                }
                else
                {
                    currentSetupList.Add(result);
                }
            }
            return result;
        }

        //public Dictionary<string, List<TemplateRow>> GenTemplatesByPatList(string sheetname, Dictionary<string, PatternData> patternList)
        //{
        //    var results = new Dictionary<string, List<TemplateRow>>();
        //    string block = sheetname == "" ? "HardIP_All" : sheetname;
        //    var blockTemplates = new List<TemplateRow>();
        //    results.Add(block, blockTemplates);
        //    int itemIndex = 1;

        //    foreach (var patDataDic in patternList)
        //    {

        //        // add filter , only add Hard_IP pattern by JN 20171031
        //        if (patDataDic.Value.TpCategory.ToUpper() != "HARD_IP")
        //            continue;

        //        int stepIndex = 0;
        //        var patInfo = SearchInfo.GetHardIpInfo(patDataDic.Key);

        //        //Describtion row for block+item+mode
        //        var templateRow1 = new TemplateRow(itemIndex, itemIndex + "." + stepIndex);
        //        blockTemplates.Add(templateRow1);
        //        templateRow1.Description = "HardIP Pattern from pattern list file";
        //        if (patInfo.MeasSeqStr != "")
        //            templateRow1.Description += "  MeasSeq: " + patInfo.MeasSeqStr;
        //        templateRow1.TestItem = itemIndex;
        //        templateRow1.Step = itemIndex + "." + stepIndex;
        //        stepIndex++;
        //        //Pattern row
        //        var templateRow2 = new TemplateRow(itemIndex, itemIndex + "." + stepIndex);
        //        blockTemplates.Add(templateRow2);
        //        templateRow2.Pattern = patDataDic.Key;
        //        templateRow2.RegisterAssignment = GenerateRegsisterAssignment(patInfo, new Dictionary<string, string>());
        //        templateRow2.Description = "Run the pattern provided";
        //        templateRow2.Step = itemIndex + "." + stepIndex;
        //        stepIndex++;

        //        foreach (var seqinfo in patInfo.SeqInfo)
        //        {
        //            if (Regex.IsMatch(seqinfo.SeqName, "VDIFF|IDIFF", RegexOptions.IgnoreCase))
        //            {
        //                if (seqinfo.PinList.Contains("::"))
        //                    _GenPlanWithVdiff(blockTemplates, describe, seqinfo, itemIndex, stepIndex);
        //                else
        //                {
        //                    string outString = string.Format("The name of VDIFF {0} is wrong", seqinfo.PinList);
        //                    Response.Report(outString, MessageLevel.Error, Convert.ToInt32(100));
        //                }
        //            }
        //            else
        //                _GenPlanWithNonVdiff(blockTemplates, describe, seqinfo, itemIndex, stepIndex);
        //            stepIndex++;
        //        }

        //        if (patInfo.DsscOut != "" && ProjectConfigLocalSpecs.Device != DeviceEnum.LCD)
        //        {
        //            _GenMeasCTemplateItem(patInfo, blockTemplates, itemIndex, stepIndex);
        //        }
        //        itemIndex++;
        //    }
        //    return results;
        //}

        /// <summary>
        /// Generate template register assignment by SentBitStr and SendBitName
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        /// <summary>
        /// initial register bit string to all "0" according to bit width
        /// </summary>
        /// <param name="bitWidth"></param>
        /// <returns></returns>
        internal static string GenBitStr(string bitWidth)
        {
            string bitStr = "";
            for (int i = 0; i < Convert.ToInt32(bitWidth); i++)
            {
                bitStr += "0";
            }
            return bitStr;
        }

        internal static void GenerateCalcEquation(List<TemplateRow> templateRows, HardIpSeqInfoNew hardIpSeqInfoNew, string seqindex, int blockindex, int stepIndex)
        {
            int pinIndex = 1;

            foreach (MeasPin calcItem in hardIpSeqInfoNew.Calc)
            {
                //Meas information row
                var newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex)
                {
                    TestItem = blockindex,
                    TestName = calcItem.TestName,
                    Step = blockindex + "." + stepIndex
                };

                templateRows.Add(newTempRow);
                newTempRow.Seqindex = seqindex;
                if (!string.IsNullOrEmpty(calcItem.CusStr))
                {
                    newTempRow.Meas = string.Format(@"Calc ({0})({1})", calcItem.CalcEqn, calcItem.CusStr);
                }
                else
                {
                    newTempRow.Meas = string.Format("Calc {0}", calcItem.CalcEqn);
                }

                if (!string.IsNullOrEmpty(calcItem.CusStr))
                {
                    newTempRow.Description = string.Format("CalcEquation:{0},With Hi/Lo Limit{1} / {2}", calcItem.CalcEqn, calcItem.HighLimit, calcItem.LowLimit);
                }

                newTempRow.HiLimit.Add("CP1", calcItem.HighLimit);
                newTempRow.LoLimit.Add("CP1", calcItem.LowLimit);

                newTempRow.Step = blockindex + "." + stepIndex;
                newTempRow.Step += "." + pinIndex;
                pinIndex++;
            }
        }

        internal static void GetBestValue(WirelessTemplateRow wirelessTemplateRow, HardIpInfoNew hardIpInfoNew, bool isReadCap)
        {
            if (isReadCap)
            {
                return;
            }

            MeasPin? target =
                hardIpInfoNew.SeqInfo.SelectMany(p => p.MeasPins)
                    .FirstOrDefault(p => !string.IsNullOrEmpty(p.LowLimit) || !string.IsNullOrEmpty(p.HighLimit));
            string lowLimit = "";
            string highLimit = "";
            if (target != null)
            {
                lowLimit = MyRegex10().Replace(target.LowLimit, "");
                highLimit = MyRegex10().Replace(target.HighLimit, "");
            }
            if (!wirelessTemplateRow.TestName.Contains("TARGET-DIFF"))
            {
                wirelessTemplateRow.HiLimit.Add("CP1", highLimit);
                wirelessTemplateRow.LoLimit.Add("CP1", lowLimit);
            }
            //foreach (var meas in info.SeqInfo.SelectMany(p => p.MeasPins))
            //{
            //        var lowLimit = Regex.Replace(meas.LowLimit, "V|Hz|A", "", RegexOptions.IgnoreCase);
            //        var highLimit = Regex.Replace(meas.HighLimit, "V|Hz|A", "", RegexOptions.IgnoreCase);
            //        //if (meas.MeasType.ToLower().Contains("v"))
            //        //    unit = "V";
            //        //else if (meas.MeasType.ToLower().Contains("f"))
            //        //    unit = "Hz";
            //        //else if (meas.MeasType.ToLower().Contains("i"))
            //        //    unit = "A";
            //        template.HiLimit.Add("CP1", highLimit);
            //        template.LoLimit.Add("CP1",lowLimit);
            //        if (!string.IsNullOrEmpty(lowLimit) || !string.IsNullOrEmpty(highLimit)) break;

            //}
        }

        internal static EnumVbtFuncType GetFuncType(ProdCharSheetRow prodCharSheetRow)
        {
            EnumVbtFuncType type = EnumVbtFuncType.HardIP;
            HardIpInfo info = LocalSpecs.HardIpInfos.GetHardIpInfo(prodCharSheetRow.PayloadValue);
            if (info.TrimTarget.Split(['|', '>'], StringSplitOptions.RemoveEmptyEntries).Length != 0)
            {
                if (info.NewInfo != null &&
                    Regex.IsMatch(info.NewInfo.MeasSeq, string.Format("{0}|{1}", MeasType.WiSrc, MeasType.WiMeas), RegexOptions.IgnoreCase))
                {
                    type = EnumVbtFuncType.RFTrim;
                }
                else
                {
                    type = EnumVbtFuncType.WiTrim;
                }
            }
            else if (LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                type = EnumVbtFuncType.LCD;
            }
            else if (info.NewInfo != null && (!string.IsNullOrEmpty(MyRegex16().Replace(info.NewInfo.RfSetup, ""))))
            {
                type = EnumVbtFuncType.RF;
            }
            if (info.MiscInfo.Exists(p => MyRegex17().IsMatch(p)))
            {
                type = EnumVbtFuncType.RF;
            }
            return type;
        }

        internal static List<int> GetTrimRegBit(HardIpInfo hardIpInfo)
        {
            var result = new List<int>();
            string trimRegName = string.Join(";", hardIpInfo.TrimRegName);
            if (hardIpInfo.TrimRegName.Count > 0)
            {
                foreach (string reg in hardIpInfo.TrimRegName)
                {
                    List<string> allregs = [.. hardIpInfo.SendBitName.Split('+')];
                    List<string> allregs_bits = [.. hardIpInfo.SendBitStr.Split('+')];
                    var bitList = new List<int>();
                    {
                        int trim_reg_index = allregs.FindIndex(p => p.EqualsIgnoreCase(reg));
                        if (trim_reg_index != -1)
                        {
                            result.Add(int.Parse(allregs_bits[trim_reg_index].Split('_')[1]));
                        }
                        else
                        {
                            bitList.Add(0);
                        }
                    }
                }
                return result;
            }
            return result;
        }

        internal static string MergeMiscInfo(string referenceItem, List<string> inits)
        {
            //replace subblock
            var result = new List<string>();
            foreach (string misc in referenceItem.Split(';'))
            {
                List<string> miscSgmt = [.. misc.Split(':')];
                if (miscSgmt.Count != 2 || !MyRegex21().IsMatch(miscSgmt[0]))
                {
                    result.Add(misc);
                }
                else
                {
                    string subblock = miscSgmt[1];
                    if (!string.IsNullOrEmpty(CommonGenerator.GetInitSubBlockName(inits)))
                    {
                        result.Add(string.Format("SubBlock:{0}-{1}", subblock, CommonGenerator.GetInitSubBlockName(inits)));
                    }
                }
            }

            return string.Join(";", result);
        }
    }
}
