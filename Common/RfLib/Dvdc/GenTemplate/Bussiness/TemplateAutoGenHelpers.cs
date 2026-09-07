using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.InputObject;

using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Static;

using RfLib.Dvdc.GenTemplate.TestPlanFormat;

using ScghLib.Reader;

using TestPlanLib.Basic.Relay;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    internal static class TemplateAutoGenHelpers
    {
        internal static string AccessWirelessTName(ProdCharSheetRow prodCharSheetRow, HardIpInfo hardIpInfo)
        {
            var resultList = new List<string>();
            List<string> nameGroup = [.. prodCharSheetRow.Item.Split('_')];
            string block = prodCharSheetRow.Mode.Replace("_", "") + "_" + nameGroup[0];
            //if (!string.IsNullOrEmpty(patinfo.MeasName.Trim(',').Trim()))
            //    return patinfo.MeasName;
            resultList.Add(block);
            string subBlock = "X";
            string itemField = "X";
            string efuseField = "X";
            string patField = "X";
            if (nameGroup.Count > 1)
            {
                subBlock = nameGroup[1];
            }

            if (nameGroup.Count > 2)
            {
                itemField = nameGroup[2];
            }

            if (!string.IsNullOrEmpty(hardIpInfo.TrimFuseName))
            {
                efuseField = hardIpInfo.TrimFuseName.Split(',').Length == 1 ? hardIpInfo.TrimFuseName.Replace("_", "") : "MultiFuse";
            }
            if (!string.IsNullOrEmpty(CommonGenerator.GetSubBlockNameByPattern(prodCharSheetRow.PayloadValue, prodCharSheetRow.Item)))
            {
                patField = CommonGenerator.GetSubBlockNameByPattern(prodCharSheetRow.PayloadValue, prodCharSheetRow.Item).Replace("_", "");
            }

            resultList.Add(subBlock);
            resultList.Add(itemField);
            resultList.Add(patField);
            resultList.Add(efuseField);

            return string.Join("_", resultList);
        }

        internal static TemplateRow GenerateHardIpReferencedPayloadRows(ProdCharSheetRow prodCharSheetRow, PatternRow patternRow, Dictionary<string, int> patternIndex, string block, List<TemplateRow> templateRows, ref int stepIndex)
        {
            #region pattern row
            var templateRow1 = new TemplateRow(patternIndex[block],
                patternIndex[block] + "." + stepIndex);
            templateRows.Add(templateRow1);
            templateRow1.Description = prodCharSheetRow.Block + "," + prodCharSheetRow.Item + "," + prodCharSheetRow.Mode + "\n";
            templateRow1.Description += "This item already reference from test plan";
            stepIndex++;
            TemplateRow templateRow2 = new TemplateRow(patternIndex[block],
                patternIndex[block] + "." + stepIndex)
            {
                TTR = GetEnableHLNInfo(prodCharSheetRow.EnableHlnv),
                Step = stepIndex.ToString(),
                Description = patternRow.Description,
                Pattern = patternRow.Pattern.TestPlanPatternName,

                ForceCondition =
                    string.Join(";", MergeForceCondition(prodCharSheetRow.ForceCondition, patternRow.ForceCondition.ForceCondition)),
                RegisterAssignment = patternRow.RegisterAssignment,
                MiscInfo = TemplateAutoGenHelpers1.MergeMiscInfo(patternRow.MiscInfo, prodCharSheetRow.InitList)
            };
            #region check TTR
            if (!string.IsNullOrEmpty(templateRow2.TTR))
            {//remove TTR specified items in test flow
                var miscinfoList = new List<string>();
                foreach (string ttrVoltage in templateRow2.TTR.Split(','))
                {
                    miscinfoList.Add(string.Format("Remove{0}", ttrVoltage));
                }
                templateRow2.MiscInfo = string.Format("{0};{1}", templateRow2.MiscInfo, string.Join(";", miscinfoList));
            }
            #endregion
            templateRows.Add(templateRow2);
            #endregion


            #region Non pattern row => Meas column

            foreach (PatChildRow childRow in patternRow.PatChildRows)
            {
                foreach (TestPlanRow row in ((PatSubChildRow)childRow).TpRows)
                {
                    var templateRow3 = new TemplateRow(patternIndex[block],
                        patternIndex[block] + "." + stepIndex)
                    {
                        Step = stepIndex.ToString(),
                        Description = row.Description,
                        ForceCondition =
                            row.ForceCondition,
                        MiscInfo = row.MiscInfo,
                        Meas = row.Meas
                    };
                    List<MeasLimit> limitSets = row.Limits;
                    foreach (MeasLimit measLimit in limitSets)
                    {
                        templateRow3.LoLimit.Add(measLimit.JobName, measLimit.LoLimit);
                        templateRow3.HiLimit.Add(measLimit.JobName, measLimit.HiLimit);

                    }
                    templateRows.Add(templateRow3);
                }
            }
            #endregion

            return templateRow2;
        }

        internal static string GetRelayOnMiscInfo(RelayItem relayItem)
        {
            var result = new List<string>();
            var relayOnGroup = relayItem.RelayStatusDictionary.Where(p => p.Value == "1").Select(p => p.Key).ToList();
            if (relayOnGroup.Count > 0)
            {
                result.Add(string.Format("RelayOn:{0};", string.Join(",", relayOnGroup)));
            }

            var relayOffGroup = relayItem.RelayStatusDictionary.Where(p => p.Value == "0").Select(p => p.Key).ToList();
            if (relayOffGroup.Count > 0)
            {
                result.Add(string.Format("RelayOff:{0};", string.Join(",", relayOffGroup)));
            }

            return string.Join("", result) + ";";
        }

        internal static string SelectPatternNameByType(ProdCharSheetRow prodCharSheetRow, EnumVbtFuncType enumVbtFuncType)
        {
            string patternName = "";
            if (enumVbtFuncType == EnumVbtFuncType.WiTrim)
            {
                patternName = prodCharSheetRow.PayloadValue;
            }
            else if (enumVbtFuncType == EnumVbtFuncType.RF || enumVbtFuncType == EnumVbtFuncType.RFTrim || enumVbtFuncType == EnumVbtFuncType.LCD)
            {
                patternName = prodCharSheetRow.PayloadValue;
            }
            else// HardIP
            {

                if (prodCharSheetRow.GetInitList().Count > 0)
                {
                    patternName += string.Join(",", prodCharSheetRow.GetInitList()) + ",";
                }

                patternName += string.Join(",", prodCharSheetRow.GetPayloadList());
            }
            return patternName;
        }

        internal static string SelectTestNameByType(ProdCharSheetRow prodCharSheetRow, HardIpInfo hardIpInfo)
        {
            if (LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                return "";
            }

            if (!string.IsNullOrEmpty(hardIpInfo.TestName))
            {
                return hardIpInfo.TestName;
            }

            return AccessWirelessTName(prodCharSheetRow, hardIpInfo);
        }

        internal static string GenerateRegsisterAssignment(HardIpInfo hardIpInfo, Dictionary<string, string> digSrcInfo)
        {
            string regsisterStr = "";
            var digSrcItems = new List<string>();
            var digSrcDict = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            if (!string.IsNullOrEmpty(hardIpInfo.DigSrcAssignment))
            {

                string regdigAssign = @"(?<reg>\w+):(?<data>((sweep\w*[\(\[].*[\)\]]|(\w))+))";
                foreach (string assignment in hardIpInfo.DigSrcAssignment.Split(';'))
                {
                    string regPartialBits = @"(?<Partial>\[\d:\d\])";
                    string part = "";
                    if (assignment.Contains('[') && Regex.IsMatch(assignment, regPartialBits, RegexOptions.IgnoreCase))
                    {
                        part = Regex.Match(assignment, regPartialBits, RegexOptions.IgnoreCase).Groups["Partial"].Value;

                    }

                    if (assignment.Contains(':'))
                    {
                        string assignment_tmp = assignment;
                        if (!string.IsNullOrEmpty(part))
                        {
                            assignment_tmp = assignment_tmp.Replace(part, "");
                        }

                        string reg = Regex.Match(assignment_tmp, regdigAssign, RegexOptions.IgnoreCase).Groups["reg"].Value;
                        string data = Regex.Match(assignment_tmp, regdigAssign, RegexOptions.IgnoreCase).Groups["data"].Value;
                        digSrcDict[reg] = data + part;
                    }
                    else
                    {
                        digSrcDict[assignment] = "";
                    }
                }
                //return string.Join(";", regItems);
            }
            //if (!digSrcDict.ContainsKey(info.TrimRegName) && !string.IsNullOrEmpty(info.TrimRegName))
            //    digSrcDict.Add(info.TrimRegName, "");
            if (hardIpInfo.SendBitStr.Length != 0 && hardIpInfo.SendBitName.Length != 0)
            {
                string multiTrimreg = hardIpInfo.TrimRegName.Count > 1 ? hardIpInfo.TrimRegName.Last() : "";
                double bitValue;
                string[] sendbits = hardIpInfo.SendBitStr.Split('+');
                string[] regs = hardIpInfo.SendBitName.Split('+');
                int i = 0;
                var usedReg = new List<string>();
                foreach (string reg_tmp in regs)
                {
                    string reg = reg_tmp.Split('[').FirstOrDefault()!;
                    if (usedReg.Contains(reg) || reg.EqualsIgnoreCase(multiTrimreg))
                    {
                        i++;
                        continue;

                    }
                    if (digSrcDict.ContainsKey(reg))
                    {
                        if (digSrcDict[reg].Length != 0)
                        {
                            if (double.TryParse(digSrcDict[reg], out bitValue))
                            {
                                digSrcItems.Add(string.Format("{0}=0b{1}", reg, digSrcDict[reg]));
                            }
                            else
                            {
                                digSrcItems.Add(string.Format("{0}={1}", reg, digSrcDict[reg]));
                            }
                        }
                        else
                        {
                            digSrcItems.Add(reg);
                        }
                    }
                    else if (digSrcInfo.ContainsKey(reg))
                    {
                        if (digSrcInfo[reg].Length != 0)
                        {
                            if (double.TryParse(digSrcInfo[reg], out bitValue))
                            {
                                digSrcItems.Add(string.Format("{0}=0b{1}", reg, digSrcInfo[reg]));
                            }
                            else
                            {
                                digSrcItems.Add(string.Format("{0}={1}", reg, digSrcInfo[reg]));
                            }
                        }
                        else
                        {
                            digSrcItems.Add(reg);
                        }
                    }
                    else
                    {
                        try
                        {
                            string bitWidth = sendbits[i].Split('_')[1];
                            string bitStr = TemplateAutoGenHelpers1.GenBitStr(bitWidth);
                            if (double.TryParse(bitStr, out bitValue))
                            {
                                digSrcItems.Add(string.Format("{0}=0b{1}", reg, bitStr));
                            }
                            else
                            {
                                digSrcItems.Add(string.Format("{0}={1}", reg, bitStr));
                            }
                        }
                        catch (Exception ex)
                        {
                            Response.Report("GenerateRegsisterAssignment: " + ex.Message);
                        }
                    }
                    usedReg.Add(reg);
                    i++;
                }
                //if (!string.IsNullOrEmpty(info.TrimRegName))
                //    digSrcItems.Add(info.TrimRegName);
                regsisterStr = string.Join(";", digSrcItems);
            }
            return regsisterStr;
        }

        internal static string SelectVBTFuncName(HardIpInfo hardIpInfo)
        {
            string funcName = "";
            if (hardIpInfo.MiscInfo.Count > 0)
            {
                foreach (string misc in hardIpInfo.MiscInfo)
                {
                    foreach (string item in misc.Split(';'))
                    {
                        if (item.Split(':').Length > 1 &&
                            item.Split(':')[0].EqualsIgnoreCase("Func"))
                        {
                            return "";
                        }
                    }
                }
            }
            if (!string.IsNullOrEmpty(funcName))
            {
                return "Func:" + funcName + ";";
            }
            return funcName;
        }

        internal static List<string> MergeForceCondition(string forceASet, string forceBSet)
        {
            var forceSets = new List<string> { forceASet, forceBSet };
            var result = new List<string>();
            var tmpForceDic = new Dictionary<string, List<string>>();
            #region collect force information
            foreach (string forceStr in forceSets.SelectMany(p => p.Split(';')))
            {
                if (string.IsNullOrEmpty(forceStr))
                {
                    continue;
                }
                if (forceStr.Split(':').Length == 3)
                {
                    string[] forcesgmts = forceStr.Split(':');
                    string forceKey = string.Format("{0}:{1}", forcesgmts[0], forcesgmts[1]);
                    if (!tmpForceDic.TryGetValue(forceKey, out List<string>? value))
                    {
                        value = [];
                        tmpForceDic.Add(forceKey, value);
                    }

                    value.AddRange(forcesgmts[2].Split(','));
                }
                else
                {
                    result.Add(forceStr);
                }
            }
            #endregion
            #region merge general force condition

            foreach (KeyValuePair<string, List<string>> forceCondition in tmpForceDic)
            {
                string forceValue = string.Join(",", forceCondition.Value.Distinct());
                result.Add(string.Format("{0}:{1}", forceCondition.Key, forceValue));
            }
            #endregion
            return [.. result.Distinct()];
        }

        internal static string GetEnableHLNInfo(string enableList)
        {
            if (string.IsNullOrEmpty(enableList))
            {
                return "";
            }
            var hLN = new List<string> { "HV", "LV", "NV" };
            foreach (string enableV in enableList.Split(','))
            {
                string? targetV = hLN.FirstOrDefault(p => p.EqualsIgnoreCase(enableV));
                if (!string.IsNullOrEmpty(targetV))
                {
                    hLN.Remove(targetV);
                }
            }
            return string.Join(",", hLN);
        }
    }
}
