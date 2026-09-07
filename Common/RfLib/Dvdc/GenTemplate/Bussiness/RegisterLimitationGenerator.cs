using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.GenerateIgxl.Wireless.DVDC.WirelessConst;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Extension;

using LogLib.Utility;

using ScghLib.Reader;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    internal class RegisterLimitationGenerator(TemplateAutoGen templateAutoGen)
    {
        private readonly TemplateAutoGen _owner = templateAutoGen;

        internal string DetectRegisterAssignLimitation(ProdCharSheetRow prodCharSheetRow, string assignData)
        {
            string result = assignData;
            if (assignData.Length >= 6000)
            {
                var row = new HardIpPattern
                {
                    Pattern = new PatternClass(prodCharSheetRow.PayloadValue),
                    SheetName = string.Format("HardIP_{0}", prodCharSheetRow.Block)
                };
                var registerAssignItem = new HardIpRegAssign
                {
                    SubBlockName = CommonGenerator.GetRegAssignName(row),
                    Type = RegisterAssignType.DigSrc_Assignment
                };
                var regAssginList = new List<List<string>>();
                foreach (string digSrcitem in assignData.Split(';'))
                {
                    List<string> data = [];
                    List<string> arr = [.. digSrcitem.Split('=')];
                    data.Add(arr[0]);
                    if (arr.Count > 1)
                    {
                        data.Add("'" + arr[1]);
                    }
                    regAssginList.Add(data);
                }
                registerAssignItem.RegAssignList = regAssginList;
                _owner.HardIpInputDataInternal.HardIpRegAssigns.Add(registerAssignItem);
                result = string.Format("Reg_assign:{0}", registerAssignItem.SubBlockName);
            }
            return result;
        }

        internal string DetectDigSrcEqnLimitation(ProdCharSheetRow prodCharSheetRow, string assignData)
        {
            string result = "";
            if (assignData.Length >= 6000)
            {
                var row = new HardIpPattern
                {
                    Pattern = new PatternClass(prodCharSheetRow.PayloadValue),
                    SheetName = string.Format("HardIP_{0}", prodCharSheetRow.Block)
                };
                var registerAssignItem = new HardIpRegAssign
                {
                    SubBlockName = CommonGenerator.GetRegAssignName(row),
                    Type = RegisterAssignType.DigSrc_Equation
                };
                var regAssginList = new List<List<string>>();
                foreach (string digSrcitem in assignData.Split('+'))
                {
                    List<string> data = [];
                    List<string> arr = [.. digSrcitem.Split('=')];
                    data.Add(arr[0]);
                    if (arr.Count > 1)
                    {
                        data.Add("'" + arr[1]);
                    }
                    regAssginList.Add(data);
                }
                registerAssignItem.RegAssignList = regAssginList;
                _owner.HardIpInputDataInternal.HardIpRegAssigns.Add(registerAssignItem);
                result = string.Format("{1}:\"Reg_assign:{0}\";", registerAssignItem.SubBlockName, WirelessConstData.DigSrcEquation);
            }
            return result;
        }

        internal static string GetDCSpec(string scghDC, List<string> initList)
        {
            if (!string.IsNullOrEmpty(scghDC))
            {
                return string.Format("DC:{0}", scghDC);
            }
            string result = "";
            string type = "";
            string domain = "X";
            string pmode = "X";
            foreach (string init in initList)
            {
                List<string> sgmts = [.. init.Split('_')];
                if (sgmts.Count <= 9)
                {
                    continue;
                }

                {
                    if (sgmts[2].EqualsIgnoreCase("C"))
                    {
                        domain = "cpu";
                    }
                    else if (sgmts[2].EqualsIgnoreCase("L"))
                    {
                        domain = "gfx";
                    }
                    else if (sgmts[2].EqualsIgnoreCase("S"))
                    {
                        domain = "soc";
                    }
                }

                if (sgmts[4].EqualsIgnoreCase("BI"))
                {
                    type = "BIST";
                }

                if (!sgmts[9].EqualsIgnoreCase("ALLFRV"))
                {
                    pmode = sgmts[9];
                }

            }

            if (!string.IsNullOrEmpty(type))
            {
                string chiplet = MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.GetChipletList(TestPlanStatic.PowerInfoSheet).FirstOrDefault() ?? "";
                result = MultiTestSettingSheetsSingleton.Instance()
                    .FindMbistCatgeoryName(domain, type, pmode, initList, out _, out _, chiplet);
            }
            if (!string.IsNullOrEmpty(result))
            {
                return string.Format("DC:{0};", result);
            }
            return result;
        }

        internal static Dictionary<string, Dictionary<string, string>> ProcessDigSrcDictionary(ProdCharSheetRow prodCharSheetRow)
        {
            var digSrcDictionary = new Dictionary<string, Dictionary<string, string>>();
            var allPatterns = new List<string>();
            allPatterns.AddRange(prodCharSheetRow.GetInitList());
            allPatterns.AddRange(prodCharSheetRow.GetPayloadList());
            var curDigSrcDictionary = new Dictionary<string, string>();
            foreach (string pattern in allPatterns)
            {
                HardIpInfo patinfo = LocalSpecs.HardIpInfos.GetHardIpInfo(pattern);
                SearchInfo.ModDuplicateRegName(patinfo);
                if (string.IsNullOrEmpty(patinfo.SendBitName))
                {
                    continue;
                }
                List<string> allreg = [.. patinfo.SendBitName.Split('+')];
                List<string> allsgmt = [.. patinfo.SendBitStr.Split('+')];
                List<string> assignment = [.. patinfo.DigSrcAssignment.Split(';')];

                try
                {
                    int i = 0;
                    if (digSrcDictionary.ContainsKey(pattern))
                    {
                        continue;
                    }
                    digSrcDictionary.Add(pattern, []);
                    //store pattern digsrc information

                    foreach (string regs in allreg)
                    {
                        string reg = regs.Split('[').FirstOrDefault()!;
                        if (patinfo.TrimRegName.Exists(p => p.EqualsIgnoreCase(reg)))
                        {
                            int regIndex = patinfo.TrimRegName.FindIndex(p => p.EqualsIgnoreCase(reg));
                            if (regIndex != -1)
                            {
                                if (regIndex < patinfo.TrimFuseName.Split(',').Length)
                                {
                                    curDigSrcDictionary[reg] = patinfo.TrimFuseName.Split(',')[regIndex];
                                }
                            }
                        }
                        else if (!curDigSrcDictionary.ContainsKey(reg))
                        {
                            curDigSrcDictionary[reg] = TemplateAutoGenHelpers1.GenBitStr(allsgmt[i].Split('_')[1]);
                        }
                        if (
                            assignment.Exists(
                                p => p.Contains(':') && p.Split(':')[0].EqualsIgnoreCase(reg)))
                        {
                            string registerMatch = assignment.FirstOrDefault(
                                    p => p.Split(':')[0].EqualsIgnoreCase(reg))!;
                            string regPartialBits = @"(?<Partial>\[\d:\d\])";
                            string part = Regex.Match(registerMatch, regPartialBits, RegexOptions.IgnoreCase).Groups["Partial"].Value;
                            if (!string.IsNullOrEmpty(part))
                            {
                                curDigSrcDictionary[reg] = registerMatch.Replace(part, "").Split(':')[1] + part;
                            }
                            else
                            {
                                curDigSrcDictionary[reg] = registerMatch.Split(':')[1] + part;
                            }
                        }
                        if (!digSrcDictionary[pattern].ContainsKey(reg) && curDigSrcDictionary.TryGetValue(reg, out string? value))
                        {
                            digSrcDictionary[pattern].Add(reg, value);
                        }
                        i++;
                    }
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
                if (digSrcDictionary[pattern].Count > 0)
                {
                    ;
                }
            }

            return digSrcDictionary;
        }
    }
}
