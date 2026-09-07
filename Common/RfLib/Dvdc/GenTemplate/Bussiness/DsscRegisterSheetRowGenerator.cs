using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Extension;

using LogLib.Static;

using RfLib.Dvdc.Reader.CapturePostProcess;
using RfLib.Dvdc.Reader.DsscSetup;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    internal static partial class DsscRegisterSheetRowGenerator
    {
        [GeneratedRegex(@".*\[\w+\]", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"\[(?<Partial>\d)\]", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();

        internal static List<DsscSetupSheetRow>? CreateRegisterSheetRows(string pattern, string dsscSetupName, ref bool writeDssc, EnumVbtFuncType enumVbtFuncType, Dictionary<string, Dictionary<string, string>> digSrcDictionary, ref List<PostProcessSheetRow> postProcessSheetRows, List<string> dsscPostProcessNameSheet)
        {
            HardIpInfo patInfo = LocalSpecs.HardIpInfos.GetHardIpInfo(pattern);
            if (patInfo == null)
            {
                return null;
            }
            Dictionary<string, string> curPatDigSrcInfo = digSrcDictionary.TryGetValue(pattern, out Dictionary<string, string>? patternDigSrcInfo)
                ? patternDigSrcInfo
                : [];
            HashSet<string> regSrcList = BuildRegSrcList(patInfo);

            GetCaptureInfoFromMiscInfo(patInfo, out string? cppSetupName, out string? cppBitWidth);

            //DSSCSetup_PostProcess
            if (regSrcList.Count == 0 && string.IsNullOrEmpty(patInfo.DsscOut) && string.IsNullOrEmpty(cppSetupName))
            {
                return CreateSingleSetupOnlyRow(pattern, dsscSetupName, ref writeDssc, enumVbtFuncType);
            }

            string eqnStr = patInfo.SendBitName;
            string dsscOutStr = string.IsNullOrEmpty(patInfo.CapStoreName) ? patInfo.DsscOut : AddStoreNameToDsscout(patInfo);
            if (regSrcList.Count == 0)
            {
                return CreateRegisterSheetRowsWithoutRegSrc(pattern, dsscSetupName, ref writeDssc, enumVbtFuncType, patInfo, cppSetupName, cppBitWidth, dsscOutStr, ref postProcessSheetRows, dsscPostProcessNameSheet);
            }

            return CreateRegisterSheetRowsWithRegSrc(pattern, dsscSetupName, ref writeDssc, enumVbtFuncType, patInfo, cppSetupName, cppBitWidth, curPatDigSrcInfo, regSrcList, eqnStr, dsscOutStr, ref postProcessSheetRows, dsscPostProcessNameSheet);
        }

        private static HashSet<string> BuildRegSrcList(HardIpInfo hardIpInfo)
        {
            var regSrcList_orig = hardIpInfo.SendBitName.Split('+').Distinct().ToList().Where(item => !string.IsNullOrEmpty(item)).ToList();
            var regSrcList = new HashSet<string>();
            foreach (string regs in regSrcList_orig)
            {
                regSrcList.Add(regs.Split('[').FirstOrDefault()!);
            }
            return regSrcList;
        }

        private static void GetCaptureInfoFromMiscInfo(HardIpInfo hardIpInfo, out string? cppSetupName, out string? cppBitWidth)
        {
            //Add PostProcess
            #region Get CaptureInfo from MiscInfo
            //var cpp = Regex.Match(string.Join(";", patInfo.MiscInfo), @"(?<cpp>\w+DSSCSetup_Post_Process)", RegexOptions.IgnoreCase).Groups["cpp"].Value;
            cppSetupName = hardIpInfo.MiscInfo.FirstOrDefault(misc => misc.StartsWith("Post_Process"));
            cppBitWidth = hardIpInfo.MiscInfo.FirstOrDefault(misc => misc.StartsWith("DigCapBits"));
            //var cppSet = Cpps.ContainsKey(pattern) ? Cpps[pattern] :null;
            if (!string.IsNullOrEmpty(cppSetupName))
            {
                cppSetupName = cppSetupName.Split(':')[1];
            }

            if (!string.IsNullOrEmpty(cppBitWidth))
            {
                cppBitWidth = cppBitWidth.Split(':')[1];
            }
            #endregion
        }

        private static List<DsscSetupSheetRow> CreateSingleSetupOnlyRow(string pattern, string dsscSetupName, ref bool writeDssc, EnumVbtFuncType enumVbtFuncType)
        {
            string setupName = "";
            if (writeDssc)
            {
                setupName = dsscSetupName;
                writeDssc = false;
            }
            return [new() { Pattern = pattern, DsscSetup = setupName, Type = enumVbtFuncType }];
        }

        private static List<DsscSetupSheetRow> CreateRegisterSheetRowsWithoutRegSrc(string pattern, string dsscSetupName, ref bool writeDssc, EnumVbtFuncType enumVbtFuncType, HardIpInfo hardIpInfo, string? cppSetupName, string? cppBitWidth, string dsscOutStr, ref List<PostProcessSheetRow> postProcessSheetRows, List<string> dsscPostProcessNameSheet)
        {
            var registerSheetRows = new List<DsscSetupSheetRow>();
            var row = new DsscSetupSheetRow { Type = enumVbtFuncType, Pattern = pattern };
            if (!string.IsNullOrEmpty(hardIpInfo.Payload))
            {
                row.Pattern = hardIpInfo.Payload;
            }
            //row.DigSrcEqn = EqnStr;
            row.DigCapPin = !string.IsNullOrEmpty(cppSetupName) ? "JTAG_TDO" : hardIpInfo.CapPinName;
            row.DigCapSampleSize = !string.IsNullOrEmpty(cppBitWidth) ? cppBitWidth : hardIpInfo.CapBit != 0 ? hardIpInfo.CapBit.ToString() : "";

            if (writeDssc)
            {
                row.DsscSetup = dsscSetupName;
                writeDssc = false;
            }

            if (enumVbtFuncType == EnumVbtFuncType.LCD && dsscOutStr.Length > 0)
            {
                string dsscPostProcessName = dsscSetupName + "_PostProcess";
                row.CusStrDigCapData = "POST_PROCESS," + dsscPostProcessName;
                dsscPostProcessNameSheet.Add(dsscPostProcessName);

                List<PostProcessSheetRow> postProcessRows = CreatePostProcessSheetRows(pattern, dsscPostProcessName, dsscOutStr);
                if (postProcessRows != null)
                {
                    postProcessSheetRows.AddRange(postProcessRows);
                }
            }
            else
            {
                row.CusStrDigCapData = !string.IsNullOrEmpty(cppSetupName) ? "POST_PROCESS," + cppSetupName : dsscOutStr.Length >= 6000 ? WriteCusStrDigCapData(ref dsscOutStr) : dsscOutStr;
            }

            registerSheetRows.Add(row);

            //Add excessive row
            while (!row.CusStrDigCapData.StartsWith("POST_PROCESS") && dsscOutStr.Length > 0)
            {
                registerSheetRows.Add(new DsscSetupSheetRow
                {
                    CusStrDigCapData = WriteCusStrDigCapData(ref dsscOutStr)
                });
            }

            return registerSheetRows;
        }

        private static List<DsscSetupSheetRow> CreateRegisterSheetRowsWithRegSrc(string pattern, string dsscSetupName, ref bool writeDssc, EnumVbtFuncType enumVbtFuncType, HardIpInfo hardIpInfo, string? cppSetupName, string? cppBitWidth, Dictionary<string, string> curPatDigSrcInfo, HashSet<string> regSrcList, string eqnStr, string dsscOutStr, ref List<PostProcessSheetRow> postProcessSheetRows, List<string> dsscPostProcessNameSheet)
        {
            var registerSheetRows = new List<DsscSetupSheetRow>();
            bool firstRow = true;
            foreach (string reg in regSrcList)
            {
                var row = new DsscSetupSheetRow
                {
                    Type = enumVbtFuncType
                };
                if (firstRow)
                {
                    row.Pattern = pattern;
                    if (!string.IsNullOrEmpty(hardIpInfo.Payload))
                    {
                        row.Pattern = hardIpInfo.Payload;
                    }
                    //row.DigSrcEqn = EqnStr;
                    row.DigSrcPin = hardIpInfo.SendPinName;
                    row.DigCapPin = !string.IsNullOrEmpty(cppSetupName) ? "JTAG_TDO" : hardIpInfo.CapPinName;
                    row.DigCapSampleSize = !string.IsNullOrEmpty(cppBitWidth) ? cppBitWidth : hardIpInfo.CapBit != 0 ? hardIpInfo.CapBit.ToString() : "";

                    //row.PatModuleInfo = patInfo.VM_Vector;
                    row.PatModuleInfo = "";

                    if (writeDssc)
                    {
                        row.DsscSetup = dsscSetupName;
                        writeDssc = false;
                    }
                }

                row.DigSrcEqn = WriteDigSrcEqn(ref eqnStr);
                row.DigSrcReg = reg;
                row.DigSrcAssignment = curPatDigSrcInfo.TryGetValue(reg, out string? digSrcAssignment) ? digSrcAssignment : "";
                //var dicRegAssignment = CreateRegsisterAssignment(patInfo, CurPatDigSrcInfo);
                //if (dicRegAssignment != null)
                //{
                //    if (dicRegAssignment.ContainsKey(reg))
                //        row.DigSrcAssignment = dicRegAssignment[reg];
                //}

                row.DigSrcSampleSize = GetBitWidth(reg, hardIpInfo);
                if (enumVbtFuncType == EnumVbtFuncType.LCD && dsscOutStr.Length > 0)
                {
                    string dsscPostProcessName = dsscSetupName + "_PostProcess";
                    row.CusStrDigCapData = "POST_PROCESS," + dsscPostProcessName;
                    dsscPostProcessNameSheet.Add(dsscPostProcessName);

                    List<PostProcessSheetRow> postProcessRows = CreatePostProcessSheetRows(pattern, dsscPostProcessName, dsscOutStr);
                    if (postProcessRows != null)
                    {
                        postProcessSheetRows.AddRange(postProcessRows);
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(cppSetupName))
                    {
                        if (firstRow)
                        {
                            row.CusStrDigCapData = "POST_PROCESS," + cppSetupName;
                        }
                    }
                    else
                    {
                        if (dsscOutStr.Length >= 6000)
                        { row.CusStrDigCapData = WriteCusStrDigCapData(ref dsscOutStr); }
                        else
                        {
                            row.CusStrDigCapData = dsscOutStr;
                            dsscOutStr = "";
                        }
                    }
                    //row.CusStrDigCapData = dsscOutStr.Length >= 6000 ? WriteCusStrDigCapData(ref dsscOutStr) : dsscOutStr;
                }

                registerSheetRows.Add(row);
                if (firstRow)
                {
                    firstRow = false;
                }
            }

            //Add excessive row
            while (eqnStr.Length > 0 || (dsscOutStr.Length > 0 && string.IsNullOrEmpty(cppSetupName)))
            {
                registerSheetRows.Add(new DsscSetupSheetRow
                {
                    DigSrcEqn = WriteDigSrcEqn(ref eqnStr),
                    CusStrDigCapData = !string.IsNullOrEmpty(cppSetupName) ? "" : WriteCusStrDigCapData(ref dsscOutStr)
                });
            }
            return registerSheetRows;
        }

        private static List<PostProcessSheetRow> CreatePostProcessSheetRows(string pattern, string dsscPostProcessSetupName, string dsscOut)
        {
            var postProcessRows = new List<PostProcessSheetRow>();

            foreach (string dssc in dsscOut.Split(','))
            {
                if (dssc.EqualsIgnoreCase("dssc_out"))
                {
                    continue;
                }

                var row = new PostProcessSheetRow
                {
                    SetupName = dsscPostProcessSetupName,
                    PatternName = pattern,
                    BlockName = dsscPostProcessSetupName.Split('_')[1],
                };
                string[] info = dssc.Split(':');
                if (info.Length >= 2)
                {
                    row.BitWidth = info[0].Trim();
                    row.TestName = info[1].Trim();
                }
                if (info.Length == 3)
                {
                    row.StoreName = info[2].Trim();
                }
                postProcessRows.Add(row);
            }

            return postProcessRows;
        }

        private static string WriteDigSrcEqn(ref string sendBitName)
        {
            ArgumentNullException.ThrowIfNull(sendBitName);

            string result = sendBitName;
            if (sendBitName.Length > 6000)
            {
                List<string> sendRegs = [.. sendBitName.Split('+')];
                var newSendRegs = new List<string>();
                while (string.Join("+", newSendRegs).Length < 6000 && sendRegs.Count != 0)
                {
                    newSendRegs.Add(sendRegs[0]);
                    sendRegs.RemoveAt(0);
                }
                sendBitName = string.Join("+", sendRegs);
                result = string.Join("+", newSendRegs);
            }
            else
            {
                sendBitName = "";
            }

            return result;
        }

        private static string WriteCusStrDigCapData(ref string dsscOut)
        {
            string result = dsscOut;
            if (dsscOut.Length > 6000)
            {
                List<string> dsscOutRegs = [.. dsscOut.Split(',')];
                var newDsscOutRegs = new List<string>();
                while (string.Join(",", newDsscOutRegs).Length < 6000 && dsscOutRegs.Count != 0)
                {
                    newDsscOutRegs.Add(dsscOutRegs[0]);
                    dsscOutRegs.RemoveAt(0);
                }
                dsscOut = string.Join(",", dsscOutRegs);
                result = string.Join(",", newDsscOutRegs);
            }
            else
            {
                dsscOut = "";
            }

            return result;
        }

        private static string GetBitWidth(string targetReg, HardIpInfo hardIpInfo)
        {
            var digSrcItems = new Dictionary<string, string>();
            string[] sendbits = hardIpInfo.SendBitStr.Split('+');
            string[] regs = hardIpInfo.SendBitName.Split('+');
            int i = 0;
            foreach (string reg in regs)
            {
                if (string.IsNullOrEmpty(reg))
                {
                    continue;
                }
                if (reg.Contains('[') && reg.Contains(']'))
                {
                    if (MyRegex().IsMatch(reg))
                    {
                        string reg_split = reg.Split('[').FirstOrDefault()!;
                        string newbit =
                            MyRegex1().Match(reg).Groups["Partial"].Value;
                        int newbitvalue = 0;
                        if (digSrcItems.TryGetValue(reg_split, out string? value) && int.TryParse(newbit, out newbitvalue) &&
                            int.TryParse(value, out int origbitvalue))
                        {
                            if (newbitvalue + 1 > origbitvalue)
                            {
                                digSrcItems[reg_split] = (newbitvalue + 1).ToString();
                            }

                            continue;
                        }
                        if (!digSrcItems.ContainsKey(reg_split))
                        {
                            digSrcItems.Add(reg_split, (newbitvalue + 1).ToString());
                        }
                    }
                }
                else
                {
                    string bitWidth = sendbits[i].Split('_')[1];
                    digSrcItems.TryAdd(reg, bitWidth);
                    i++;
                }

            }
            return digSrcItems[targetReg];
        }

        private static string AddStoreNameToDsscout(HardIpInfo hardIpInfo)
        {
            if (hardIpInfo.CapBit > 200000)
            {
                return hardIpInfo.DsscOut.ToString();
            }
            string dsscout = hardIpInfo.DsscOut;
            string capStoreName = hardIpInfo.CapStoreName;
            var newDsscOut = new StringBuilder();
            Response.Report(string.Format("AddStoreNameToDsscout ...{0}", hardIpInfo.Payload));
            if (!string.IsNullOrEmpty(dsscout.ToString()) && !string.IsNullOrEmpty(capStoreName))
            {
                Dictionary<string, string> capStoreNameDic = capStoreName.Split(';').ToDictionary(x => x.Split(':')[0]);
                List<string> dsscOutDic = [.. dsscout.Replace("DSSC_OUT,", "").Split(',')];

                if (capStoreNameDic.Count == dsscOutDic.Count)
                {
                    for (int idx = 0; idx < capStoreNameDic.Count; idx++)
                    {
                        string testname = dsscOutDic[idx].Split(':')[1];
                        if (capStoreNameDic.TryGetValue(testname, out string? storeName))
                        {
                            dsscOutDic[idx] = dsscOutDic[idx].Replace(testname, storeName);
                        }
                    }
                }
                return "DSSC_OUT," + string.Join(",", dsscOutDic);
                //foreach (var key in capStoreNameDic.Keys)
                //{
                //    dsscout = dsscout.Replace(key, capStoreNameDic[key]);
                //}
            }

            return dsscout.ToString();
        }
    }
}
