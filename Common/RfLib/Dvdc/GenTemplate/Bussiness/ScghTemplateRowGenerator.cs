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
using LogLib.Utility;

using ProjectConfigLib.ProjectConfig;

using RfLib.Dvdc.GenTemplate.TestPlanFormat;

using ScghLib.Reader;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    internal sealed partial class ScghTemplateRowGenerator(TemplateAutoGen templateAutoGen)
    {
        [GeneratedRegex("IsReadCapTrim", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex("VDIFF|IDIFF", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex5();

        private readonly TemplateAutoGen _owner = templateAutoGen;

        public Dictionary<string, List<TemplateRow>> GenTemplatesByScgh(List<PatternRow> patternRows)
        {
            Response.Report(string.Format("Generate Templates By Scgh ..."), percentage: Convert.ToInt32(75));

            var results = new Dictionary<string, List<TemplateRow>>();
            var patternIndex = new Dictionary<string, int>();
            int i = -1;

            var scghItemsByBlockModeItem =
                ScghStatic.ScghData.ConvertedPatternRowListByHardip.GroupBy(p => p.Block + "#" + p.Mode + "#" + p.Item + "#" + p.PayloadValue)
                    .ToDictionary(p => p.Key, p => p.ToList());
            foreach (KeyValuePair<string, List<ProdCharSheetRow>> scghItems in scghItemsByBlockModeItem)
            {
                ProcessScghItemGroup(scghItems.Value, patternRows, results, patternIndex, ref i);
            }
            return results;
        }

        private void ProcessScghItemGroup(List<ProdCharSheetRow> prodCharSheetRows, List<PatternRow> patternRows, Dictionary<string, List<TemplateRow>> results, Dictionary<string, int> patternIndex, ref int i)
        {
            bool isAllHardIP = prodCharSheetRows.All(p => TemplateAutoGenHelpers1.GetFuncType(p) == EnumVbtFuncType.HardIP);
            string scghPayload = prodCharSheetRows.Select(p => p.PayloadValue).First();
            List<PatternRow> planTargets =
                patternRows.FindAll(
                    p =>
                        p.Pattern.GetPatternName().EqualsIgnoreCase(scghPayload));
            bool isFoundReferenceItems = planTargets.Count > 0;
            int generateLoopCount = isFoundReferenceItems ? planTargets.Count : 1;
            if (generateLoopCount > 1)
            {
                ;
            }
            //per row generate, if need to reference multipl same test plan payload
            //=> still need to follow :  All initExpandSet+payload1 then All initExpandSet+payload2
            // Limitation: payload might suggest need to be unique => it might generate redundant set.
            for (int loopCount = 0; loopCount < generateLoopCount; loopCount++)
            {
                PatternRow? referenceItem = isFoundReferenceItems ? planTargets[loopCount] : null;
                foreach (ProdCharSheetRow scgh in prodCharSheetRows)
                {
                    #region update main flow => to be improve code

                    i++;
                    try
                    {
                        ProcessSingleScghTemplateRow(scgh, referenceItem, results, patternIndex, isAllHardIP);
                    }
                    catch (Exception ex)
                    {
                        ErrorMessageBox.Show(string.Format(ex.ToString()));
                    }

                    #endregion
                }
            }
        }

        private void ProcessSingleScghTemplateRow(ProdCharSheetRow prodCharSheetRow, PatternRow? patternRow, Dictionary<string, List<TemplateRow>> results, Dictionary<string, int> patternIndex, bool isAllHardIP)
        {
            if (!prodCharSheetRow.IsGenFlow || prodCharSheetRow.Block.ContainsIgnoreCase("init"))
            {
                return;
            }

            const string blockPrefix = "HardIP_";
            string block = blockPrefix + prodCharSheetRow.Block.Replace(" ", "_");
            List<string> testNameFromScghBlocks = [.. ProjectConfigSingleton.Instance().GetValue("Template", "TestNameFromScghBlocks").Split(',')];
            bool testNameFromScgh = testNameFromScghBlocks.Any(x => x.Trim().EqualsIgnoreCase(prodCharSheetRow.Block));
            int stepIndex = 0;
            HardIpInfo patInfo = LocalSpecs.HardIpInfos.GetHardIpInfo(prodCharSheetRow.PayloadValue);
            EnumVbtFuncType itemType = TemplateAutoGenHelpers1.GetFuncType(prodCharSheetRow);

            Dictionary<string, Dictionary<string, string>> digSrcInfo = RegisterLimitationGenerator.ProcessDigSrcDictionary(prodCharSheetRow);
            //Determine function type with RFFunc/RFTrim/BBFunc/WiTrim/HardIP

            List<TemplateRow> blockTemplates;
            if (results.TryGetValue(block, out List<TemplateRow>? existingBlockTemplates))
            {
                blockTemplates = existingBlockTemplates;
                patternIndex[block]++;
            }
            else
            {
                blockTemplates = [];
                results.Add(block, blockTemplates);
                patternIndex.Add(block, 1);
            }

            TemplateRow templateRow2 = null!;
            if (patInfo.NewInfo != null || LocalSpecs.Options.Device == EnumDevice.RF || itemType == EnumVbtFuncType.RF ||
                LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                templateRow2 = GenerateLcdRfBlockRows(prodCharSheetRow, patInfo, itemType, isAllHardIP, digSrcInfo, blockTemplates, patternIndex, block, testNameFromScgh, ref stepIndex);
            }
            else // HardIP
            //Describtion row for block+item+mode
            {
                templateRow2 = GenerateHardIpBlockRows(prodCharSheetRow, patInfo, patternRow, itemType, isAllHardIP, digSrcInfo, blockTemplates, patternIndex, block, ref stepIndex);
            }

            GenerateSetEfuseRowIfNeeded(prodCharSheetRow, patInfo, blockTemplates, patternIndex, block, templateRow2, ref stepIndex);
        }

        #region Update LCD/RF items
        private TemplateRow GenerateLcdRfBlockRows(ProdCharSheetRow prodCharSheetRow, HardIpInfo hardIpInfo, EnumVbtFuncType enumVbtFuncType, bool isAllHardIP, Dictionary<string, Dictionary<string, string>> digSrcInfo, List<TemplateRow> templateRows, Dictionary<string, int> patternIndex, string block, bool testNameFromScgh, ref int stepIndex)
        {
            var templateRow1 = new WirelessTemplateRow(patternIndex[block],
                patternIndex[block] + "." + stepIndex)
            {
                Description = prodCharSheetRow.Block + "," + prodCharSheetRow.Item + "," + prodCharSheetRow.Mode
            };
            if (hardIpInfo.NewInfo == null)
            {
                if (hardIpInfo.MeasSeqStr.Length != 0)
                {
                    templateRow1.Description += "  MeasSeq: " + hardIpInfo.MeasSeqStr;
                }
            }
            else if (hardIpInfo.NewInfo.MeasSeq.Length != 0)
            {
                templateRow1.Description += "  MeasSeq: " + hardIpInfo.NewInfo.MeasSeq;
            }

            templateRow1.TestItem = patternIndex[block];
            templateRow1.Step = patternIndex[block] + "." + stepIndex;
            templateRows.Add(templateRow1);
            stepIndex++;

            #region PatternRow

            TemplateRow templateRow2 = new WirelessTemplateRow(patternIndex[block],
                patternIndex[block] + "." + stepIndex)
            {
                TestItem = patternIndex[block],
                Step = patternIndex[block] + "." + stepIndex
            };

            GenerateCommonPatternRow(templateRow2, prodCharSheetRow, hardIpInfo, patternIndex[block], stepIndex, enumVbtFuncType, isAllHardIP, digSrcInfo);
            _owner.PatternRowModifierInternal.ModifyPatternRow((WirelessTemplateRow)templateRow2, prodCharSheetRow, hardIpInfo, enumVbtFuncType, isAllHardIP, testNameFromScgh, _owner.DsscSetupNameTestPlanInternal);
            string? miscBCCF = hardIpInfo.MiscInfo.FirstOrDefault(misc => misc.StartsWith("BestCodeCalcFunc"));
            if (miscBCCF != null)
            {
                hardIpInfo.BestCodeCalcFunc = miscBCCF.Split(';')[0].Split(':')[1];
            }
            templateRows.Add(templateRow2);

            #endregion

            bool isReadCapTrim = hardIpInfo.MiscInfo.Exists(p => MyRegex2().IsMatch(p));

            if (enumVbtFuncType == EnumVbtFuncType.WiTrim || enumVbtFuncType == EnumVbtFuncType.RFTrim || isReadCapTrim)
            //If contains Trim Target => Generate Trim Items
            {
                stepIndex++;
                _owner.WitrimBestCodeGeneratorInternal.GenPlanWithWiTrimItem(templateRows, hardIpInfo, isReadCapTrim, patternIndex[block], ref stepIndex);
            }

            else if (enumVbtFuncType == EnumVbtFuncType.LCD)
            {
                stepIndex++;
                GenerateRFNonPatternRow(hardIpInfo, templateRows, patternIndex, isReadCapTrim, block, ref stepIndex);
            }
            else //if (itemType == VBTFuncType.RF)
            // If contains RF Items => Generate RFItems its way. The RF Items way would be HardIP-Like
            {
                stepIndex++;
                GenerateRFNonPatternRow(hardIpInfo, templateRows, patternIndex, isReadCapTrim, block, ref stepIndex);
            }
            //else // otherwise, choose HardIP way to generate
            //{
            //    //templateRow2.RegisterAssignment = GenerateRegsisterAssignment(patInfo);
            //    stepIndex++;
            //    _GenerateNonPatternRow(patInfo, blockTemplates, patternIndex, block, ref stepIndex);
            //}
            return templateRow2;
        }
        #endregion

        #region update pure HardIP(RF/LCD) or AP items
        private TemplateRow GenerateHardIpBlockRows(ProdCharSheetRow prodCharSheetRow, HardIpInfo hardIpInfo, PatternRow? patternRow, EnumVbtFuncType enumVbtFuncType, bool isAllHardIP, Dictionary<string, Dictionary<string, string>> digSrcInfo, List<TemplateRow> templateRows, Dictionary<string, int> patternIndex, string block, ref int stepIndex)
        {
            bool isAP = LocalSpecs.Options.Device == EnumDevice.AP;
            bool isSplitWithInitPayload = isAP;
            if (isSplitWithInitPayload && prodCharSheetRow.InitList.Count > 0)
            {
                GenerateHardIpInitPartRows(prodCharSheetRow, templateRows, patternIndex, block, ref stepIndex);
            }
            //else
            TemplateRow templateRow2 = GenerateHardIpPayloadRows(prodCharSheetRow, hardIpInfo, patternRow, isSplitWithInitPayload, enumVbtFuncType, isAllHardIP, digSrcInfo, patternIndex, block, templateRows, ref stepIndex);

            return templateRow2;
        }

        private void GenerateHardIpInitPartRows(ProdCharSheetRow prodCharSheetRow, List<TemplateRow> templateRows, Dictionary<string, int> patternIndex, string block, ref int stepIndex)
        {
            #region generate Init part

            var templateRow1 = new TemplateRow(patternIndex[block],
                patternIndex[block] + "." + stepIndex);
            templateRows.Add(templateRow1);
            templateRow1.Description = prodCharSheetRow.Block + "," + prodCharSheetRow.Item + "," + prodCharSheetRow.Mode + "init";

            stepIndex++;
            //Pattern row

            TemplateRow templateRow2 = new TemplateRow(patternIndex[block],
                patternIndex[block] + "." + stepIndex)
            {
                TTR = TemplateAutoGenHelpers.GetEnableHLNInfo(prodCharSheetRow.EnableHlnv)
            };
            string initSubBlock = string.Format("SubBlock:Init-{0}-{1}", CommonGenerator.GetSubBlockNameByPattern(prodCharSheetRow.PayloadValue, prodCharSheetRow.Item, false), CommonGenerator.GetInitSubBlockName(prodCharSheetRow.InitList));
            #region deal with Misc info Check whether Pattern type mismatch/ TTR remove

            var miscinfoList = new List<string>
            {
                initSubBlock
            };
            #region check TTR
            if (!string.IsNullOrEmpty(templateRow2.TTR))
            {//remove TTR specified items in test flow
                foreach (string ttrVoltage in templateRow2.TTR.Split(','))
                {
                    miscinfoList.Add(string.Format("Remove{0}", ttrVoltage));
                }
            }
            #endregion
            #region Check Pattern type between init and payload
            //if not match => put payload enableword on init
            if (
                !prodCharSheetRow.InitList.Last().Split('_')[0].EqualsIgnoreCase(prodCharSheetRow.PayloadValue.Split('_')[0]))
            {
                var hLNList = new List<string> { "HV", "LV", "NV" };
                foreach (string voltage in hLNList)
                {
                    if (templateRow2.TTR.Contains(voltage))
                    {
                        continue;
                    }

                    string enableWord = CommonGenerator.GenEnableWord(prodCharSheetRow.PayloadValue, "", voltage, _owner.HardIpInputDataInternal);
                    miscinfoList.Add(string.Format("EnableWord:{0}", enableWord));
                }
            }

            #endregion

            templateRow2.MiscInfo = string.Join(";", miscinfoList);
            #endregion
            templateRows.Add(templateRow2);
            //List<string> nameList = ProdCharSheetRowList.Select(x => x.Mode.ToLower()).Distinct().ToList();
            //templateRow2.Pattern = string.Join(@",",GetPatternName(scgh.Key, nameList, ProdCharSheetRowList));

            templateRow2.ForceCondition = RegisterLimitationGenerator.GetDCSpec(prodCharSheetRow.DcUsed, prodCharSheetRow.InitList);

            templateRow2.Pattern =
                string.Join(",", prodCharSheetRow.InitList);

            templateRow2.Description = "Run the pattern provided";

            stepIndex++;

            #endregion
        }

        private TemplateRow GenerateHardIpPayloadRows(ProdCharSheetRow prodCharSheetRow, HardIpInfo hardIpInfo, PatternRow? patternRow, bool isSplitWithInitPayload, EnumVbtFuncType enumVbtFuncType, bool isAllHardIP, Dictionary<string, Dictionary<string, string>> digSrcInfo, Dictionary<string, int> patternIndex, string block, List<TemplateRow> templateRows, ref int stepIndex)
        {
            #region Convert Payload

            TemplateRow templateRow2;
            if (patternRow != null)
            {
                templateRow2 = TemplateAutoGenHelpers.GenerateHardIpReferencedPayloadRows(prodCharSheetRow, patternRow, patternIndex, block, templateRows, ref stepIndex);
            }
            else
            {
                templateRow2 = GenerateHardIpOriginPayloadRows(prodCharSheetRow, hardIpInfo, digSrcInfo, isSplitWithInitPayload, enumVbtFuncType, isAllHardIP, patternIndex, block, templateRows, ref stepIndex);
            }

            #endregion

            return templateRow2;
        }

        private TemplateRow GenerateHardIpOriginPayloadRows(ProdCharSheetRow prodCharSheetRow, HardIpInfo hardIpInfo, Dictionary<string, Dictionary<string, string>> digSrcInfo, bool isSplitWithInitPayload, EnumVbtFuncType enumVbtFuncType, bool isAllHardIP, Dictionary<string, int> patternIndex, string block, List<TemplateRow> templateRows, ref int stepIndex)
        {
            #region Origin path => based on patinfo to generate payload
            var templateRow1 = new TemplateRow(patternIndex[block],
                patternIndex[block] + "." + stepIndex);
            templateRows.Add(templateRow1);
            templateRow1.Description = prodCharSheetRow.Block + "," + prodCharSheetRow.Item + "," + prodCharSheetRow.Mode;
            if (hardIpInfo.MeasSeqStr.Length != 0)
            {
                templateRow1.Description += "  MeasSeq: " + hardIpInfo.MeasSeqStr;
            }

            stepIndex++;
            //Pattern row

            TemplateRow templateRow2 = new TemplateRow(patternIndex[block],
                patternIndex[block] + "." + stepIndex);
            if (
                !string.IsNullOrEmpty(CommonGenerator.GetSubBlockNameByPattern(prodCharSheetRow.PayloadValue, prodCharSheetRow.Item, false)))
            {
                templateRow2.MiscInfo += string.Format("SubBlock:{0}-{1};", CommonGenerator.GetSubBlockNameByPattern(prodCharSheetRow.PayloadValue, prodCharSheetRow.Item, false), CommonGenerator.GetInitSubBlockName(prodCharSheetRow.InitList));
            }

            templateRows.Add(templateRow2);
            //List<string> nameList = ProdCharSheetRowList.Select(x => x.Mode.ToLower()).Distinct().ToList();
            //templateRow2.Pattern = string.Join(@",",GetPatternName(scgh.Key, nameList, ProdCharSheetRowList));
            templateRow2.ForceCondition = _owner.PlanGeneratorInternal.UpdateForceConditionNew(hardIpInfo.ForceCondition);
            templateRow2.TTR = TemplateAutoGenHelpers.GetEnableHLNInfo(prodCharSheetRow.EnableHlnv);

            if (digSrcInfo.TryGetValue(prodCharSheetRow.PayloadValue, out Dictionary<string, string>? payloadDigSrcInfo))
            {
                templateRow2.RegisterAssignment = TemplateAutoGenHelpers.GenerateRegsisterAssignment(hardIpInfo, payloadDigSrcInfo);
            }

            templateRow2.Description = "Run the pattern provided";

            GenerateCommonPatternRow(templateRow2, prodCharSheetRow, hardIpInfo, patternIndex[block], stepIndex, enumVbtFuncType, isAllHardIP, digSrcInfo);

            if (!string.IsNullOrEmpty(templateRow2.TTR))
            {//remove TTR specified items in test flow
                var miscinfoList = new List<string>();
                foreach (string ttrVoltage in templateRow2.TTR.Split(','))
                {
                    miscinfoList.Add(string.Format("Remove{0}", ttrVoltage));
                }
                templateRow2.MiscInfo = string.Format("{0};{1}", templateRow2.MiscInfo, string.Join(";", miscinfoList));
            }

            templateRow2.Pattern = isSplitWithInitPayload && prodCharSheetRow.InitList.Count > 0
                ? string.Join(",", prodCharSheetRow.PayloadList).Trim(',')
                : (string.Join(",", prodCharSheetRow.InitList) + "," +
                     string.Join(",", prodCharSheetRow.PayloadList)).Trim(',');
            stepIndex++;
            GenerateNonPatternRow(hardIpInfo, templateRows, patternIndex, block, ref stepIndex);
            #endregion

            return templateRow2;
        }
        #endregion

        private void GenerateSetEfuseRowIfNeeded(ProdCharSheetRow prodCharSheetRow, HardIpInfo hardIpInfo, List<TemplateRow> templateRows, Dictionary<string, int> patternIndex, string block, TemplateRow templateRow, ref int stepIndex)
        {
            if (hardIpInfo.MiscInfo.Count > 0 && hardIpInfo.MiscInfo.Any(info => info.StartsWithIgnoreCase("SetEfuse")))
            {
                stepIndex++;
                var templateRowMiscInfo = new WirelessTemplateRow(patternIndex[block],
                    patternIndex[block] + "." + stepIndex);
                string fuseSetupName = prodCharSheetRow.Block + "_" +
CommonGenerator.GetSubBlockNameByPattern(prodCharSheetRow.PayloadValue, prodCharSheetRow.Item, false).Replace("_", "");
                GenerateMiscInfoRow(templateRowMiscInfo, fuseSetupName, hardIpInfo.MiscInfo, templateRow);
                templateRows.Add(templateRowMiscInfo);
            }
        }

        private void GenerateCommonPatternRow(TemplateRow templateRow, ProdCharSheetRow prodCharSheetRow, HardIpInfo hardIpInfo, int stepIndex, int stepSubIndex, EnumVbtFuncType enumVbtFuncType, bool isHardIPSheet, Dictionary<string, Dictionary<string, string>> digSrcInfo)
        {
            templateRow.MiscInfo = TemplateAutoGenHelpers.SelectVBTFuncName(hardIpInfo);
            if (hardIpInfo.MiscInfo.Count > 0 && !hardIpInfo.MiscInfo.All(info => info.StartsWith("SetEfuse:")))
            {
                templateRow.MiscInfo = templateRow.MiscInfo + string.Join(";", hardIpInfo.MiscInfo).Trim(';') + ";";
            }
            if (!string.IsNullOrEmpty(CommonGenerator.GetSubBlockNameByPattern(prodCharSheetRow.PayloadValue, prodCharSheetRow.Item, false)) && LocalSpecs.Options.Device != EnumDevice.RF && isHardIPSheet)
            {
                templateRow.MiscInfo += string.Format("SubBlock:{0};", CommonGenerator.GetSubBlockNameByPattern(prodCharSheetRow.PayloadValue, prodCharSheetRow.Item, false));
            }

            if (enumVbtFuncType == EnumVbtFuncType.RFTrim || enumVbtFuncType == EnumVbtFuncType.WiTrim)
            {
                if (hardIpInfo.NewInfo != null && !string.IsNullOrEmpty(hardIpInfo.NewInfo.MeasName))
                {
                    templateRow.MiscInfo += string.Format("MeasName:{0};", hardIpInfo.NewInfo.MeasName);
                }
            }
            if (!string.IsNullOrEmpty(prodCharSheetRow.Item) && _owner.RelayInfoInternal.RelayItems.Any(p => p.Item.EqualsIgnoreCase(prodCharSheetRow.Item)))
            {
                templateRow.MiscInfo += TemplateAutoGenHelpers.GetRelayOnMiscInfo(_owner.RelayInfoInternal.RelayItems.FirstOrDefault(p => p.Item.EqualsIgnoreCase(prodCharSheetRow.Item))!);
            }

            #region Force Condition
            if (hardIpInfo.ForceCondition != null)
            {
                var sweepList = new List<string>();
                var allSweepInPat = new List<ForcePin>();
                var allSweepInSeq = new List<ForcePin>();
                if (hardIpInfo.ForceCondition.ForcePins.Exists(p => p.ForceType.ContainsIgnoreCase("sweep")))
                {
                    allSweepInPat =
                        hardIpInfo.ForceCondition.ForcePins.FindAll(p => p.ForceType.ContainsIgnoreCase("sweep"));
                }
                //Not PMIC or not Trim would collect sweep information in misc info
                if (hardIpInfo.NewInfo != null && !(enumVbtFuncType == EnumVbtFuncType.WiTrim || LocalSpecs.Options.Device == EnumDevice.LCD))
                {
                    allSweepInSeq = hardIpInfo.NewInfo.SeqInfo.SelectMany(p => p.ForceConditions.ForcePins)
                        .ToList()
                        .FindAll(p => p.ForceType.ContainsIgnoreCase("sweep"));
                    allSweepInPat.AddRange(allSweepInSeq);
                }
                if (allSweepInPat.Count > 0)
                {
                    foreach (ForcePin sweepItem in allSweepInPat)
                    {
                        string sweepStr = string.Format(@"sweep({0}:{1});", sweepItem.PinName, sweepItem.ForceValue);
                        sweepList.Add(sweepStr);
                    }
                    templateRow.MiscInfo += string.Join("\n", sweepList);
                }
                //fill PrePatForce todo
                templateRow.ForceCondition = _owner.PlanGeneratorInternal.UpdateForceConditionNew(hardIpInfo.ForceCondition);
            }
            #endregion


            templateRow.Pattern = TemplateAutoGenHelpers.SelectPatternNameByType(prodCharSheetRow, enumVbtFuncType);
            templateRow.TestName = TemplateAutoGenHelpers.SelectTestNameByType(prodCharSheetRow, hardIpInfo);
            Dictionary<string, string> patDigSRcInfo = digSrcInfo.TryGetValue(prodCharSheetRow.PayloadValue, out Dictionary<string, string>? scghPayloadDigSrcInfo)
                ? scghPayloadDigSrcInfo
                : [];
            templateRow.RegisterAssignment = TemplateAutoGenHelpers.GenerateRegsisterAssignment(hardIpInfo, patDigSRcInfo);
            templateRow.RegisterAssignment = _owner.RegisterLimitationGeneratorInternal.DetectRegisterAssignLimitation(prodCharSheetRow, templateRow.RegisterAssignment);
            templateRow.MiscInfo += _owner.RegisterLimitationGeneratorInternal.DetectDigSrcEqnLimitation(prodCharSheetRow, hardIpInfo.SendBitName);
            templateRow.Description = "Run the pattern provided";
            templateRow.Step = stepIndex + "." + stepSubIndex;

        }

        private void GenerateMiscInfoRow(TemplateRow template, string name, List<string> miscinfo, TemplateRow patternRow)
        {
            template.Description = "Set Efuse Write " + name;
            template.Pattern = "Instance:Wireless_HIP_eFuse_Write_" + name;
            var item = new SetEfuseItem();
            string vbtName = "Wireless_HIP_eFuse_Write";
            //Collect setEfuse information
            foreach (string miscstring in miscinfo)
            {
                foreach (string permisc in miscstring.Replace("\"", "").Split(';'))
                {
                    string regfuse = @"SetEfuse\s*:\s*(?<Fuse>.*)";
                    string regfuseOther = @"(?<vbt>SetEfuse\w+)\s*:\s*(?<Fuse>.*)";
                    if (Regex.IsMatch(permisc, regfuse, RegexOptions.IgnoreCase))
                    {
                        string fuseStr = Regex.Match(permisc, regfuse, RegexOptions.IgnoreCase).Groups["Fuse"].Value;
                        foreach (string fuseHIP in fuseStr.Replace(",", "@").Split('@'))
                        {
                            string regStorefuse = @"(?<HIPCodeStore>.*)\s*=\s*(?<efuseField>.*)";
                            var fuseinfo = new EFuseStoreRow();
                            Match match = Regex.Match(fuseHIP, regStorefuse, RegexOptions.IgnoreCase);
                            fuseinfo.FieldName = match.Groups["efuseField"].ToString().Trim();
                            fuseinfo.CaptureStoreName = match.Groups["HIPCodeStore"].ToString().Trim();
                            fuseinfo.FuseEnable = "TRUE";
                            item.Contents.Add(fuseinfo);
                        }
                        patternRow.MiscInfo = patternRow.MiscInfo.Replace(permisc, "").Trim(';');
                    }
                    else if (Regex.IsMatch(permisc, regfuseOther, RegexOptions.IgnoreCase))
                    {//MiscInfo=SetEFuseFromLookUpTable:ARF_PADIO_DCTPP_5G(ValA)=tx_5g_ppa_casc_bias_tune_1p0
                        //captureStoredName	efuseNames
                        Match matchFirst = Regex.Match(permisc, regfuseOther, RegexOptions.IgnoreCase);
                        vbtName = matchFirst.Groups["vbt"].Value;
                        string regStorefuse = @"(?<HIPCodeStore>.*)\s*=\s*(?<efuseField>.*)";
                        var fuseinfo = new EFuseStoreRow();
                        Match match = Regex.Match(matchFirst.Groups["Fuse"].Value, regStorefuse, RegexOptions.IgnoreCase);
                        fuseinfo.FieldName = match.Groups["efuseField"].ToString().Trim();
                        fuseinfo.CaptureStoreName = match.Groups["HIPCodeStore"].ToString().Trim();
                        fuseinfo.FuseEnable = "TRUE";
                        item.Contents.Add(fuseinfo);
                        patternRow.MiscInfo = patternRow.MiscInfo.Replace(permisc, "").Trim(';');
                    }
                }

            }
            //CapFuseSetupName
            var miscList = new List<string>();
            if (item.Contents.Count > 0)
            {
                miscList.Add("Func:" + vbtName);
                if (item.Contents.Count >= 5)
                {
                    _owner.ReferenceSetEfuseItems.Add(item);
                    item.ItemName = string.Format("SetEfuse_{0}", name);
                    miscList.Add(string.Format("CapFuseSetupName:{0}", item.ItemName));
                }
                else
                {
                    if (vbtName == "SetEFuseFromLookUpTable")
                    {
                        miscList.Add(string.Format("efuseNames:{0}", string.Join(",", item.Contents.Select(p => p.FieldName))));
                        miscList.Add(string.Format("captureStoredName:{0}", string.Join(",", item.Contents.Select(p => p.CaptureStoreName))));
                    }
                    else
                    {
                        miscList.Add(string.Format("efuseFieldName:{0}", string.Join(",", item.Contents.Select(p => p.FieldName))));
                        miscList.Add(string.Format("HIPCodeStoreName:{0}", string.Join(",", item.Contents.Select(p => p.CaptureStoreName))));
                    }
                }
            }

            template.MiscInfo = string.Join(";", miscList);
        }

        private void GenerateRFNonPatternRow(HardIpInfo hardIpInfo, List<TemplateRow> templateRows, Dictionary<string, int> patternIndex, bool isReadCap, string block, ref int stepIndex)
        {
            /*
             * New Seq: ,,,F,I,V,Vdiff
New Seq MeasPin: |||PAD_CLKREF_0|PAD_DCTPN|DAC_ANATESTP::DAC_ANATESTN|
New Seq MeasName: |||220M_OUT|IPP|VREF|VDIFF
New Seq ForceType: I|V|V,SweepV|
New Seq ForcePin: ANALOG_TEST_P|ANALOG_TEST_N|PAD_ASG_TXN_0,PAD_ASG_TXN_1,PAD_ASG_TXN_2,PAD_ASG_TXN_3,PAD_ASG_TXN_4,PAD_ASG_TXN_5,PAD_ASG_TXN_6,PAD_ASG_TXN_7||
New Seq SweepPin: ||PAD_ASG_TXP_0,PAD_ASG_TXP_1,PAD_ASG_TXP_2,PAD_ASG_TXP_3,PAD_ASG_TXP_4,PAD_ASG_TXP_5,PAD_ASG_TXP_6,PAD_ASG_TXP_7||
New Seq ForceValue: 0.000025|0.5|0.4||
New Seq SweepValue: ||0.25,0.55,0.1|||
New Seq ExpectValue: |||2.2GHz|0.000025|0.6|0.5|
New Seq Hlimit:|||||0.7||
New Seq Llimit:|||||0.5||
             */
            int i = 0;

            var template = new WirelessTemplateRow(patternIndex[block], patternIndex[block] + "." + stepIndex)
            {
                TestItem = patternIndex[block],
                Step = patternIndex[block] + "." + stepIndex
            };
            try
            {
                if (hardIpInfo.NewInfo != null)
                {
                    bool isRFItem = hardIpInfo.NewInfo.SeqInfo.SelectMany(p => p.MeasPins).ToList()
                        .Exists(p => TemplateAutoGen.MyRegex3().IsMatch(p.MeasType)) ||
                        hardIpInfo.MiscInfo.Exists(p => TemplateAutoGen.MyRegex4().IsMatch(p));
                    foreach (HardIpSeqInfoNew seq in hardIpInfo.NewInfo.SeqInfo)
                    {
                        switch (seq.MeasSeq.ToUpper())
                        {
                            case "VDIFF":
                            case "IDIFF":
                                if (seq.MeasPin.Any(p => p.Contains("::")))
                                {
                                    _owner.PlanGeneratorInternal.GenPlanWithVdiffNew(templateRows, seq, patternIndex[block], stepIndex);
                                }
                                else
                                {
                                    string outString = string.Format("The pin name syntax of VDIFF/IDIFF is wrong");
                                    Response.Report(outString, EnumMessageLevel.Error, Convert.ToInt32(100));
                                }
                                break;
                            case "MEASWAIT":
                                template = new WirelessTemplateRow(patternIndex[block],
                                    patternIndex[block] + "." + stepIndex)
                                {
                                    TestItem = patternIndex[block],
                                    Step = patternIndex[block] + "." + stepIndex,
                                    Description = string.Format("Measure Wait Time : {0}", seq.MeasWait),
                                    Meas = string.Format("MeasWait {0}", seq.MeasWait)
                                };
                                templateRows.Add(template);
                                break;
                            default:
                                _owner.GenPlanWithNonVdiffNew(hardIpInfo.Payload, templateRows, seq, i.ToString(), patternIndex[block], stepIndex, isRFItem);
                                TemplateAutoGenHelpers1.GenerateCalcEquation(templateRows, seq, i.ToString(), patternIndex[block], stepIndex);
                                break;
                        }

                        i++;
                        stepIndex++;
                    }
                }

                if (hardIpInfo.DsscOut.Length != 0 || _owner.Cpps != null)
                {
                    if (_owner.Cpps.ContainsKey(hardIpInfo.Payload.ToUpper()))
                    {
                        MeasCTemplateItemGenerator.GenMeasCTemplateItemCpp(_owner.Cpps[hardIpInfo.Payload.ToUpper()], templateRows, patternIndex[block], stepIndex);
                    }
                    else
                    {
                        MeasCTemplateItemGenerator.GenMeasCTemplateItemNew(hardIpInfo, templateRows, patternIndex[block], stepIndex);
                    }
                }
                if (!string.IsNullOrEmpty(hardIpInfo.TrimTarget.Replace(">", "").Replace("|", "")))
                {
                    List<int> trimBits = !string.IsNullOrEmpty(hardIpInfo.TrimFuseName) ? TemplateAutoGenHelpers1.GetTrimRegBit(hardIpInfo) : [];
                    var newTempRow = new WirelessTemplateRow(patternIndex[block], patternIndex[block] + "." + stepIndex);
                    if (trimBits.Count > 0)
                    {
                        foreach (int bit in trimBits)
                        {
                            newTempRow = new WirelessTemplateRow(patternIndex[block], patternIndex[block] + "." + stepIndex);
                            templateRows.Add(newTempRow);

                            newTempRow.Description = string.Format("TestName for {0}", "BestCode");
                            newTempRow.TestName = "BSTC";
                            newTempRow.LoLimit.Add("CP1", "0");
                            newTempRow.HiLimit.Add("CP1", (Math.Pow(2, bit) - 1).ToString());
                        }
                    }

                    newTempRow.Step = patternIndex[block] + "." + stepIndex;
                    newTempRow = new WirelessTemplateRow(patternIndex[block], patternIndex[block] + "." + stepIndex);
                    templateRows.Add(newTempRow);
                    //

                    newTempRow.Description = string.Format("TestName for {0}", "BestValue");
                    newTempRow.TestName = "BSTV";
                    TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo!, isReadCap);

                    newTempRow.Step = patternIndex[block] + "." + stepIndex;
                    newTempRow = new WirelessTemplateRow(patternIndex[block], patternIndex[block] + "." + stepIndex);
                    if (trimBits.Count > 0)
                    {
                        foreach (int bit in trimBits)
                        {
                            newTempRow = new WirelessTemplateRow(patternIndex[block], patternIndex[block] + "." + stepIndex);
                            templateRows.Add(newTempRow);

                            newTempRow.Description = string.Format("TestName for {0}", "BestCode");
                            newTempRow.TestName = "BSTC";
                            newTempRow.LoLimit.Add("CP1", "0");
                            newTempRow.HiLimit.Add("CP1", (Math.Pow(2, bit) - 1).ToString());
                        }
                    }

                    newTempRow.Step = patternIndex[block] + "." + stepIndex;

                    //newTempRow.HiLimit = highlimit;
                    //newTempRow.LoLimit = lowlimit;

                    newTempRow = new WirelessTemplateRow(patternIndex[block], patternIndex[block] + "." + stepIndex);
                    templateRows.Add(newTempRow);
                    //
                    newTempRow.Description = string.Format("TestName for {0}", "VerificationValue");
                    newTempRow.TestName = "VRFV";
                    TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo!, isReadCap);
                    newTempRow.Step = patternIndex[block] + "." + stepIndex;
                }

                MeasCTemplateItemGenerator.GenExtraLimit(hardIpInfo, templateRows, patternIndex[block]);
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        private void GenerateNonPatternRow(HardIpInfo hardIpInfo, List<TemplateRow> templateRows, Dictionary<string, int> patternIndex, string block, ref int stepIndex)
        {
            foreach (HardIpSeqInfo seqinfo in hardIpInfo.SeqInfo)
            {
                if (MyRegex5().IsMatch(seqinfo.SeqName))
                {
                    if (seqinfo.PinList.Contains("::"))
                    {
                        _owner.PlanGeneratorInternal.GenPlanWithVdiff(templateRows, seqinfo, patternIndex[block], stepIndex);
                    }
                    else
                    {
                        string outString = string.Format("The name of VDIFF {0} is wrong", seqinfo.PinList);
                        Response.Report(outString, EnumMessageLevel.Error, Convert.ToInt32(100));
                    }
                }
                else
                {
                    _owner.PlanGeneratorInternal.GenPlanWithNonVdiff(templateRows, seqinfo, patternIndex[block], stepIndex);
                }

                stepIndex++;
            }

            if (hardIpInfo.DsscOut.Length != 0 && LocalSpecs.Options.Device != EnumDevice.LCD)
            {
                MeasCTemplateItemGenerator.GenMeasCTemplateItem(hardIpInfo, templateRows, patternIndex[block], stepIndex);
            }
            MeasCTemplateItemGenerator.GenExtraLimit(hardIpInfo, templateRows, patternIndex[block]);
        }
    }
}
