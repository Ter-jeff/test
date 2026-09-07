using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Static;
using LogLib.Utility;

using ProjectConfigLib.ProjectConfig;

using RfLib.Dvdc.Reader.CapturePostProcess;
using RfLib.Dvdc.Reader.DsscSetup;

using ScghLib.Reader;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    internal class DsscRegisterTableBuilder(TemplateAutoGen templateAutoGen)
    {
        private readonly TemplateAutoGen _owner = templateAutoGen;
        private readonly List<string> _dSSCSetupName_Sheet = [];
        private readonly List<string> _dSSCPostProcessName_Sheet = [];

        internal DataTable? CreateInstrumentSetup()
        {
            Response.Report(string.Format("Create Instrument Setup ..."), percentage: Convert.ToInt32(70));

            if (SettingStatic.BasicConfigWorkbook == null || SettingStatic.BasicConfigWorkbook.Worksheets["InstrumentSetup"] == null)
            {
                return null;
            }
            if (ScghStatic.ScghData.ConvertedPatternRowListByHardip.All(p => TemplateAutoGenHelpers1.GetFuncType(p) != EnumVbtFuncType.RF &&
TemplateAutoGenHelpers1.GetFuncType(p) != EnumVbtFuncType.RFTrim && TemplateAutoGenHelpers1.GetFuncType(p) != EnumVbtFuncType.LCD)
                )
            {
                return null;
            }
            var genInstrumentSetupTable = new GenInstrumentSetupTable();
            genInstrumentSetupTable.Init(_owner.MultiChannelMapInternal);
            //var setupIndex = 0;

            foreach (ProdCharSheetRow scgh in ScghStatic.ScghData.ConvertedPatternRowListByHardip)
            {
                HardIpInfo patInfo = LocalSpecs.HardIpInfos.GetHardIpInfo(scgh.PayloadValue);
                EnumVbtFuncType funcType = TemplateAutoGenHelpers1.GetFuncType(scgh);
                if (funcType != EnumVbtFuncType.RF && funcType != EnumVbtFuncType.RFTrim && funcType != EnumVbtFuncType.LCD)
                {
                    continue;
                }

                if (patInfo.NewInfo != null)
                {
                    if (!string.IsNullOrEmpty(patInfo.NewInfo.RfSetup))
                    {
                        //setupIndex++;
                        //var patternName = "";
                        //if(scgh.GetInitList().Count > 0)
                        //    patternName += string.Join(",",scgh.GetInitList()) + ",";
                        //patternName += string.Join(",",scgh.GetPayloadList());

                        genInstrumentSetupTable.WriteInstTableNew(patInfo.NewInfo, scgh.PayloadValue, scgh.Item);
                    }
                    else
                    {
                        patInfo.NewInfo.RfSetup = null;
                        var setup = new List<string>();
                        foreach (HardIpSeqInfoNew seq in patInfo.NewInfo.SeqInfo)
                        {
                            setup.Add("[]");
                        }
                        patInfo.NewInfo.RfSetup = string.Join("|", setup);
                        genInstrumentSetupTable.WriteInstTableNew(patInfo.NewInfo, scgh.PayloadValue, scgh.Item);
                    }
                }
            }

            _owner.InstrumentSetupForPatList = genInstrumentSetupTable.InstrumentSetupRowList;

            return genInstrumentSetupTable.DtInstrumentSetup;
        }

        internal Dictionary<EnumVbtFuncType, DsscSetupSheet> CreateRegisterTable()
        {
            Response.Report(string.Format("Create Register Table ..."), percentage: Convert.ToInt32(65));

            Dictionary<EnumVbtFuncType, DsscSetupSheet> registerSheets = [];
            List<string> setupNameList = [];
            List<string> testNameFromScghBlocks = [.. ProjectConfigSingleton.Instance().GetValue("Template", "TestNameFromScghBlocks").Split(',')];
            foreach (ProdCharSheetRow scgh in ScghStatic.ScghData.ConvertedPatternRowListByHardip)
            {
                bool testNameFromScgh = testNameFromScghBlocks.Any(x => x.Trim().EqualsIgnoreCase(scgh.Block));
                DsscSetupSheet registerSheet = null!;
                EnumVbtFuncType funcType = TemplateAutoGenHelpers1.GetFuncType(scgh);
                if (!scgh.IsGenFlow)
                {
                    continue;
                }

                if (scgh.Block.EqualsIgnoreCase("OTP"))
                {
                    funcType = EnumVbtFuncType.OTP;
                }
                //if (scgh.Block.Equals("srcdrv", StringComparison.CurrentCultureIgnoreCase))
                //    funcType = VBTFuncType.SRCDRV;

                if (LocalSpecs.Options.Device == EnumDevice.AP && !(funcType == EnumVbtFuncType.RFTrim || funcType == EnumVbtFuncType.RF))
                {
                    continue;
                }
                //HardIP would follow  Witrim
                if ((funcType == EnumVbtFuncType.WiTrim || funcType == EnumVbtFuncType.HardIP) &&
                    scgh.GetInitList().Count == 0 &&
                    scgh.GetPayloadList().Count == 1)
                {
                    continue;
                }

                string dsscSetupName = "";
                if (testNameFromScgh)
                {
                    dsscSetupName =
                    scgh.Item + "_DSSCSetup";
                }
                else
                {
                    dsscSetupName =
                    funcType + "_" + scgh.Block + "_" + CommonGenerator.GetSubBlockNameByPattern(scgh.PayloadValue, scgh.Item, false).Replace("_", "") + "_DSSCSetup";
                }
                dsscSetupName = TemplateAutoGenHelpers1.CheckDuplicateDSSCName(dsscSetupName, _dSSCSetupName_Sheet);

                setupNameList.Add(dsscSetupName);
                if (!registerSheets.TryGetValue(funcType, out DsscSetupSheet? existingRegisterSheet))
                {
                    existingRegisterSheet = new DsscSetupSheet { SheetName = funcType.ToString() };
                    registerSheets.Add(funcType, existingRegisterSheet);
                }

                registerSheet = existingRegisterSheet;
                Dictionary<string, Dictionary<string, string>> digSrcDic = RegisterLimitationGenerator.ProcessDigSrcDictionary(scgh);
                _dSSCSetupName_Sheet.Add(dsscSetupName);
                bool writeDsscSetupName = true;
                int patternIndex = 0;
                foreach (string initPat in scgh.InitList)
                {
                    try
                    {
                        string initSetupName = GetBBSetupName(funcType.ToString(), scgh.InitAliasList[patternIndex]) +
                                            "_DSSCSetup";
                        var postProcessRows = new List<PostProcessSheetRow>();
                        List<DsscSetupSheetRow>? regRows = DsscRegisterSheetRowGenerator.CreateRegisterSheetRows(initPat, dsscSetupName, ref writeDsscSetupName, funcType, digSrcDic, ref postProcessRows, _dSSCPostProcessName_Sheet);
                        if (regRows != null)
                        {
                            registerSheet.RowList.AddRange(regRows);
                        }

                        if (postProcessRows != null)
                        {
                            registerSheet.PostProcessRowList.AddRange(postProcessRows);
                        }

                        if (setupNameList.Contains(initSetupName))
                        {
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorMessageBox.Show(string.Format(ex.ToString()));
                    }
                    patternIndex++;
                }
                patternIndex = 0;
                foreach (string payload in scgh.PayloadList)
                {
                    try
                    {
                        string initSetupName = GetBBSetupName(funcType.ToString(), scgh.PayloadAliasList.FirstOrDefault() ?? "") +
                                            "_DSSCSetup";
                        //if (payload.Equals(scgh.PayloadValue, StringComparison.OrdinalIgnoreCase) &&
                        //    (funcType == VBTFuncType.WiTrim || funcType == VBTFuncType.RFTrim || funcType == VBTFuncType.HardIP))
                        //    continue;
                        var postProcessRows = new List<PostProcessSheetRow>();
                        List<DsscSetupSheetRow>? regRows = DsscRegisterSheetRowGenerator.CreateRegisterSheetRows(payload, dsscSetupName, ref writeDsscSetupName, funcType, digSrcDic, ref postProcessRows, _dSSCPostProcessName_Sheet);
                        if (regRows != null)
                        {
                            registerSheet.RowList.AddRange(regRows);
                        }

                        if (postProcessRows != null)
                        {
                            registerSheet.PostProcessRowList.AddRange(postProcessRows);
                        }

                        if (setupNameList.Contains(initSetupName))
                        {
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorMessageBox.Show(string.Format(ex.ToString()));
                    }
                    patternIndex++;
                }

            }

            return registerSheets;
        }

        private static string GetBBSetupName(string funcType, string item)
        {
            List<string> itemList = [.. item.Split('_')];
            if (!itemList[0].EqualsIgnoreCase(funcType))
            {
                itemList.Insert(0, funcType);
            }

            return string.Join("_", itemList);
        }
    }
}
