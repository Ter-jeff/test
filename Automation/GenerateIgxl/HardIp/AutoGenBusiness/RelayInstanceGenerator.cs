using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.PostAction.Relay.RelayConst;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    public class RelayInstanceGenerator
    {
        /// <summary>
        /// Add HardIP relay instances to "TestInst_Common" sheet
        /// </summary>
        /// <param name="planDic"></param>
        public void GenRelayInstance(Dictionary<string, HardIpSheet> planDic)
        {
            InstanceSheet commonInstanceSheet;
            bool exist = false;
            string key = "";
            foreach (KeyValuePair<string, InstanceSheet> insSheet in TestProgram.IgxlWorkBk.InsSheets)
            {
                if (string.Equals(insSheet.Value.Name, OutputConst.RelayInstName, StringComparison.OrdinalIgnoreCase))
                {
                    exist = true;
                    key = insSheet.Key;
                    break;
                }
            }

            if (!exist)
            {
                commonInstanceSheet = new InstanceSheet(OutputConst.RelayInstName);
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirCommonSheets, commonInstanceSheet);
            }
            else
            {
                commonInstanceSheet = TestProgram.IgxlWorkBk.InsSheets[key];
            }


            //Record all relay instances names to prevent duplicate relay instance
            foreach (string sheet in planDic.Keys)
            {
                foreach (HardIpPattern pattern in planDic[sheet].Rows)
                {
                    //relay setting may be "ReverseSetting"+"&"+"NewSetting"
                    foreach (KeyValuePair<string, string> item in pattern.RelaySetting)
                    {
                        string settingInJob = item.Value;
                        foreach (string setting in settingInJob.Split(';'))
                        {
                            GenerateRelay(commonInstanceSheet, setting);
                        }
                    }
                }

                #region Reset Relay Setting
                if (planDic[sheet].Rows.Count == 0)
                {
                    continue;
                }

                var lastSetting = new Dictionary<string, string>();
                HardIpPattern lastPlanItem = planDic[sheet].Rows.LastOrDefault(x => x.RelaySetting.Count != 0);
                if (lastPlanItem != null)
                {
                    lastSetting = lastPlanItem.RelaySetting;
                }

                if (lastSetting.Count > 0)
                {
                    foreach (KeyValuePair<string, string> item in lastSetting)
                    {
                        string setting = item.Value;
                        //If the last setting has both reverse setting and newSetting, just reverse the newSetting
                        if (item.Value.Contains(';'))
                        {
                            setting = item.Value.Split(';')[1];
                        }

                        string reverseSetting = SearchInfo.ReverseRelaySetting(setting);
                        GenerateRelay(commonInstanceSheet, reverseSetting);
                    }
                }

                #endregion
            }
        }

        private void GenerateRelay(InstanceSheet commonInstanceSheet, string setting)
        {
            string atgrelayName = HardIpConstData.PrefixAtgRelay + setting.Replace(":", "_").Replace("&", "_");
            if (setting == "" ||
                commonInstanceSheet.Rows.Any(comInst => comInst.TestName.EqualsIgnoreCase(atgrelayName)))
            {
                return;
            }

            SearchInfo.GetRelayArgs(setting, out string argOn, out string argOff);
            var relayInstance = new InstanceRow { TestName = atgrelayName };

            var relayFunction = new OutputConst(argOn, argOff);
            relayInstance.VbtType = relayFunction.Function.Type;
            relayInstance.VbtName = relayFunction.Function.FullFunctionName;
            relayInstance.ArgList = relayFunction.Function.Parameters;
            relayInstance.Args = relayFunction.Function.ArgList;

            commonInstanceSheet.AddRow(relayInstance);
        }
    }
}
