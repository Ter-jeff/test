using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.BistBira.Base;
using Automation.GenerateIgxl.BistBira.NonLogicData;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Singleton;

using CommonLib.Enums;

using ScghLib.Reader;

namespace Automation.GenerateIgxl.BistBira.BistInputLib
{
    public class VoltageConverter
    {
        private readonly MultiTestSettingSheetsSingleton _multiTestSetting;
        private static readonly Regex _regex = new Regex(@"\w+_(?<chiplet>[A-z]\d+$)", RegexOptions.Compiled);

        public VoltageConverter(MultiTestSettingSheetsSingleton multiTestSetting)
        {
            _multiTestSetting = multiTestSetting;
        }

        public void WorkFlow(ref BistProdFlowSheet prodFlow)
        {
            BistNaming naming = new BistNaming(new MbistConfig());
            for (int i = 0; i < prodFlow.Rows.Count; i++)
            {
                string retentionType = BistNonLogicalLib.CheckRetentionNew(prodFlow.Rows[i].Pattern);
                if (string.IsNullOrEmpty(retentionType) || !BistNaming.GetVoltageType(prodFlow.Rows[i].Voltage).Equals(BistConst.ConNv, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string module = naming.GetModule(prodFlow.MbistSheet.SheetName);
                string pModeModule = naming.GetPatternModule(prodFlow.Rows[i].VoltageMode, module);
                string rowPerformanceMode = prodFlow.Rows[i].IsDsscRow ? prodFlow.Rows[i].OriPerformance : prodFlow.Rows[i].VoltageMode;
                string chiplet = _regex.Match(prodFlow.MbistSheet.SheetName).Groups["chiplet"].ToString();
                string dcCategory = _multiTestSetting.FindMbistCatgeoryName(module, retentionType, rowPerformanceMode, null, out EnumMessageLevel _, out _, chiplet, pModeModule);

                // DC category mode in SCGH Mbist_SOC_NRT_X,LV
                string[] arr = prodFlow.Rows[i].Voltage.Split(',', ' ');
                if (_multiTestSetting.DcCategoryInfos.Exists(s => s.CategoryName.Equals(arr[0], StringComparison.OrdinalIgnoreCase)))
                {
                    dcCategory = arr[0];
                    string volt = arr.Length > 1 ? arr[1] : BistConst.ConNv;
                    List<string> voltagesInTs = _multiTestSetting.GetDcCategoryVoltages(dcCategory);
                    if (!voltagesInTs.Contains(volt))
                    {
                        volt = voltagesInTs.Contains(BistConst.ConHv) ? BistConst.ConHv : BistConst.ConLv;
                    }

                    prodFlow.Rows[i].Voltage = dcCategory + "," + volt;
                }
                else
                {
                    if (!string.IsNullOrEmpty(dcCategory))
                    {
                        List<string> voltagesInTs = _multiTestSetting.GetDcCategoryVoltages(dcCategory);
                        string voltage = BistConst.ConNv;
                        if (!voltagesInTs.Contains(BistConst.ConNv))
                        {
                            if (voltagesInTs.Contains(BistConst.ConHv) && voltagesInTs.Contains(BistConst.ConLv))
                            {
                                voltage = BistConst.ConNv;
                            }
                            else
                            {
                                voltage = voltagesInTs.Contains(BistConst.ConHv) ? BistConst.ConHv : BistConst.ConLv;
                            }
                        }
                        prodFlow.Rows[i].Voltage = voltage;
                    }
                }
            }
        }
    }
}
