using System;
using System.Collections.Generic;

using Automation.Singleton;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.GenWireByPattern
{
    public class NWireByPatternMain
    {
        public void WorkFlow(Dictionary<string, SubFlowSheet> subFlowSheets)
        {
            Dictionary<string, string> patternDic = NwireSingleton.Instance().SettingInfo.PatternDic;
            if (patternDic == null || patternDic.Count == 0)
            {
                return;
            }

            foreach (KeyValuePair<string, SubFlowSheet> subFlowSheet in subFlowSheets)
            {
                bool allowFlag = true;
                foreach (KeyValuePair<string, string> pattern in patternDic)
                {
                    for (int i = 0; i < subFlowSheet.Value.Rows.Count; i++)
                    {
                        if (subFlowSheet.Value.Rows[i].Opcode == null)
                        {
                            continue;
                        }

                        if (subFlowSheet.Value.Rows[i].Opcode.Equals("Test", StringComparison.OrdinalIgnoreCase) || subFlowSheet.Value.Rows[i].Opcode.Equals("characterize", StringComparison.OrdinalIgnoreCase))
                        {
                            if (subFlowSheet.Value.Rows[i].Parameter.ContainsIgnoreCase(pattern.Key.ToUpper()))
                            {
                                if (allowFlag)
                                {
                                    FlowRow call = NwireSingleton.Instance().SettingInfo.GetNwirePatternCall(pattern.Value);
                                    subFlowSheet.Value.Rows.Insert(i, call);
                                    allowFlag = false;
                                }
                            }
                            else
                            {
                                allowFlag = true;
                            }
                        }
                    }
                }
            }
        }
    }
}
