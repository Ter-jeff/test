using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;

namespace Automation.GenerateIgxl.HardIp.InputReader
{
    public class ParseTestPlanByCondition
    {
        public void ParseTestPlanPatternByCondition(Dictionary<string, HardIpSheet> testPlanDic)
        {
            foreach (KeyValuePair<string, HardIpSheet> pair in testPlanDic)
            {
                List<HardIpPattern> patternList = pair.Value.Rows;
                for (int index = 0; index < patternList.Count; index++)
                {
                    HardIpPattern hardIpPattern = patternList[index];
                    if (hardIpPattern.ForceConditionList.Count > 1)
                    {
                        patternList.Remove(hardIpPattern);
                        int conditionIndex = 1;
                        for (int i = 0; i < hardIpPattern.ForceConditionList.Count; i++)
                        {
                            var forceVoltages = hardIpPattern.ForceConditionList[i].ForcePins.Where(n => n.ForceLabelVoltages.Count > 0).SelectMany(n => n.ForceLabelVoltages).Distinct().ToList();
                            HardIpPattern newPattern = hardIpPattern.Copy();
                            newPattern.ConditionIndex = conditionIndex;
                            newPattern.ForceConditionList = new List<ForceCondition> { hardIpPattern.ForceConditionList[i] };
                            ForcePin prefixForcePin = newPattern.ForceConditionList.SelectMany(q => q.ForcePins).First();
                            newPattern.ForceVoltageFlag = "_" + $"{prefixForcePin.PinName}{prefixForcePin.ForceValue}".Replace(".", "p").Replace("_", "");
                            if (i != 0 && forceVoltages.Count != 0 && forceVoltages.Count != 3)
                            {
                                var skipVoltage = new List<string>
                                {
                                    HardIpConstData.LabelLv,
                                    HardIpConstData.LabelHv,
                                    HardIpConstData.LabelNv
                                };
                                foreach (string voltage in forceVoltages)
                                {
                                    skipVoltage.Remove(voltage);
                                }
                                newPattern.SkipList = skipVoltage;
                            }
                            patternList.Insert(index, newPattern);
                            index++;
                            conditionIndex++;
                        }
                        index--;
                    }
                }
            }
        }
    }
}
