using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Static;

namespace Automation.GenerateIgxl.HardIp.HardIpData.DataBase
{
    public class ConfigData
    {
        public Dictionary<string, string> VbtNameMapping;
        public Dictionary<string, string> InstSpecialSetting;
        public bool PatSetNameUsingSubBlock;
        public HashSet<string> AdaptiveCoolingItem;
        public bool NeedGenerateTmpsBin;
        public string CoolingLimitValue;
        public bool NoBinOutForAll;

        public ConfigData()
        {
            VbtNameMapping = new Dictionary<string, string>();
            InstSpecialSetting = new Dictionary<string, string>();

            string adpCoolBlock = LocalSpecs.Options.AdpativeCoolingBlock;
            string adpCoolSuBlock = LocalSpecs.Options.AdpativeCoolingSuBlock;
            if (!string.IsNullOrEmpty(adpCoolBlock) && !string.IsNullOrEmpty(adpCoolSuBlock))
            {
                string[] blockList = adpCoolBlock.Split(',');
                string[] subBlockList = adpCoolSuBlock.Split(',');
                AdaptiveCoolingItem = Enumerable
                    .Range(0, Math.Max(blockList.Length, subBlockList.Length))
                    .Select(i => (blockList.ElementAtOrDefault(i)?.Trim().Replace("_", "").ToUpper(),
                                    subBlockList.ElementAtOrDefault(i)?.Trim().Replace("_", "").ToUpper()))
                    .Where(p => !string.IsNullOrEmpty(p.Item1) &&
                              !string.IsNullOrEmpty(p.Item2))
                    .Select(p => $"{p.Item1}_{p.Item2}").ToHashSet(StringComparer.OrdinalIgnoreCase);
                NeedGenerateTmpsBin = true;
            }
            else
            {
                AdaptiveCoolingItem = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                NeedGenerateTmpsBin = false;
            }

            CoolingLimitValue = LocalSpecs.Options.CoolingLimitValue;
            NoBinOutForAll = LocalSpecs.Options.NoBinOutForAll;
        }
    }

}
