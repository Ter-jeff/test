using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Static;

namespace Automation.Reader.ConfigFile.NamingRule.Business
{
    internal class MbistConfigReader : ConfigReaderBase<MbistConfig>
    {
        public override MbistConfig ReadConfig()
        {
            string configFilePath = Path.Combine(LocalSpecs.SettingFolder, "Settings", "SCGH", $"Mbist_Config_{LocalSpecs.CurrentProject}.xml");
            if (!string.IsNullOrEmpty(LocalSpecs.SettingFiles.MbistConfig))
            {
                configFilePath = LocalSpecs.SettingFiles.MbistConfig;
            }

            if (!File.Exists(configFilePath))
            {
                string file = Path.Combine(Directory.GetCurrentDirectory(), "Settings", "SCGH", "Mbist_Config_Default.xml");
                if (File.Exists(file))
                {
                    configFilePath = file;
                }
            }

            var namingRules = new Dictionary<string, RtosConfig>();
            var settings = new XmlReaderSettings
            {
                IgnoreComments = true
            };
            var config = new MbistConfig();
            var reader = XmlReader.Create(configFilePath, settings);
            var doc = new XmlDocument();
            doc.Load(reader);
            XmlNode configNode = doc.SelectSingleNode("MbistConfig");
            if (configNode != null)
            {
                XmlNode node = configNode.SelectSingleNode("NamingRule");
                if (node != null)
                {
                    namingRules = ReadNamingRule(node);
                }

                node = configNode.SelectSingleNode("SpecialFlagSetting");
                if (node != null)
                {
                    config.SpecialFlagSetting = ReadFlagNamingRule(node);
                }

                node = configNode.SelectSingleNode("RepairFlag");
                if (node != null)
                {
                    config.RepairFlagSetting = ReadRepairFlagSetting(node);
                }

                node = configNode.SelectSingleNode("SpecialNaming");
                if (node != null)
                {
                    config.SpecialNamingRules = ReadSpecialNamingRule(node);
                }
            }
            config.NamingRulesList = namingRules;
            return config;
        }

        private Dictionary<string, List<KeyAndPosition>> ReadRepairFlagSetting(XmlNode pNode)
        {
            var flagSetting = new Dictionary<string, List<KeyAndPosition>>();
            XmlNodeList nodeList = pNode.ChildNodes;
            foreach (XmlNode node in nodeList)
            {
                string name = "";
                var rule = new List<KeyAndPosition>();
                if (node != null)
                {
                    name = node.Name;
                    XmlNodeList subNodeList = node.ChildNodes;
                    var keys = new KeyAndPosition();
                    foreach (XmlNode subNode in subNodeList)
                    {
                        if (subNode.Name.Equals("KeyPosition", StringComparison.OrdinalIgnoreCase))
                        {
                            keys.KeyPositions = subNode.InnerText;
                        }
                        if (subNode.Name.Equals("Key", StringComparison.OrdinalIgnoreCase))
                        {
                            keys.Keys = subNode.InnerText;
                            rule.Add(keys);
                            keys = new KeyAndPosition();
                        }
                    }

                }
                flagSetting.Add(name, rule);
            }
            return flagSetting;
        }

        private Dictionary<string, MbistProductionSpecialNamingRule> ReadSpecialNamingRule(XmlNode pNode)
        {

            var specialNamingRuleDic = new Dictionary<string, MbistProductionSpecialNamingRule>();
            XmlNodeList nodeList = pNode.ChildNodes;
            foreach (XmlNode node in nodeList)
            {
                if (node != null)
                {
                    string name = node.Name;
                    XmlNodeList subNodeList = node.ChildNodes;
                    var specialNaming = new MbistProductionSpecialNamingRule();
                    foreach (XmlNode subNode in subNodeList)
                    {
                        if (subNode.Name.Equals("FieldFromLabel", StringComparison.OrdinalIgnoreCase))
                        {
                            specialNaming.UseLabelPosition = subNode.InnerText;
                        }
                        if (subNode.Name.Equals("VmarginHasNumber", StringComparison.OrdinalIgnoreCase))
                        {
                            specialNaming.VMarginNeedNumber = subNode.InnerText.Equals("1");
                        }
                        if (subNode.Name.Equals("FlagNameFromPattern", StringComparison.OrdinalIgnoreCase))
                        {
                            specialNaming.FlagUsePatternPosition = subNode.InnerText;
                        }
                    }
                    specialNamingRuleDic.Add(name, specialNaming);
                }
            }

            return specialNamingRuleDic;
        }
    }
}
