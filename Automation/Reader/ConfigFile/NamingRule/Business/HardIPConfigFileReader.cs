using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Xml;

using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Static;

namespace Automation.Reader.ConfigFile.NamingRule.Business
{
    internal class HardIpConfigFileReader : ConfigReaderBase<HardIpConfig>
    {
        private const string ConKeyPosition = "KeyPosition";
        private const string ConKey = "Key";

        public override HardIpConfig ReadConfig()
        {
            try
            {
                string configFilePath = Path.Combine(LocalSpecs.SettingFolder, "Settings", "SCGH", $"HardIP_Config_{LocalSpecs.CurrentProject}.xml");
                if (!string.IsNullOrEmpty(LocalSpecs.SettingFiles.HardipConfig))
                {
                    configFilePath = LocalSpecs.SettingFiles.HardipConfig;
                }

                if (!File.Exists(configFilePath))
                {
                    string file = Path.Combine(Directory.GetCurrentDirectory(), "Settings", "SCGH", "HardIP_Config_Default.xml");
                    if (File.Exists(file))
                    {
                        configFilePath = file;
                    }
                }

                var config = new HardIpConfig();
                if (!File.Exists(configFilePath))
                {
                    config.NamingRulesList = new Dictionary<string, RtosConfig>();
                    return config;
                }
                var doc = new XmlDocument();
                var settings = new XmlReaderSettings
                {
                    IgnoreComments = true
                };
                using (var reader = XmlReader.Create(configFilePath, settings))
                {
                    doc.Load(reader);
                    XmlNode configNode = doc.SelectSingleNode("HardIPConfig");
                    XmlNode node = configNode?.SelectSingleNode("NamingRule");
                    if (node != null)
                    {
                        Dictionary<string, RtosConfig> namingRules = ReadNamingRule(node);
                        config.NamingRulesList = namingRules;
                    }
                    if (configNode != null && configNode.SelectSingleNode("SplitRule") != null)
                    {
                        config.SplitRuleList.AddRange(ReadSplitRule(configNode.SelectSingleNode("SplitRule")));
                    }
                    if (configNode != null && configNode.SelectSingleNode("PayloadType") != null)
                    {
                        config.PayloadType = ReadPayloadType(configNode.SelectSingleNode("PayloadType"));
                    }
                    if (configNode != null && configNode.SelectSingleNode("SpecialFlagSetting") != null)
                    {
                        config.SpecialFlagSetting = ReadFlagNamingRule(configNode.SelectSingleNode("SpecialFlagSetting"));
                    }
                }
                return config;
            }
            catch (Exception e)
            {

                throw new Exception(e.StackTrace + "Error occurs when Reading HardIP Config, please Check Config file!");
            }

        }

        private List<HardIpConfigSplitRule> ReadSplitRule(XmlNode splitRulNode)
        {
            var splitRuleList = new List<HardIpConfigSplitRule>();
            XmlNodeList splitRuleNodeList = splitRulNode.ChildNodes;
            foreach (XmlNode node in splitRuleNodeList)
            {
                var sRule = new HardIpConfigSplitRule
                {
                    SubName = node.Name,
                    SubRule = node.InnerText
                };
                splitRuleList.Add(sRule);
            }
            return splitRuleList;
        }

        internal DataTable ReadPayloadType(XmlNode pNode)
        {
            var payloadTable = new DataTable();
            XmlNode positionNode = pNode.SelectSingleNode(ConKeyPosition);
            if (positionNode != null)
            {
                string[] value = positionNode.InnerText.Split(',');
                payloadTable.Columns.Add("key");
                foreach (string s in value)
                {
                    payloadTable.Columns.Add(s.Trim());
                }
                XmlNode keyNode = pNode.SelectSingleNode(ConKey);
                if (keyNode != null)
                {
                    XmlNodeList nodeList = keyNode.ChildNodes;
                    foreach (XmlNode node in nodeList)
                    {
                        DataRow row = payloadTable.NewRow();
                        row[0] = node.Name;
                        value = node.InnerText.Split(',');
                        for (int i = 0; i < value.Length; i++)
                        {
                            row[i + 1] = value[i].Trim();
                        }
                        payloadTable.Rows.Add(row);
                    }
                }
            }

            return payloadTable;
        }
    }
}
