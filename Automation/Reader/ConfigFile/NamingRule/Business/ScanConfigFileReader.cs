using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Regs;
using Automation.Static;

using CommonLib.Extension;

namespace Automation.Reader.ConfigFile.NamingRule.Business
{
    public class ScanConfigFileReader : ConfigReaderBase<ScanConfig>
    {
        private const string ConKeyPosition = "KeyPosition";
        private const string ConKey = "Key";

        public override ScanConfig ReadConfig()
        {
            try
            {
                string configFilePath = Path.Combine(LocalSpecs.SettingFolder, "Settings", "SCGH", $"Scan_Config_{LocalSpecs.CurrentProject}.xml");
                if (!string.IsNullOrEmpty(LocalSpecs.SettingFiles.ScanConfig))
                {
                    configFilePath = LocalSpecs.SettingFiles.ScanConfig;
                }

                if (!File.Exists(configFilePath))
                {
                    string file = Path.Combine(Directory.GetCurrentDirectory(), "Settings", "SCGH", "Scan_Config_Default.xml");
                    if (File.Exists(file))
                    {
                        configFilePath = file;
                    }
                }

                var config = new ScanConfig();
                var doc = new XmlDocument();
                var settings = new XmlReaderSettings
                {
                    IgnoreComments = true
                };
                var reader = XmlReader.Create(configFilePath, settings);
                doc.Load(reader);
                XmlNode configNode = doc.SelectSingleNode("ScanConfig");
                XmlNode node = configNode?.SelectSingleNode("NamingRule");
                if (node != null)
                {
                    Dictionary<string, List<RtosConfig>> namingRules = ReadNamingRuleNew(node);
                    config.NewNamingRulesList = namingRules;
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
                return config;
            }
            catch (Exception e)
            {
                throw new Exception(e.StackTrace + "Error occurs when Reading Scan Config, please Check Config file!");
            }
        }

        private List<ScanConfigSplitRule> ReadSplitRule(XmlNode splitRulNode)
        {
            var splitRuleList = new List<ScanConfigSplitRule>();
            XmlNodeList splitRuleNodeList = splitRulNode.ChildNodes;
            foreach (XmlNode node in splitRuleNodeList)
            {
                var sRule = new ScanConfigSplitRule
                {
                    SubName = node.Name,
                    SubRule = node.InnerText
                };
                splitRuleList.Add(sRule);
            }
            return splitRuleList;
        }

        private Dictionary<string, List<RtosConfig>> ReadNamingRuleNew(XmlNode namingNode)
        {
            var configNamingRules = new List<RtosConfig>();
            XmlNodeList splitRuleNodeList = namingNode.ChildNodes;
            foreach (XmlNode node in splitRuleNodeList)
            {
                if (node != null)
                {
                    string name;
                    if (node.Name == "Sa" || node.Name == "Td")
                    {
                        string block = node.Name == "Sa" ? "Sa" : "Td";

                        foreach (XmlNode initNode in node.ChildNodes)
                        {
                            var rule = new RtosConfig
                            {
                                Block = block
                            };
                            name = initNode.Name;
                            rule.Module = name;
                            Regex regexInit = RegCommon.RegexInit;
                            foreach (XmlNode childNode in initNode.ChildNodes)
                            {
                                Match initMatch = regexInit.Match(childNode.Name);
                                if (initMatch.Success)
                                {
                                    rule.NamingRuleInitList.Add(childNode.InnerText);
                                    rule.NamingRuleInitSequenceList.Add(Convert.ToInt16(initMatch.Groups["idx"].ToString()));
                                }

                                if (childNode.Name.ContainsIgnoreCase("payload"))
                                {
                                    rule.NamingRulePayloadList.Add(childNode.InnerText);
                                }
                            }
                            configNamingRules.Add(rule);
                        }
                    }
                    else
                    {
                        var rule = new RtosConfig();
                        name = node.Name;
                        rule.Module = name;
                        rule.Block = "Sa";
                        Regex regexInit = RegCommon.RegexInit;
                        foreach (XmlNode childNode in node.ChildNodes)
                        {
                            Match initMatch = regexInit.Match(childNode.Name);
                            if (initMatch.Success)
                            {
                                rule.NamingRuleInitList.Add(childNode.InnerText);
                                rule.NamingRuleInitSequenceList.Add(Convert.ToInt16(initMatch.Groups["idx"].ToString()));
                            }


                            if (childNode.Name.ContainsIgnoreCase("payload"))
                            {
                                rule.NamingRulePayloadList.Add(childNode.InnerText);
                            }
                        }
                        var rule2 = new RtosConfig
                        {
                            Block = "Td",
                            Module = rule.Module,
                            NamingRuleInitList = rule.NamingRuleInitList,
                            NamingRulePayloadList = rule.NamingRulePayloadList,
                            NamingRuleInitSequenceList = rule.NamingRuleInitSequenceList
                        };

                        configNamingRules.Add(rule);
                        configNamingRules.Add(rule2);
                    }


                }
            }
            var temp = configNamingRules.GroupBy(p => p.Module).ToDictionary(p => p.Key, p => p.ToList());
            return temp;
        }

        private DataTable ReadPayloadType(XmlNode xmlNode)
        {
            var payloadTable = new DataTable();
            XmlNode singleNode = xmlNode.SelectSingleNode(ConKeyPosition);
            if (singleNode != null)
            {
                string[] value = singleNode.InnerText.Split(',');
                payloadTable.Columns.Add("key");
                foreach (string s in value)
                {
                    payloadTable.Columns.Add(s.Trim());
                }

                XmlNode keyNode = xmlNode.SelectSingleNode(ConKey);
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
