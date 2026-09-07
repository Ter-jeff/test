using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;

using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Regs;

using CommonLib.Extension;

namespace Automation.Reader.ConfigFile.NamingRule.Business
{
    public abstract class ConfigReaderBase<T> where T : ConfigBase
    {
        protected const string ConPayload = "Payload";
        public abstract T ReadConfig();

        protected Dictionary<string, RtosConfig> ReadNamingRule(XmlNode namingNode)
        {
            var namingRulesList = new Dictionary<string, RtosConfig>();
            Regex regexInit = RegCommon.RegexInit;
            XmlNodeList splitRuleNodeList = namingNode.ChildNodes;
            foreach (XmlNode node in splitRuleNodeList)
            {
                string name = "";
                var rule = new RtosConfig();
                if (node != null)
                {
                    name = node.Name;
                    for (int i = 1; i <= 10; i++)
                    {
                        string initName = "init" + i.ToString(CultureInfo.InvariantCulture);
                        XmlNode initNode = node.SelectSingleNode(initName);
                        if (initNode != null)
                        {
                            Match initMatch = regexInit.Match(initNode.Name);
                            if (initMatch.Success)
                            {
                                rule.NamingRuleInitList.Add(initNode.InnerText);
                                rule.NamingRuleInitSequenceList.Add(Convert.ToInt16(initMatch.Groups["idx"].ToString()));
                            }
                        }
                    }
                    foreach (XmlNode childNode in node.ChildNodes)
                    {
                        if (childNode.Name.ContainsIgnoreCase("payload"))
                        {
                            rule.NamingRulePayloadList.Add(childNode.InnerText);
                        }
                    }
                }
                namingRulesList.Add(name, rule);
            }
            return namingRulesList;
        }

        protected string ReadFlagNamingRule(XmlNode pNode)
        {
            XmlNode keyNode = pNode.SelectSingleNode(ConfigReaderBase<HardIpConfig>.ConPayload);
            if (keyNode != null)
            {
                return keyNode.InnerText;
            }

            return "";
        }
    }
}
