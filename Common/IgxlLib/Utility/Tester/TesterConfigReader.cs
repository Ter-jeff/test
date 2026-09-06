using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

using CommonLib.Extension;

using IgxlLib.Regs;

namespace IgxlLib.Utility.Tester
{
    public static class TesterConfigReader
    {

        public static Dictionary<string, TesterConfig>? GetTesterConfigs(string configName = "")
        {
            string cfgFile = !string.IsNullOrEmpty(configName) ? configName : Path.Combine(AppContext.BaseDirectory, "Config", "Tester", "TesterConfig_Default.xml");
            if (!File.Exists(cfgFile))
            {
                return null;
            }

            var doc = new XmlDocument();
            var settings = new XmlReaderSettings { IgnoreComments = true };
            var reader = XmlReader.Create(cfgFile, settings);
            doc.Load(reader);
            var dic = new Dictionary<string, TesterConfig>(StringExtensions.IgnoreCase);
            XmlNode? configNode = doc.SelectSingleNode("UflexConfig");
            if (configNode != null)
            {
                XmlNode? nodeCp = configNode.SelectSingleNode("CP");
                if (nodeCp != null)
                {
                    TesterConfig testerCfg = GetUltraFlexConfig(nodeCp);
                    dic.Add("CP", testerCfg);
                }

                XmlNode? nodeFt = configNode.SelectSingleNode("FT");
                if (nodeFt != null)
                {
                    TesterConfig testerCfg = GetUltraFlexConfig(nodeFt);
                    dic.Add("FT", testerCfg);
                }

                if (nodeCp == null && nodeFt == null)
                {
                    TesterConfig testerCfgAll = GetUltraFlexConfig(configNode);
                    dic.Add("All", testerCfgAll);
                }
            }

            return dic;
        }

        private static TesterConfig GetUltraFlexConfig(XmlNode xmlNode)
        {
            var testerCfg = new TesterConfig();
            XmlNode? node = xmlNode.SelectSingleNode("VSM");
            if (node != null)
            {
                testerCfg.Vsm = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("IO");
            if (node != null)
            {
                testerCfg.Io = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("HexVS");
            if (node != null)
            {
                testerCfg.HexVs = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("UVS256");
            if (node != null)
            {
                testerCfg.Uvs256 = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("UVS256HP");
            if (node != null)
            {
                testerCfg.Uvs256Hp = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("UVS64");
            if (node != null)
            {
                testerCfg.Uvs64 = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("UVI80");
            if (node != null)
            {
                testerCfg.Uvi80 = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("Support");
            if (node != null)
            {
                testerCfg.Support = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("UltraPAC");
            if (node != null)
            {
                testerCfg.UltraPac = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("DC30");
            if (node != null)
            {
                testerCfg.Dc30 = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("US10G");
            if (node != null)
            {
                testerCfg.Us10G = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("UW24Source");
            if (node != null)
            {
                testerCfg.Uw24Source = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("UW24Measure");
            if (node != null)
            {
                testerCfg.Uw24Measure = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            node = xmlNode.SelectSingleNode("MX8");
            if (node != null)
            {
                testerCfg.Mx8 = Reg.WhitespaceRegex.Replace(node.InnerText, "");
            }

            return testerCfg;
        }
    }
}
