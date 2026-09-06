using System.Collections.Generic;
using System.Xml;

namespace TestPlanLib.Efuse
{

    public class EfuseScriptConfig
    {
        public Dictionary<string, string> CrCdescription = [];
        public List<string> PassBin = [];
        public Dictionary<string, string> AliasBlockList = [];
        public Dictionary<string, string> UdrcmpMapping = [];
        public int BaseVoltage { get; set; }
        public double BaseVoltageResolution { get; set; }
        public List<string> DatalogBlockList = [];
        public bool IsUsePostCheck;
        public Dictionary<string, string> AliasProgStageDic = [];

        public void UpdateCRCBlockInformaiton(XmlDocument xmlDocument)
        {
            XmlNodeList? multinode = xmlDocument.SelectNodes("EfuseCheck/Block");
            if (multinode == null)
            {
                return;
            }

            foreach (object node in multinode)
            {
                string blockname = SearchBlockName((XmlNode)node);
                string? bits = SearchBits((XmlNode)node);
                string? type = SearchType((XmlNode)node);
                string result = type + " = " + bits;
                CrCdescription.Add(blockname, result);
            }
        }

        private static string SearchBlockName(XmlNode xmlNode)
        {
            var element = (XmlElement)xmlNode;
            return element.GetAttribute("Name");
        }

        private static string? SearchType(XmlNode xmlNode)
        {
            var element = (XmlElement?)xmlNode.SelectSingleNode("Type");

            return element?.InnerText;
        }

        private static string? SearchBits(XmlNode xmlNode)
        {
            var element = (XmlElement?)xmlNode.SelectSingleNode("Bits");

            return element?.InnerText;
        }

        public void UpdateAliasBlockList(XmlNode xmlNode)
        {
            XmlNodeList? subnodes = xmlNode.SelectNodes("EfuseCheck/AliasBlockName");
            if (subnodes != null)
            {
                foreach (object subnode in subnodes)
                {
                    var element = (XmlElement)subnode;
                    string bdfBlock = element.GetAttribute("BDFBlock");
                    string datalogBlock = element.GetAttribute("DataLogBlock");
                    AliasBlockList.Add(datalogBlock, bdfBlock);
                }
            }
        }

        public static string? AccessBaseVoltage(XmlNode xmlNode)
        {
            var element = (XmlElement?)xmlNode.SelectSingleNode("EfuseCheck/BASE_VOLTAGE");

            return element?.InnerText;
        }

        public static string? AccessBaseVoltageResolution(XmlNode xmlNode)
        {
            var element = (XmlElement?)xmlNode.SelectSingleNode("EfuseCheck/BASE_VOLTAGE_Resoution");

            return element?.InnerText;
        }

        public void AccessUDRCMPMappingList(XmlNode xmlNode)
        {
            XmlNodeList? subnodes = xmlNode.SelectNodes("EfuseCheck/UDR");
            if (subnodes != null)
            {
                foreach (object subnode in subnodes)
                {
                    var element = (XmlElement)subnode;
                    string refBlock = element.GetAttribute("REFBLOCK");
                    string cmpBlock = element.GetAttribute("COMPAREBLOCK");
                    UdrcmpMapping.Add(cmpBlock, refBlock);
                }
            }
        }

        public void AccessProgStageMappingDic(XmlNode xmlNode)
        {
            XmlNodeList? subnodes = xmlNode.SelectNodes("EfuseCheck/ProgStage");
            if (subnodes != null)
            {
                foreach (object subnode in subnodes)
                {
                    var element = (XmlElement)subnode;
                    string jobMapping = element.GetAttribute("JobMapping");
                    if (!AliasProgStageDic.ContainsKey(jobMapping.Split(':')[1]))
                    {
                        AliasProgStageDic.Add(jobMapping.Split(':')[1], jobMapping.Split(':')[0]);
                    }
                }
            }
        }

        public void AccessDatalogBlockList(XmlNode xmlNode)
        {
            var element = (XmlElement?)xmlNode.SelectSingleNode("EfuseCheck/DatalogBlockList");
            if (element != null)
            {
                DatalogBlockList = [.. element.InnerText.Replace(" ", "").Split(',')];
            }
            else
            {
                DatalogBlockList = [];
            }
        }

        public static bool AccessPostCheckFlag(XmlNode xmlNode)
        {
            var element = (XmlElement?)xmlNode.SelectSingleNode("EfuseCheck/POSTCHECK");
            if (element == null || element.InnerText.Length == 0)
            {
                return false;
            }

            return element.InnerText == "1";
        }
    }
}
