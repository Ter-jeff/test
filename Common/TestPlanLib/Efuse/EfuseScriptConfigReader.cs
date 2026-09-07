using System.Xml;

namespace TestPlanLib.Efuse
{
    public class EfuseScriptConfigReader
    {
        public static EfuseScriptConfig LoadXml(string fileName)
        {
            var efuseScriptConfig = new EfuseScriptConfig();
            efuseScriptConfig.CrCdescription.Clear();
            efuseScriptConfig.AliasBlockList.Clear();
            efuseScriptConfig.UdrcmpMapping.Clear();
            efuseScriptConfig.DatalogBlockList.Clear();
            efuseScriptConfig.AliasProgStageDic.Clear();
            efuseScriptConfig.IsUsePostCheck = false;
            var doc = new XmlDocument();

            var settings = new XmlReaderSettings { IgnoreComments = true };
            var reader = XmlReader.Create(fileName, settings);
            doc.Load(reader);

            efuseScriptConfig.UpdateCRCBlockInformaiton(doc);
            efuseScriptConfig.UpdateAliasBlockList(doc);
            efuseScriptConfig.BaseVoltage = int.Parse(EfuseScriptConfig.AccessBaseVoltage(doc)!);
            efuseScriptConfig.BaseVoltageResolution = double.Parse(EfuseScriptConfig.AccessBaseVoltageResolution(doc)!);
            efuseScriptConfig.AccessUDRCMPMappingList(doc);
            efuseScriptConfig.AccessDatalogBlockList(doc);
            efuseScriptConfig.AccessProgStageMappingDic(doc);
            efuseScriptConfig.IsUsePostCheck = EfuseScriptConfig.AccessPostCheckFlag(doc);
            return efuseScriptConfig;
        }
    }
}
