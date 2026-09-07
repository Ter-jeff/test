using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.Singleton;

using IgxlLib.IgxlBase;

using LogLib.Static;

using TestPlanLib.BinNumberLegacy;
using TestPlanLib.Singleton;

namespace Automation.GenerateIgxl.Scan.Base
{
    public class ScanNaming
    {
        private readonly Dictionary<string, string> _modes = PerformanceModeSingleton.Instance().GetAllPerformanceModeDic();

        public static string ConModuleCpu { get; private set; } = "";

        public static string ConModuleGpu { get; private set; } = "";

        public static string ConModuleSoc { get; private set; } = "";

        public ScanNaming()
        {
            ConModuleCpu = ModuleSingleton.Instance().ModuleCpu;
            ConModuleGpu = ModuleSingleton.Instance().ModuleGfx;
            ConModuleSoc = ModuleSingleton.Instance().ModuleSoc;
        }

        public string GetSoftBinNumber(string binName, string module, string payloadType, string performanceMode, string level, bool alarm = true)
        {
            var para = new SoftBinNumPara();
            if (Regex.IsMatch(payloadType, "sa", RegexOptions.IgnoreCase))
            {
                para.Category = "Low speed Scan";
            }

            para.Module = "SCAN";
            para.Block = module;
            para.Level = level;
            para.PerformanceMode = performanceMode;
            para.SubModule = payloadType;
            return BinNumberSingletonLegacy.Instance().GetSoftBinNumber(_modes, Response.Report, alarm, binName, para);
        }

        public string GetHardBinNumber(BinTableRow binTableRow)
        {
            var para = new BinNumDefPara(EnumBinNumDefBlock.Scan, "TD");
            BinNumberSingletonLegacy.Instance().GetHardBinNum(para, out string hardBin);
            binTableRow.Comment = $"Block : {para.Block} , Condition : {para.Condition}";
            return hardBin;
        }
    }
}
