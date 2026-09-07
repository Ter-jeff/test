using CommonReaderLib.PatternListCsv;

namespace Automation.PreCheck.AllParaData
{
    public class HardIpParaData : ParaData
    {
        public string TargetPattern;
        public bool HvEnable { get; set; } = true;
        public bool LvEnable { get; set; } = true;
        public bool NvEnable { get; set; } = true;
        public bool CzHvEnable { get; set; } = true;
        public bool CzLvEnable { get; set; } = true;
        public bool CzNvEnable { get; set; } = true;
        public bool SplitCzFlow { get; set; }
        public EnumBlock Block { get; private set; }

        public HardIpParaData(EnumBlock block)
        {
            Block = block;
        }
    }
}
