namespace Automation.Utility.Pattern
{
    public class MbistInfoCfg
    {
        public bool Pp { get; set; }
        public bool Dd { get; set; }
        public bool Fa { get; set; }
        public bool Cz { get; set; }
        public int SocOffset { get; set; }
        public int CpuOffset { get; set; }
        public int GfxOffset { get; set; }
        public bool SocMoreThan1Block { get; set; }
        public bool CpuMoreThan1Block { get; set; }
        public bool GfxMoreThan1Block { get; set; }
        public bool IgnoreDontCare { get; set; }
    }
}
