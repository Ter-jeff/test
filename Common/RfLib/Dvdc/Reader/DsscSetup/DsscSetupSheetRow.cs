using RfLib.Dvdc.GenTemplate.Bussiness;

namespace RfLib.Dvdc.Reader.DsscSetup
{
    public class DsscSetupSheetRow
    {
        public EnumVbtFuncType Type { get; set; }
        public string Block { get; set; } = "";
        public string DsscSetup { get; set; } = "";
        public string Pattern { get; set; } = "";
        public string DigSrcEqn { get; set; } = "";
        public string DigSrcReg { get; set; } = "";
        public string DigSrcAssignment { get; set; } = "";
        public string DigSrcPin { get; set; } = "";
        public string DigSrcSampleSize { get; set; } = "";
        public string DigCapPin { get; set; } = "";
        public string DigCapSampleSize { get; set; } = "";
        public string CusStrDigCapData { get; set; } = "";
        public string PatModuleInfo { get; set; } = "";
    }
}
