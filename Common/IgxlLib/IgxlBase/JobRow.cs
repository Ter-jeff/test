namespace IgxlLib.IgxlBase
{
    public class JobRow : IgxlRow
    {
        public string JobName { get; set; } = string.Empty;
        public string PinMap { get; set; } = string.Empty;
        public string TestInstances { get; set; } = string.Empty;
        public string FlowTable { get; set; } = string.Empty;
        public string AcSpecs { get; set; } = string.Empty;
        public string DcSpecs { get; set; } = string.Empty;
        public string PatternSets { get; set; } = string.Empty;
        public string PatternGroups { get; set; } = string.Empty;
        public string BinTable { get; set; } = string.Empty;
        public string Characterization { get; set; } = string.Empty;
        public string TestProcedures { get; set; } = string.Empty;
        public string MixedSignalTiming { get; set; } = string.Empty;
        public string WaveDefinitions { get; set; } = string.Empty;
        public string PSets { get; set; } = string.Empty;
        public string Signals { get; set; } = string.Empty;
        public string PortMap { get; set; } = string.Empty;
        public string FractionalBus { get; set; } = string.Empty;
        public string ConcurrentSequence { get; set; } = string.Empty;
        public string SpikeCheckConfig { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
    }
}
