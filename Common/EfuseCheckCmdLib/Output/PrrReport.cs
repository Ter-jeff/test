namespace EfuseCheckCmdLib.Output
{
    internal class PrrReport
    {
        public int Site { get; internal set; }
        public int Line { get; internal set; }
        public string Type { get; internal set; } = "";
        public string PrrInLog { get; internal set; } = "";
        public string PrrByEfuse { get; internal set; } = "";
        public string Result { get; internal set; } = "";
    }
}
