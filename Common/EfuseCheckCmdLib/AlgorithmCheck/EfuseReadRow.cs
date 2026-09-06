namespace EfuseCheckCmdLib.AlgorithmCheck
{
    public class EfuseReadRow
    {
        public string Name { get; internal set; } = "";
        public int Site { get; internal set; }
        public string Value { get; internal set; } = "";
        internal EfuseReadLine Line { get; set; } = null!;
    }
}
