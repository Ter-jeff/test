namespace EfuseCheckCmdLib.Datalog
{
    public class PrrRow
    {
        public int Site { get; internal set; }
        public string Type { get; internal set; } = "";
        public string Prr { get; internal set; } = "";
        public PrrLine Line { get; internal set; } = null!;
    }
}
