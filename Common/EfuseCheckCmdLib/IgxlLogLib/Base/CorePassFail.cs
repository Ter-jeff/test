namespace EfuseCheckCmdLib.IgxlLogLib.Base
{
    public class CorePassFail(string block, string subBlock, string data)
    {
        public string Block { get; set; } = block;
        public string SubBlock { get; set; } = subBlock;
        public string Data { get; set; } = data;
    }
}
