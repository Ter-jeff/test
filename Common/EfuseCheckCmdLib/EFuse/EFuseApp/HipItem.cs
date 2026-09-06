using System.Collections.Generic;

namespace EfuseCheckCmdLib.EFuse.EFuseApp
{
    public class HipItem(string pid)
    {
        public string Name = pid;
        public string X = "";
        public string Y = "";
        public Dictionary<string, string> HipData = [];
    }
}
