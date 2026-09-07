namespace Cautogen.AutoCZ.CharPostProcessor.DataStructure
{
    public class TimeSetItem
    {
        public int Version { get; set; }
        public string TimeMod { get; set; }

        public TimeSetItem()
        {
            Version = -1;
            TimeMod = "";
        }

    }
    public class CompileITem
    {
        public string Product;
        public string Version;
        public string TpCategory;
        public string AtpName;
        public string OpCode;
        public string ScanMode;
        public string Halt;
        public string Compilation;
        public string Md5;
        public string HLv;
        public string ScanSetupTSet;

        public CompileITem()
        {
            Product = "";
            Version = "";
            TpCategory = "";
            AtpName = "";
            OpCode = "";
            ScanMode = "";
            Halt = "";
            Compilation = "";
            Md5 = "";
            HLv = "";
            ScanSetupTSet = "";
        }
    }

}
