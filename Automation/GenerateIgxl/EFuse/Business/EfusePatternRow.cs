using System.Collections.Generic;

namespace Automation.GenerateIgxl.EFuse.Business
{
    public class EfusePatternRow
    {
        public string SheetName = "";
        public int RowNum;
        public string BankName = "";
        public EfusePatternType PatternType = new EfusePatternType();
        public List<string> InitList = new List<string>();
        public List<string> PayloadList = new List<string>();
        public List<string> PatList = new List<string>();
        public string ReadWritePin = "";
        public int SendStoreCount;
        public EfusePatternWorR ReadOrWrite = EfusePatternWorR.Non;
    }
}
