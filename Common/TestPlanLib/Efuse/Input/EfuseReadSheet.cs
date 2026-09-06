using System.Collections.Generic;

using CommonReaderLib;

namespace TestPlanLib.Efuse.Input
{
    public class EfuseReadSheet : MySheet
    {
        public List<EfuseReadRow> Rows = [];

        public int TestNameColIdx = -1;
        public int TypeColIdx = -1;
        public int BankColIdx = -1;
        public int WriteReadColIdx = -1;
        public int PurposeColIdx = -1;
        public int UserDefinedColIdx = -1;
        public int InitPatColIdx = -1;
        public int PayloadPatColIdx = -1;
        public List<int> JobColIdx = [];

        public EfuseReadSheet(string sheetName)
        {
            SheetName = sheetName;
        }
    }

    public class EfuseReadRow : MyRow
    {
        public string TestName = "";
        public string Type = "";
        public string Bank = "";
        public string WriteRead = "";
        public string Purpose = "";
        public string UserDefined = "";
        public List<string> InitList = [];
        public List<string> PayloadList = [];
        public Dictionary<string, string> JobDictionary = [];
    }
}
