using System.Collections.Generic;

using CommonReaderLib;

namespace TestPlanLib.CPM
{
    public class EfuseCpmSheet : MySheet
    {
        public List<EfuseCpmSheetRow> Rows = [];

        public int EfuseBankColNumber = -1;
        public int CpmEfuseName2ColNumber = -1;
        public int BitsColNumber = -1;
        public int DefaultValuesColNumber = -1;
        public int AlternateValuesColNumber = -1;
        public int CpmColNumber = -1;
        public int CommentColNumber = -1;
        public Dictionary<string, int> FlagColNumber = [];

        public EfuseCpmSheet(string sheetName)
        {
            SheetName = sheetName;
        }
    }

    public class EfuseCpmSheetRow : MyRow
    {
        public string EfuseBank = "";
        public string CpmEfuseName2 = "";
        public string Bits = "";
        public Dictionary<string, string> FlagsValue = [];
        public string AlternateValues = "";
        public string Cpm = "";
        public string Comment = "";
    }
}
