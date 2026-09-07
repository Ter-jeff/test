using System.Collections.Generic;

using CommonReaderLib;

namespace TestPlanLib.Efuse.Input
{
    public class EfuseArraySizeSheet : MySheet
    {
        public List<EfuseArraySizeRow> Rows = [];

        public int FuseArrayColNumber = -1;
        public int OfBitsColNumber = -1;
        public int OfFuseArrayColNumber = -1;
        public int ProgrammedAtColNumber = -1;
        public int SupplyColNumber = -1;

        public EfuseArraySizeSheet(string sheetName)
        {
            SheetName = sheetName;
        }
    }

    public class EfuseArraySizeRow : MyRow
    {
        public string FuseArray = "";
        public string OfBits = "";
        public string OfFuseArray = "";
        public string ProgrammedAt = "";
        public string Supply = "";
        public List<string> SupplyPinList = [];
    }
}
