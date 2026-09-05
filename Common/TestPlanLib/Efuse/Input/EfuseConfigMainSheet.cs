using System.Collections.Generic;

using CommonReaderLib;

namespace TestPlanLib.Efuse.Input
{
    public class EfuseConfigMainSheet : MySheet
    {
        public string AccessMode = "";
        public string Header = "";
        public bool HasSvm = false;
        public List<EfuseConfigMainRow> Rows = [];
        public int FuseColNumber = -1;
        public int DescriptionColNumber = -1;
        public int DataColCount = -1;
        public int DataStartCol = -1;
        public int ValueToFuseColNumber = -1;
        public int FuseBlowLocationColNumber = -1;

        public EfuseConfigMainSheet(string sheetName)
        {
            SheetName = sheetName;
        }
    }

    public class EfuseConfigMainRow : MyRow
    {
        public string Fuse = "";
        //format LSB:MSB is default
        public bool NeedSort = true;
        public int Lsb;
        public int Msb;
        public int Width;
        public string Description = "";
        public List<string> DataList = [];
        public string ValueToFuse = "";
        public string FuseBlowLocation = "";
        public bool IsAllSameData = false;
    }
}
