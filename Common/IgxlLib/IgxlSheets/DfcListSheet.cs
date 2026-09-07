using System.IO;

using IgxlLib.IgxlBase;

namespace IgxlLib.IgxlSheets
{
    public class DfcListSheet(string sheetName) : IgxlSheet<DfcRow>(sheetName)
    {
        public override string SheetType => "Dfc_List";
        public override string IgxlSheetName => "DFC_List";

        public override void Write(string fileName, string version = "")
        {
            using var sw = new StreamWriter(fileName);
            foreach (DfcRow row in Rows)
            {
                sw.WriteLine(row.Text);
            }
        }
    }
}
