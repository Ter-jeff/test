using System.IO;

using IgxlLib.IgxlBase;

namespace IgxlLib.IgxlSheets
{
    public class BasFile(string sheetName) : IgxlSheet<BasRow>(sheetName)
    {
        public override string SheetType => "BasFile";
        public override string IgxlSheetName => "BasFile";

        public override void Write(string fileName, string version = "")
        {
            using var sw = new StreamWriter(fileName);
            foreach (BasRow row in Rows)
            {
                sw.WriteLine(row.Text);
            }
        }
    }
}
