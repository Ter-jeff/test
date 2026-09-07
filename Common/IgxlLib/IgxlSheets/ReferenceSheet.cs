using System.IO;
using System.Text;

using IgxlLib.Const;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class ReferenceSheet : IgxlSheet<ReferenceRow>
    {
        public override string SheetType => "DTReferencesSheet";
        public override string IgxlSheetName => IgxlSheetNames.Reference;

        public ReferenceSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public ReferenceSheet(string sheetName) : base(sheetName)
        {
        }

        private StringBuilder WriteContent(string version = "")
        {
            return WriteContent2P0(version);
        }

        public override void Write(string fileName, string version = "")
        {
            Directory.CreateDirectory(Path.GetDirectoryName(fileName) ?? "");

            if (string.IsNullOrEmpty(version))
            {
                version = "2.0";
            }

            using var sw = new StreamWriter(fileName, false);
            StringBuilder content = WriteContent(version);
            sw.Write(content.ToString());
        }
    }
}
