using System.IO;
using System.Text;

using CommonLib.Extension;

using IgxlLib.Const;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class PatSetSubSheet : IgxlSheet<PatSetSubRow>
    {
        public override string SheetType => "DTPatternSubroutineSheet";
        public override string IgxlSheetName => IgxlSheetNames.PatternSubroutine;

        public PatSetSubSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public PatSetSubSheet(string sheetName) : base(sheetName)
        {
        }

        public PatSetSubRow? GetExist(string patternFileName)
        {
            return Rows.Find(x => !x.IsBackup && x.PatternFileName.EqualsIgnoreCase(patternFileName));
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
