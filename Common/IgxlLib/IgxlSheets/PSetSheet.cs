using System.IO;
using System.Linq;
using System.Text;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class PSetSheet : IgxlSheet<PSet>
    {
        public override string SheetType => "DTPsetsSheet";
        public override string IgxlSheetName => IgxlSheetNames.PSet;

        public PSetSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public PSetSheet(string sheetName) : base(sheetName)
        {
        }

        protected override void WriteSheet(string fileName, string version, SheetInfo sheetInfo)
        {
            if (Rows.Count == 0)
            {
                return;
            }

            using var sw = new StreamWriter(fileName, false);
            var sb = new StringBuilder();
            int maxCount = GetMaxCount(sheetInfo);
            int nameIndex = GetIndexFrom(sheetInfo, "Name");
            int pinIndex = GetIndexFrom(sheetInfo, "Pin");
            int instrumentTypeIndex = GetIndexFrom(sheetInfo, "Instrument Type");
            int thisLargeHeadingIndex = GetIndexFrom(sheetInfo, "This large heading makes AutoFit enlarge the row height");
            int parameter0Index = GetIndexFrom(sheetInfo, "Parameter0");
            int commentIndex = GetIndexFrom(sheetInfo, "Comment");

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];

                SetField(sheetInfo, i, arr);

                SetColumns(sheetInfo, i, arr);

                WriteHeader(arr, sb);
            }
            #endregion

            #region data
            for (int index = 0; index < Rows.Count; index++)
            {
                PSet row = Rows[index];
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                if (!string.IsNullOrEmpty(row.Name))
                {
                    arr[0] = row.ColumnA ?? "";
                    arr[nameIndex] = row.Name;
                    arr[pinIndex] = row.Pin;
                    arr[instrumentTypeIndex] = row.InstrumentType;
                    arr[thisLargeHeadingIndex] = row.ThisLargeHeading;
                    for (int j = 0; j < row.Parameters.Count; j++)
                    {
                        if (j < 80)
                        {
                            arr[parameter0Index + j] = row.Parameters.ElementAt(j).Value;
                        }
                    }
                    arr[commentIndex] = row.Comment;
                }
                else
                {
                    arr = ["\t"];
                }
                sb.AppendLine(string.Join("\t", arr));
            }
            #endregion

            sw.Write(sb.ToString());
        }

        public override void Write(string fileName, string version = "")
        {
            Write2P0(fileName, version);
        }
    }
}
