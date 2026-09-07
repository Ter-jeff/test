using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using CommonLib.Extension;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class GlobalSpecSheet : IgxlSheet<GlobalSpec>
    {
        public override string SheetType => "DTGlobalSpecSheet";
        public override string IgxlSheetName => IgxlSheetNames.GlobalSpec;

        public GlobalSpecSheet(ExcelWorksheet excelWorksheet, bool isAddDefault = true) : base(excelWorksheet)
        {
            if (isAddDefault)
            {
                AddDefaultRow();
            }
        }

        public GlobalSpecSheet(string sheetName, bool isAddDefault = true) : base(sheetName)
        {
            if (isAddDefault)
            {
                AddDefaultRow();
            }
        }

        public override void AddRow(GlobalSpec globalSpec)
        {
            if (!Rows.Exists(p => p.Symbol.EqualsIgnoreCase(globalSpec.Symbol) && p.Job.EqualsIgnoreCase(globalSpec.Job)))
            {
                Rows.Add(globalSpec);
            }
        }

        public List<GlobalSpec> FindRowList(string globalSpecSymbol, string globalSpecJob)
        {
            List<GlobalSpec> foundGlobalSpecs = Rows.FindAll(g => g.Symbol.EqualsIgnoreCase(globalSpecSymbol) && g.Job.EqualsIgnoreCase(globalSpecJob));
            return foundGlobalSpecs;
        }

        private void AddDefaultRow()
        {
            var vclDefault = new GlobalSpec("Vcl_default", "", "=-1", "Detector clamp voltage low");
            AddRow(vclDefault);
            var vchDefault = new GlobalSpec("Vch_default", "", "=6", "Detector clamp voltage high");
            AddRow(vchDefault);
            var vphDefault = new GlobalSpec("Vph_default", "", "=5", "Hi-V pin voltage high");
            AddRow(vphDefault);
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
            int symbolIndex = GetIndexFrom(sheetInfo, "Symbol");
            int jobIndex = GetIndexFrom(sheetInfo, "Job");
            int valueIndex = GetIndexFrom(sheetInfo, "Value");
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
                GlobalSpec row = Rows[index];
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                if (!string.IsNullOrEmpty(row.Symbol))
                {
                    arr[0] = row.ColumnA ?? "";
                    arr[symbolIndex] = row.Symbol;
                    arr[jobIndex] = row.Job;
                    arr[valueIndex] = row.Value;
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
