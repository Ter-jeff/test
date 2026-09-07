using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public partial class BinTableSheet : IgxlSheet<BinTableRow>
    {
        [GeneratedRegex(@"^Item(?<number>[\d]+)", RegexOptions.Compiled)]
        private static partial Regex ItemNumberRegex();
        private static readonly Regex _itemNumberRegex = ItemNumberRegex();

        public override string SheetType => "DTBintablesSheet";
        public override string IgxlSheetName => IgxlSheetNames.BinTable;

        public BinTableSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public BinTableSheet(string sheetName) : base(sheetName)
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
            int itemListIndex = GetIndexFrom(sheetInfo, "Item List");
            int opIndex = GetIndexFrom(sheetInfo, "Op");
            int sortIndex = GetIndexFrom(sheetInfo, "Sort");
            int binIndex = GetIndexFrom(sheetInfo, "Bin");
            int resultIndex = GetIndexFrom(sheetInfo, "Result");
            int item0Index = GetIndexFrom(sheetInfo, "Item0");
            int commentIndex = GetIndexFrom(sheetInfo, "Comment");

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                foreach (Column item in sheetInfo.Columns.Column)
                {
                    if (item.indexFrom == item.indexTo && item.rowIndex == i)
                    {
                        if (_itemNumberRegex.IsMatch(item.columnName))
                        {
                            arr[item.indexFrom] = item.columnName.Replace("Item", "Items");
                        }
                        else
                        {
                            arr[item.indexFrom] = item.columnName;
                        }
                    }
                }

                WriteHeader(arr, sb);
            }
            #endregion

            #region data
            List<BinTableRow> mainList = Rows.FindAll(x => !x.IsBackup);
            List<BinTableRow> backupList = Rows.FindAll(x => x.IsBackup && !string.IsNullOrEmpty(x.Name));
            if (backupList.Count != 0)
            {
                PrintBackupEmptyRows(mainList);
                mainList.AddRange(backupList);
            }
            for (int index = 0; index < mainList.Count; index++)
            {
                BinTableRow row = mainList[index];
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                if (!string.IsNullOrEmpty(row.Name))
                {
                    arr[0] = row.ColumnA ?? "";
                    arr[nameIndex] = row.Name;
                    arr[itemListIndex] = string.Join(",", row.ItemList);
                    arr[opIndex] = row.Op;
                    arr[sortIndex] = row.Sort;
                    arr[binIndex] = row.Bin;
                    arr[resultIndex] = row.Result;
                    for (int i = 0; i < row.Items.Count; i++)
                    {
                        string item = row.Items[i];
                        arr[item0Index + i] = item;
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

        public void Check()
        {
            CheckFlagMismatch();
        }

        private void CheckFlagMismatch()
        {
            foreach (BinTableRow row in Rows)
            {
                if (row.Items.FindLastIndex(x => !string.IsNullOrEmpty(x)) >= row.ItemList.Split(',').Length)
                {
                    throw new Exception($"BintTable {row.Name} the last target index exceed flag count !!!");
                }
            }
        }

        public override void Write(string fileName, string version = "")
        {
            Write2P0(fileName, version);
        }
    }
}
