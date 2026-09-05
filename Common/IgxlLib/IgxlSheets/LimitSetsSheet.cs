using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class LimitSetsSheet : IgxlSheet<LimitSetsRow>
    {
        public override string SheetType => "LimitSetsSheet";
        public override string IgxlSheetName => IgxlSheetNames.LimitSetsSheet;

        public string SourceSheet = string.Empty;

        public LimitSetsSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
            if (excelWorksheet.Name.Length > 31)
            {
                throw new Exception($"{excelWorksheet.Name} sheet name exceeds the maximum length");
            }
        }

        public LimitSetsSheet(string sheetName) : base(sheetName)
        {
            if (sheetName.Length > 31)
            {
                throw new Exception($"{sheetName} sheet name exceeds the maximum length");
            }
        }

        public override void Write(string fileName, string version = "")
        {
            if (string.IsNullOrEmpty(version))
            {
                version = "2.0";
            }

            WriteByOneVersion(fileName, WriteContent(version));
        }

        public StringBuilder? WriteContent(string version = "")
        {
            Dictionary<string, Dictionary<string, SheetInfo>> sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.TryGetValue(IgxlSheetName, out Dictionary<string, SheetInfo>? dic))
            {
                if (dic.TryGetValue(version, out SheetInfo? igxlSheetsVersion1))
                {
                    return WriteSheet(version, igxlSheetsVersion1);
                }
                if (dic.TryGetValue("2.0", out SheetInfo? igxlSheetsVersion))
                {
                    return WriteSheet(version, igxlSheetsVersion);
                }
            }
            return null;
        }

        protected StringBuilder? WriteSheet(string version, SheetInfo sheetInfo)
        {
            if (Rows.Count == 0)
            {
                return null;
            }

            var sb = new StringBuilder();
            int maxCount = GetMaxCount(sheetInfo);
            int rowIndex = GetIndexFrom(sheetInfo, "row");
            int flowTableIndex = GetIndexFrom(sheetInfo, "flowtable");
            int instanceIndex = GetIndexFrom(sheetInfo, "instance");
            int testNameIndex = GetIndexFrom(sheetInfo, "testname");
            int testNumberIndex = GetIndexFrom(sheetInfo, "testnumber");
            int lowCompSignIndex = GetIndexFrom(sheetInfo, "lowcompsign");
            int highCompSignIndex = GetIndexFrom(sheetInfo, "highcompsign");

            var loIndexes = new List<int>();
            var hiIndexes = new List<int>();
            for (int ind = 0; ind < Rows.First().LimitSetsCat.Count; ind++)
            {
                loIndexes.Add(GetIndexFrom(sheetInfo, "Name", "lolim") + (ind * 2));
                hiIndexes.Add(GetIndexFrom(sheetInfo, "Name", "hilim") + (ind * 2));
            }

            int categoryNum = highCompSignIndex + (Rows.First().LimitSetsCat.Count * 2);
            int scaleIndex = categoryNum + GetIndexFrom(sheetInfo, "scale");
            int unitsIndex = categoryNum + GetIndexFrom(sheetInfo, "units");
            int formatIndex = categoryNum + GetIndexFrom(sheetInfo, "format");
            int passSortIndex = categoryNum + GetIndexFrom(sheetInfo, "passsort");
            int failSortIndex = categoryNum + GetIndexFrom(sheetInfo, "failsort");
            int sortNameIndex = categoryNum + GetIndexFrom(sheetInfo, "sortname");
            int passBinIndex = categoryNum + GetIndexFrom(sheetInfo, "passbin");
            int failBinIndex = categoryNum + GetIndexFrom(sheetInfo, "failbin");
            int binNameIndex = categoryNum + GetIndexFrom(sheetInfo, "binname");
            int resultIndex = categoryNum + GetIndexFrom(sheetInfo, "result");
            int passActionIndex = categoryNum + GetIndexFrom(sheetInfo, "passaction");
            int failActionIndex = categoryNum + GetIndexFrom(sheetInfo, "failaction");
            int assumeIndex = categoryNum + GetIndexFrom(sheetInfo, "assume");
            int sitesIndex = categoryNum + GetIndexFrom(sheetInfo, "sites");
            int commentIndex = categoryNum + GetIndexFrom(sheetInfo, "comment");

            maxCount = Math.Max(maxCount, commentIndex + 1);

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            sb.AppendLine(SheetType + ",version=" + version);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];

                SetField(sheetInfo, i, arr);

                SetColumns(sheetInfo, i, arr);

                if (sheetInfo.Columns.Variant != null)
                {
                    foreach (Column item in sheetInfo.Columns.Variant)
                    {
                        for (int ind = 0; ind < Rows.First().LimitSetsCat.Count; ind++)
                        {
                            if (item.columnName == "Name" && item.rowIndex == i)
                            {
                                arr[item.indexFrom + (ind * 2)] = Rows.First().LimitSetsCat[ind].CategoryName;
                                arr[item.indexTo + (ind * 2)] = Rows.First().LimitSetsCat[ind].CategoryName;
                            }

                            foreach (Column column in item.Column1)
                            {
                                if (column.indexTo == column.indexFrom && column.rowIndex == i)
                                {
                                    arr[column.indexFrom + (ind * 2)] = column.columnName;
                                }
                            }
                        }
                    }
                }

                SetRelativeColumn(sheetInfo, i, arr, scaleIndex);

                WriteHeader(arr, sb);
            }
            #endregion

            #region data
            var limitSetsRows = Rows.Where(x => !x.IsBackup).ToList();
            WriteLimitSetsRows(limitSetsRows, maxCount, rowIndex, flowTableIndex, instanceIndex, testNameIndex, testNumberIndex, lowCompSignIndex, highCompSignIndex, loIndexes, hiIndexes, scaleIndex, unitsIndex, formatIndex, passSortIndex, failSortIndex, sortNameIndex, passBinIndex, failBinIndex, binNameIndex, resultIndex, passActionIndex, failActionIndex, assumeIndex, sitesIndex, commentIndex, sb);

            var backupRows = Rows.Where(x => x.IsBackup).ToList();
            if (backupRows.Count != 0)
            {
                PrintBackupEmptyRows(sb);
                WriteLimitSetsRows(backupRows, maxCount, rowIndex, flowTableIndex, instanceIndex, testNameIndex, testNumberIndex, lowCompSignIndex, highCompSignIndex, loIndexes, hiIndexes, scaleIndex, unitsIndex, formatIndex, passSortIndex, failSortIndex, sortNameIndex, passBinIndex, failBinIndex, binNameIndex, resultIndex, passActionIndex, failActionIndex, assumeIndex, sitesIndex, commentIndex, sb);
            }
            #endregion

            return sb;
        }

        private static void WriteLimitSetsRows(List<LimitSetsRow> limitSetsRows, int maxCount, int rowIndex, int flowTableIndex, int instanceIndex, int testNameIndex, int testNumberIndex, int lowCompSignIndex, int highCompSignIndex, List<int> loIndexes, List<int> hiIndexes, int scaleIndex, int unitsIndex, int formatIndex, int passSortIndex, int failSortIndex, int sortNameIndex, int passBinIndex, int failBinIndex, int binNameIndex, int resultIndex, int passActionIndex, int failActionIndex, int assumeIndex, int sitesIndex, int commentIndex, StringBuilder stringBuilder)
        {
            for (int index = 0; index < limitSetsRows.Count; index++)
            {
                LimitSetsRow row = limitSetsRows[index];
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                arr[0] = row.ColumnA ?? "";
                arr[rowIndex] = row.Row;
                arr[flowTableIndex] = row.FlowTable;
                arr[instanceIndex] = row.Instance;
                arr[testNameIndex] = row.TestName;
                arr[testNumberIndex] = row.TestNumber;
                arr[lowCompSignIndex] = row.LowCompSign;
                arr[highCompSignIndex] = row.HighCompSign;
                for (int i = 0; i < row.LimitSetsCat.Count; i++)
                {
                    arr[loIndexes[i]] = row.LimitSetsCat[i].LoLim;
                    arr[hiIndexes[i]] = row.LimitSetsCat[i].HiLim;
                }
                arr[scaleIndex] = row.Scale;
                arr[unitsIndex] = row.Units;
                arr[formatIndex] = row.Format;
                arr[passSortIndex] = row.PassSort;
                arr[failSortIndex] = row.FailSort;
                arr[sortNameIndex] = row.SortName;
                arr[passBinIndex] = row.PassBin;
                arr[failBinIndex] = row.FailBin;
                arr[binNameIndex] = row.BinName;
                arr[resultIndex] = row.Result;
                arr[passActionIndex] = row.PassAction;
                arr[failActionIndex] = row.FailAction;
                arr[assumeIndex] = row.Assume;
                arr[sitesIndex] = row.Sites;
                arr[commentIndex] = row.Comment;

                stringBuilder.AppendLine(string.Join("\t", arr) + "\t");
            }
        }
    }
}
