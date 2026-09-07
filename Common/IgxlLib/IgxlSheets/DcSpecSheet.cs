using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using CommonLib.Extension;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

namespace IgxlLib.IgxlSheets
{
    public class DcSpecSheet : SpecSheetBase<DcSpec>
    {
        public override string SheetType => "DTDCSpecSheet";
        public override string IgxlSheetName => IgxlSheetNames.DcSpec;

        //for UT
        public DcSpecSheet(string sheetName) : base(sheetName)
        {
        }

        public DcSpecSheet(string sheetName, List<string> categoryList, List<string> selectorNameList) : base(sheetName, categoryList, selectorNameList)
        {
        }

        protected override void WriteSheet(IEnumerable<Spec> specs, string fileName, string version, SheetInfo sheetInfo)
        {
            IEnumerable<Spec> enumerable = [.. specs];
            if (!enumerable.Any())
            {
                return;
            }
            var dcSpecs = enumerable.Cast<DcSpec>().ToList();
            using (var sw = new StreamWriter(fileName, false))
            {
                int maxCount = GetMaxCount(sheetInfo);
                int symbolIndex = GetIndexFrom(sheetInfo, "Symbol");
                int valueIndex = GetIndexFrom(sheetInfo, "Value");
                int selectorsIndex = GetIndexFrom(sheetInfo, "Selectors");
                int categoryValuesIndex = selectorsIndex + SelectorNameList.Count;
                int relativeColumnIndex = selectorsIndex + SelectorNameList.Count + (dcSpecs.Max(x => x.CategoryCount) * 3);
                maxCount = Math.Max(maxCount, relativeColumnIndex + 1);

                PrintHeader(version, sheetInfo, sw, maxCount, selectorsIndex, categoryValuesIndex, relativeColumnIndex);

                PrintData(sw, maxCount, symbolIndex, valueIndex, selectorsIndex, categoryValuesIndex, relativeColumnIndex);
            }

            DcSpec? testSetting = dcSpecs.Find(x => !string.IsNullOrEmpty(x.SpecialComment));
            if (testSetting != null)
            {
                AddVersion(fileName, testSetting.SpecialComment);
            }
        }

        private void PrintData(StreamWriter streamWriter, int maxCount, int symbolIndex, int valueIndex, int selectorsIndex, int categoryValuesIndex, int relativeColumnIndex)
        {
            #region data
            for (int index = 0; index < Rows.Count; index++)
            {
                Spec row = Rows.ElementAt(index);
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                if (!string.IsNullOrEmpty(row.Symbol))
                {
                    arr[0] = row.ColumnA ?? "";
                    arr[symbolIndex] = row.Symbol;
                    arr[valueIndex] = "#N/A";
                    for (int i = 0; i < SelectorNameList.Count; i++)
                    {
                        Selector? selectorName = row.SelectorList.Find(x => x.SelectorName!.EqualsIgnoreCase(SelectorNameList[i]));
                        if (selectorName != null)
                        {
                            arr[selectorsIndex + i] = selectorName.SelectorValue ?? "";
                        }
                        else
                        {
                            arr[selectorsIndex + i] = "";
                        }
                    }
                    int categoryCount = 0;
                    foreach (CategoryInSpec category in row.CategoryList)
                    {
                        arr[categoryValuesIndex + (categoryCount * 3)] = category.Typ ?? "";
                        arr[categoryValuesIndex + (categoryCount * 3) + 1] = category.Min ?? "";
                        arr[categoryValuesIndex + (categoryCount * 3) + 2] = category.Max ?? "";
                        categoryCount++;
                    }
                    arr[relativeColumnIndex] = row.Comment;
                }
                else
                {
                    arr = ["\t"];
                }
                streamWriter.WriteLine(string.Join("\t", arr));
            }
            #endregion
        }

        private void PrintHeader(string version, SheetInfo sheetInfo, StreamWriter streamWriter, int maxCount, int selectorsIndex, int categoryValuesIndex, int relativeColumnIndex)
        {
            var sb = new StringBuilder();

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];

                SetField(sheetInfo, i, arr);

                if (sheetInfo.Columns.Column != null)
                {
                    foreach (Column item in sheetInfo.Columns.Column)
                    {
                        if (item.rowIndex == i)
                        {
                            arr[item.indexFrom] = item.columnName;
                        }

                        if (item.Column1 != null)
                        {
                            foreach (Column column1 in item.Column1)
                            {
                                if (column1.rowIndex == i)
                                {
                                    arr[column1.indexFrom] = column1.columnName;
                                }
                            }
                        }
                    }
                }

                #region Set Variant
                if (sheetInfo.Columns.Variant != null)
                {
                    foreach (Column item in sheetInfo.Columns.Variant)
                    {
                        if (item.indexFrom == item.indexTo && item.rowIndex == i)
                        {
                            arr[item.indexFrom] = item.columnName;
                        }

                        if (item.columnName == "Selectors" && item.rowIndex + 1 == i)
                        {
                            for (int index = 0; index < SelectorNameList.Count; index++)
                            {
                                string selectorName = SelectorNameList[index];
                                arr[selectorsIndex + index] = selectorName;
                            }
                        }

                        if (item.variantName == "CategoryValues")
                        {
                            int categoryCount = 0;
                            foreach (string category in CategoryList)
                            {
                                foreach (Column column in item.Column1)
                                {
                                    if (column.indexFrom == 1 && item.rowIndex == i)
                                    {
                                        arr[categoryValuesIndex - 1 + (categoryCount * 3) + column.indexFrom] = category;
                                    }

                                    if (column.rowIndex == i)
                                    {
                                        arr[categoryValuesIndex - 1 + (categoryCount * 3) + column.indexFrom] = column.columnName;
                                    }
                                }
                                categoryCount++;
                            }
                        }
                    }
                }
                #endregion

                SetRelativeColumn(sheetInfo, i, arr, relativeColumnIndex);

                WriteHeader(arr, sb);
            }
            #endregion

            streamWriter.Write(sb.ToString());
        }

        private static void AddVersion(string filename, string version)
        {
            string[] allLine = File.ReadAllLines(filename);
            if (allLine.Length >= 5)
            {
                allLine[4] += version;
                File.WriteAllLines(filename, allLine);
            }
        }
    }
}
