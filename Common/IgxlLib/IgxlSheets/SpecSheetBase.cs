using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using CommonLib.Extension;

using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public abstract class SpecSheetBase<T> : IgxlSheet<T> where T : IgxlRow
    {
        public Dictionary<string, List<string>> TimeSetMapping = new(StringExtensions.IgnoreCase);
        public Dictionary<string, List<string>> CategoryTimeSetDic = [];

        public List<string> SelectorNameList { get; set; } = [];
        public List<string> CategoryList { get; set; } = [];

        protected SpecSheetBase(string sheetName) : base(sheetName)
        {
        }

        protected SpecSheetBase(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        protected SpecSheetBase(ExcelWorksheet excelWorksheet, List<string> categoryList, List<string> selectorNameList) : base(excelWorksheet)
        {
            CategoryList = categoryList;
            SelectorNameList = selectorNameList;
        }

        protected SpecSheetBase(string sheetName, List<string> categoryList, List<string> selectorNameList) : base(sheetName)
        {
            CategoryList = categoryList;
            SelectorNameList = selectorNameList;
        }

        public override void Write(string fileName, string version = "")
        {
            if (string.IsNullOrEmpty(version))
            {
                version = "3.0";
            }

            //Support 3.0 & 2.0
            Directory.CreateDirectory(Path.GetDirectoryName(fileName) ?? string.Empty);

            Dictionary<string, Dictionary<string, SheetInfo>> sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.TryGetValue(IgxlSheetName, out Dictionary<string, SheetInfo>? dic))
            {
                if (dic.TryGetValue(version, out SheetInfo? igxlSheetsVersion1))
                {
                    WriteSheet(Rows.OfType<Spec>(), fileName, version, igxlSheetsVersion1);
                }
                else if (dic.TryGetValue("3.0", out SheetInfo? igxlSheetsVersion))
                {
                    WriteSheet(Rows.OfType<Spec>(), fileName, version, igxlSheetsVersion);
                }
            }
        }

        protected virtual void WriteSheet(IEnumerable<Spec> specs, string fileName, string version, SheetInfo sheetInfo)
        {
            IEnumerable<Spec> enumerable = [.. specs];
            if (!enumerable.Any())
            {
                return;
            }

            using var sw = new StreamWriter(fileName, false);
            var sb = new StringBuilder();
            int maxCount = GetMaxCount(sheetInfo);
            int symbolIndex = GetIndexFrom(sheetInfo, "Symbol");
            int valueIndex = GetIndexFrom(sheetInfo, "Value");
            int selectorsIndex = GetIndexFrom(sheetInfo, "Selectors");
            if (selectorsIndex == -1)
            {
                selectorsIndex = GetIndexFrom(sheetInfo, "Selector");
            }

            int categoryValuesIndex = selectorsIndex + SelectorNameList.Count;
            int relativeColumnIndex = selectorsIndex + SelectorNameList.Count + (enumerable.First().CategoryList.Count * 3);
            maxCount = Math.Max(maxCount, relativeColumnIndex + 1);

            bool addBlankRow = false;

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];

                SetField(sheetInfo, i, arr);

                SetColumns(sheetInfo, i, arr);

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

            #region data
            for (int index = 0; index < enumerable.Count(); index++)
            {
                Spec row = enumerable.ElementAt(index);
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                if (!string.IsNullOrEmpty(row.Symbol))
                {
                    if (!addBlankRow && row.IsBackup)
                    {
                        addBlankRow = true;
                        sb.AppendLine("\t");
                    }

                    arr[0] = row.ColumnA!;
                    arr[symbolIndex] = row.Symbol;
                    arr[valueIndex] = "#N/A";

                    for (int i = 0; i < SelectorNameList.Count; i++)
                    {
                        Selector? selectorName = row.SelectorList.Find(x => x.SelectorName!.EqualsIgnoreCase(SelectorNameList[i]));
                        if (selectorName != null)
                        {
                            arr[selectorsIndex + i] = selectorName.SelectorValue!;
                        }
                        else
                        {
                            arr[selectorsIndex + i] = "";
                        }
                    }
                    int categoryCount = 0;
                    foreach (CategoryInSpec category in row.CategoryList)
                    {
                        arr[categoryValuesIndex + (categoryCount * 3)] = category.Typ!;
                        arr[categoryValuesIndex + (categoryCount * 3) + 1] = category.Min!;
                        arr[categoryValuesIndex + (categoryCount * 3) + 2] = category.Max!;
                        categoryCount++;
                    }
                    arr[relativeColumnIndex] = row.Comment;
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

        public void AddCategoryList(string category)
        {
            IEnumerable<Spec> rows = Rows.OfType<Spec>();
            foreach (Spec spec in rows)
            {
                spec.CategoryList.Add(new CategoryInSpec(category, "0", "0", "0"));
            }
            CategoryList.Add(category);
        }
    }
}
