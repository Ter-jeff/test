using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public abstract class SpecSheetReader<TSheet, TSpec> : IgxlSheetReader<TSheet> where TSpec : AcSpec where TSheet : SpecSheetBase<TSpec>
    {
        private const int StartRowIndex = 3;
        private const int EndRowIndex = 4;
        private const int SubStartColumnIndex = 4;
        private const int StartColumnIndex = 7;

        public override TSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            int maxRowCount = excelWorksheet.Dimension.End.Row;
            int maxColumnCount = excelWorksheet.Dimension.End.Column;

            var selectorNameList = new List<string>();
            for (int i = SubStartColumnIndex; i < StartColumnIndex; i++)
            {
                string head2 = GetMergeCellValue(excelWorksheet, StartRowIndex + 1, i).Trim();
                if (!string.IsNullOrEmpty(head2))
                {
                    selectorNameList.Add(head2);
                }
            }

            var categoryList = new List<string>();
            for (int i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                string head = GetMergeCellValue(excelWorksheet, StartRowIndex, i).Trim();
                if (!string.IsNullOrEmpty(head))
                {
                    categoryList.Add(head);
                }
            }

            bool isBackup = false;
            TSheet specSheet = (TSheet)Activator.CreateInstance(typeof(TSheet), excelWorksheet.Name, categoryList, selectorNameList)!;
            for (int i = StartRowIndex + 2; i <= maxRowCount; i++)
            {
                string symbol = GetMergeCellValue(excelWorksheet, i, 2).Trim();
                if (symbol.EqualsIgnoreCase("Using TSet"))
                {
                    specSheet.TimeSetMapping = GetTimeSetMapping(excelWorksheet);
                }

                TSpec row = GetSpecsRow(excelWorksheet, i);
                if (string.IsNullOrEmpty(row.Symbol))
                {
                    isBackup = true;
                    continue;
                }

                row.IsBackup = isBackup;
                specSheet.AddRow(row);
            }
            return specSheet;
        }

        public override TSheet ReadSheet(Stream stream, string sheetName)
        {
            var specSheet = (TSheet)Activator.CreateInstance(typeof(TSheet), sheetName)!;
            bool isBackup = false;
            int i = 1;
            string categoryLine = "";
            var categoryList = new List<string>();
            string selectorLine = "";
            var selectorNameList = new List<string>();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string? line = sr.ReadLine();
                    if (i > EndRowIndex)
                    {
                        TSpec spec = GetSpecsRow(line!, sheetName, i, selectorNameList, categoryLine, selectorLine);
                        if (string.IsNullOrEmpty(spec.Symbol))
                        {
                            isBackup = true;
                            continue;
                        }

                        spec.IsBackup = isBackup;
                        specSheet.AddRow(spec);
                        if (spec.Symbol.EqualsIgnoreCase("Using TSet"))
                        {
                            foreach (CategoryInSpec category in spec.CategoryList)
                            {
                                if (!string.IsNullOrEmpty(category.Typ))
                                {
                                    string timeSet = category.Typ;
                                    if (!specSheet.CategoryTimeSetDic.TryGetValue(category.Name, out List<string>? value))
                                    {
                                        specSheet.CategoryTimeSetDic.Add(category.Name, [timeSet]);
                                    }
                                    else if (!value.ToList().Contains(timeSet, StringExtensions.IgnoreCase))
                                    {
                                        value.Add(timeSet);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (line != null)
                        {
                            string[] arr = line.Split('\t');
                            int maxColumnCount = arr.Length;

                            if (i == StartRowIndex)
                            {
                                for (int col = StartColumnIndex - 1; col < maxColumnCount; col++)
                                {
                                    string value = arr[col];
                                    if (!string.IsNullOrEmpty(value))
                                    {
                                        categoryList.Add(value);
                                        categoryLine = line;
                                    }
                                }
                            }
                            else if (i == StartRowIndex + 1)
                            {
                                for (int col = SubStartColumnIndex - 1; col < StartColumnIndex; col++)
                                {
                                    string value = arr[col];
                                    if (!string.IsNullOrEmpty(value))
                                    {
                                        selectorNameList.Add(value);
                                        selectorLine = line;
                                    }
                                }

                                selectorNameList = [.. selectorNameList.Distinct()];
                                selectorNameList.Remove("Comment");
                            }
                        }
                    }

                    i++;
                }
            }

            specSheet.CategoryList = categoryList;
            specSheet.SelectorNameList = selectorNameList;
            return specSheet;
        }

        private TSpec GetSpecsRow(string line, string sheetName, int row, List<string> selectorNameList, string categoryLine, string selectorLine)
        {
            GetSpecsRow(line, selectorNameList, categoryLine, selectorLine, out string symbol, out string comment, out List<CategoryInSpec> categoryInSpecs, out List<Selector> selectorList);

            var spec = (TSpec)Activator.CreateInstance(typeof(TSpec), symbol, selectorList, "", comment)!;
            spec.RowNum = row;
            spec.SheetName = sheetName;
            foreach (CategoryInSpec categoryInSpec in categoryInSpecs)
            {
                spec.AddCategory(categoryInSpec);
            }

            return spec;
        }

        private Dictionary<string, List<string>> GetTimeSetMapping(ExcelWorksheet excelWorksheet)
        {
            var dic = new Dictionary<string, List<string>>();
            string name = "";
            string typ = "";
            string min = "";
            string max = "";
            for (int i = StartColumnIndex; i < excelWorksheet.Dimension.End.Column; i++)
            {
                string head = GetMergeCellValue(excelWorksheet, StartRowIndex, i).Trim();
                if (!string.IsNullOrEmpty(head))
                {
                    if (!string.IsNullOrEmpty(name))
                    {
                        List<string> timeSets = [.. typ.Split(',')];
                        timeSets.AddRange([.. min.Split(',')]);
                        timeSets.AddRange([.. max.Split(',')]);
                        timeSets = [.. timeSets.Distinct().Where(x => !string.IsNullOrEmpty(x))];
                        foreach (string timeSet in timeSets)
                        {
                            if (!dic.TryGetValue(timeSet, out List<string>? value))
                            {
                                dic.Add(timeSet, [name]);
                            }
                            else
                            {
                                value.Add(name);
                            }
                        }
                    }
                    name = head;
                }
            }

            return dic;
        }

        private TSpec GetSpecsRow(ExcelWorksheet excelWorksheet, int row)
        {
            GetSpecsRow(excelWorksheet, row, out string symbol, out string comment, out List<CategoryInSpec> categoryInSpecs, out List<Selector> selectorList);

            TSpec specs = (TSpec)Activator.CreateInstance(typeof(TSpec), symbol, selectorList, "", comment)!;
            foreach (CategoryInSpec categoryInSpec in categoryInSpecs)
            {
                specs.AddCategory(categoryInSpec);
            }
            specs.Comment = comment;

            return specs;
        }

        protected void GetSpecsRow(string line, List<string> selectorNameList, string categoryLine, string selectorLine, out string symbol, out string comment, out List<CategoryInSpec> categoryInSpecs, out List<Selector> selectors)
        {
            string[] arr = line.Split('\t');
            string[] arrCategory = categoryLine.Split('\t');
            string[] arrSelector = selectorLine.Split('\t');

            symbol = arr[1].Trim();
            string name = "";
            comment = "";
            string typ = "";
            string min = "";
            string max = "";
            categoryInSpecs = [];
            selectors = [];
            int index = arrCategory.ToList().IndexOf("Selectors");
            int start = index + 1;
            for (int i = start + selectorNameList.Count - 1; i < arr.Length; i++)
            {
                if (i < arrCategory.Length)
                {
                    string category = arrCategory[i].Trim();
                    if (!string.IsNullOrEmpty(category))
                    {
                        if (!string.IsNullOrEmpty(name))
                        {
                            categoryInSpecs.Add(new CategoryInSpec(name, typ, min, max));
                        }

                        name = category;
                    }
                }

                string selectorName = arrSelector[i].Trim();
                string value = arr[i].Trim();
                switch (CellDiff.FormatStringForCompare(selectorName))
                {
                    case "TYP":
                        typ = value;
                        selectors.Add(new Selector("Typ", "Typ"));
                        break;
                    case "MIN":
                        min = value;
                        selectors.Add(new Selector("Min", "Min"));
                        break;
                    case "MAX":
                        max = value;
                        selectors.Add(new Selector("Max", "Max"));
                        break;
                    case "COMMENT":
                        comment = value;
                        break;
                }
            }

            categoryInSpecs.Add(new CategoryInSpec(name, typ, min, max));
        }

        protected void GetSpecsRow(ExcelWorksheet excelWorksheet, int row, out string symbol, out string comment, out List<CategoryInSpec> categoryInSpecs, out List<Selector> selectors)
        {
            symbol = GetMergeCellValue(excelWorksheet, row, 2).Trim();
            string name = "";
            comment = "";
            string typ = "";
            string min = "";
            string max = "";
            categoryInSpecs = [];
            selectors = [];
            for (int i = StartColumnIndex; i <= excelWorksheet.Dimension.End.Column; i++)
            {
                string head = GetMergeCellValue(excelWorksheet, StartRowIndex, i).Trim();
                if (!string.IsNullOrEmpty(head))
                {
                    if (!string.IsNullOrEmpty(name))
                    {
                        categoryInSpecs.Add(new CategoryInSpec(name, typ, min, max));
                    }

                    name = head;
                }
                string head2 = GetMergeCellValue(excelWorksheet, StartRowIndex + 1, i).Trim();
                string content = GetCellText(excelWorksheet, row, i);
                switch (CellDiff.FormatStringForCompare(head2))
                {
                    case "TYP":
                        typ = content;
                        selectors.Add(new Selector("Typ", "Typ"));
                        break;
                    case "MIN":
                        min = content;
                        selectors.Add(new Selector("Min", "Min"));
                        break;
                    case "MAX":
                        max = content;
                        selectors.Add(new Selector("Max", "Max"));
                        break;
                    case "COMMENT":
                        comment = content;
                        break;
                }
            }

            categoryInSpecs.Add(new CategoryInSpec(name, typ, min, max));
        }
    }
}
