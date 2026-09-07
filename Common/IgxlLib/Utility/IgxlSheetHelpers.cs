using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Serialization;

using CommonLib.Extension;

using IgxlLib.IGDataXML.IGXLSheetsVersion;

namespace IgxlLib.Utility
{
    public static class IgxlSheetHelpers
    {
        public static Dictionary<string, Dictionary<string, SheetInfo>> GetIgxlSheetsVersion()
        {
            var assembly = Assembly.GetExecutingAssembly();
            string[] resourceNames = assembly.GetManifestResourceNames();
            var igxlConfigDic = new Dictionary<string, Dictionary<string, SheetInfo>>();
            foreach (string resourceName in resourceNames)
            {
                if (resourceName.Contains(".IGXLSheetsVersion."))
                {
                    using Stream? stream = assembly.GetManifestResourceStream(resourceName);
                    var xs = new XmlSerializer(typeof(IGXLVersion));
                    if (stream != null)
                    {
                        var igxlConfig = (IGXLVersion?)xs.Deserialize(stream);
                        foreach (SheetInfo sheetItemClass in igxlConfig!.Sheets)
                        {
                            string sheetName = sheetItemClass.sheetName;
                            string sheetVersion = sheetItemClass.sheetVersion;
                            var dic = new Dictionary<string, SheetInfo>();
                            if (dic.TryAdd(sheetVersion, sheetItemClass))
                            {
                                if (!igxlConfigDic.TryAdd(sheetName, dic))
                                {
                                    igxlConfigDic[sheetName].TryAdd(sheetVersion, sheetItemClass);
                                }
                            }
                        }
                    }
                }
            }
            return igxlConfigDic;
        }

        public static int GetIndexFrom(SheetInfo sheetInfo, string name)
        {
            if (sheetInfo.Field != null)
            {
                foreach (Field item in sheetInfo.Field)
                {
                    if (item.fieldName.EqualsIgnoreCase(name))
                    {
                        return item.columnIndex;
                    }
                }
            }

            if (sheetInfo.Columns.Column != null)
            {
                foreach (Column item in sheetInfo.Columns.Column)
                {
                    if (item.columnName.EqualsIgnoreCase(name))
                    {
                        return item.indexFrom;
                    }
                }
            }

            if (sheetInfo.Columns.Variant != null)
            {
                foreach (Column item in sheetInfo.Columns.Variant)
                {
                    if (item.columnName.EqualsIgnoreCase(name))
                    {
                        return item.indexFrom;
                    }
                }
            }

            if (sheetInfo.Columns.RelativeColumn != null)
            {
                foreach (Column item in sheetInfo.Columns.RelativeColumn)
                {
                    if (item.columnName.EqualsIgnoreCase(name))
                    {
                        return item.indexFrom;
                    }
                }
            }
            return -1;
        }

        public static int GetIndexFrom(SheetInfo sheetInfo, string name, string subString)
        {
            if (sheetInfo.Columns.Column != null)
            {
                foreach (Column item in sheetInfo.Columns.Column)
                {
                    if (item.columnName.EqualsIgnoreCase(name))
                    {
                        if (string.IsNullOrEmpty(subString))
                        {
                            return item.indexFrom;
                        }

                        if (item.Column1 != null)
                        {
                            foreach (Column column1 in item.Column1)
                            {
                                if (column1.columnName.EqualsIgnoreCase(subString))
                                {
                                    return column1.indexFrom;
                                }
                            }
                        }
                    }
                }
            }

            if (sheetInfo.Columns.Variant != null)
            {
                foreach (Column item in sheetInfo.Columns.Variant)
                {
                    if (item.columnName.EqualsIgnoreCase(name))
                    {
                        if (string.IsNullOrEmpty(subString))
                        {
                            return item.indexFrom;
                        }

                        if (item.Column1 != null)
                        {
                            foreach (Column column1 in item.Column1)
                            {
                                if (column1.columnName.EqualsIgnoreCase(subString))
                                {
                                    return column1.indexFrom;
                                }
                            }
                        }
                    }
                }
            }

            return -1;
        }

        public static int GetMaxCount(SheetInfo sheetInfo)
        {
            int max = -1;
            if (sheetInfo.Field != null)
            {
                int fieldMax = sheetInfo.Field.Max(x => x.columnIndex);
                if (fieldMax > max)
                {
                    max = fieldMax;
                }
            }

            if (sheetInfo.Columns.Column != null)
            {
                foreach (Column item in sheetInfo.Columns.Column)
                {
                    max = Math.Max(max, item.indexFrom);
                    if (item.Column1 != null)
                    {
                        foreach (Column column1 in item.Column1)
                        {
                            max = Math.Max(max, column1.indexFrom);
                        }
                    }
                }
            }

            if (sheetInfo.Columns.Variant != null)
            {
                foreach (Column item in sheetInfo.Columns.Variant)
                {
                    max = Math.Max(max, item.indexFrom);
                    if (item.Column1 != null)
                    {
                        foreach (Column column1 in item.Column1)
                        {
                            max = Math.Max(max, column1.indexFrom);
                        }
                    }
                }
            }

            if (sheetInfo.Columns.RelativeColumn != null)
            {
                int relativeColumnMax = sheetInfo.Columns.RelativeColumn.Max(x => x.indexFrom);
                if (relativeColumnMax > max)
                {
                    max = relativeColumnMax;
                }
            }

            return max + 1;
        }

        public static void SetColumns(SheetInfo sheetInfo, int i, string[] arr)
        {
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
        }

        public static void SetField(SheetInfo sheetInfo, int i, string[] arr)
        {
            if (sheetInfo.Field != null)
            {
                foreach (Field item in sheetInfo.Field)
                {
                    if (item.rowIndex == i)
                    {
                        arr[item.columnIndex] = item.fieldName;
                    }
                }
            }
        }

        public static void WriteHeader(string[] arr, StringBuilder stringBuilder)
        {
            stringBuilder.AppendLine(WriteHeader(arr));
        }

        public static string WriteHeader(string[] arr)
        {
            if (arr == null || arr.Length == 0)
            {
                return "\t";
            }

            if (arr.Any(x => !string.IsNullOrEmpty(x)))
            {
                return string.Join("\t", arr).TrimEnd('\t');
            }

            return "\t";
        }
    }
}
