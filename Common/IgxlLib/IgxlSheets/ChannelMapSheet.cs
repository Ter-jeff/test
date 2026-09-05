using System;
using System.IO;
using System.Linq;
using System.Text;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class ChannelMapSheet : IgxlSheet<ChannelMapRow>
    {
        public override string SheetType => "DTChanMap";
        public override string IgxlSheetName => IgxlSheetNames.ChannelMap;

        public bool IsPogo { get; set; }

        public int SiteNum { get; set; }

        public ChannelMapSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
            SiteNum = 1;
        }

        public ChannelMapSheet(string sheetName) : base(sheetName)
        {
            SiteNum = 1;
        }

        public override void Write(string fileName, string version = "")
        {
            Write2P4(fileName, version);
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
            int pinNameIndex = GetIndexFrom(sheetInfo, "Device Under Test", "Pin Name");
            int packagePinIndex = GetIndexFrom(sheetInfo, "Device Under Test", "Package Pin");
            int typeIndex = GetIndexFrom(sheetInfo, "Tester Channel", "Type");
            int siteIndex = GetIndexFrom(sheetInfo, "Site");
            int siteCount = Rows.Max(p => p.Sites.Count);
            int relativeColumnIndex = siteIndex + siteCount;
            maxCount = Math.Max(maxCount, relativeColumnIndex + 1);

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            if (version == "2.4")
            {
                sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            }
            else
            {
                sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1:dataformat=signal\t" + IgxlSheetName);
            }

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
                        if (item.rowIndex == i)
                        {
                            arr[item.indexFrom] = item.columnName;
                        }

                        if (item.columnName == "Site" && item.rowIndex == i)
                        {
                            for (int index = 0; index < siteCount; index++)
                            {
                                arr[item.indexFrom + index] = "Site " + index;
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
            for (int j = 0; j < Rows.Count; j++)
            {
                ChannelMapRow row = Rows[j];
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                if (!string.IsNullOrEmpty(row.DeviceUnderTestPinName))
                {
                    arr[0] = row.ColumnA ?? "";
                    arr[pinNameIndex] = row.DeviceUnderTestPinName;
                    arr[packagePinIndex] = row.DeviceUnderTestPackagePin;
                    arr[typeIndex] = row.Type;
                    for (int i = 0; i < row.Sites.Count; i++)
                    {
                        string site = row.Sites[i];
                        arr[siteIndex + i] = site;
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
    }
}
