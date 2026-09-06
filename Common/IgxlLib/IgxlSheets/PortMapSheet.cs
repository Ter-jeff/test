using System.IO;
using System.Linq;
using System.Text;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class PortMapSheet : IgxlSheet<PortSet>
    {
        public override string SheetType => "DTPortMapSheet";
        public override string IgxlSheetName => IgxlSheetNames.PortMap;

        public PortMapSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public PortMapSheet(string sheetName) : base(sheetName)
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
            int portNameIndex = GetIndexFrom(sheetInfo, "Port Name");
            int familyIndex = GetIndexFrom(sheetInfo, "Protocol", "Family");
            int typeIndex = GetIndexFrom(sheetInfo, "Protocol", "Type");
            int settingsIndex = GetIndexFrom(sheetInfo, "Protocol", "Settings");
            int setting0Index = GetIndexFrom(sheetInfo, "Protocol", "Setting0");
            int nameIndex = GetIndexFrom(sheetInfo, "Function", "Name");
            int pinIndex = GetIndexFrom(sheetInfo, "Function", "Pin");
            int propertiesIndex = GetIndexFrom(sheetInfo, "Function", "Properties");
            int property0Index = GetIndexFrom(sheetInfo, "Function", "Property0");
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
                PortSet portSet = Rows[index];
                for (int i = 0; i < portSet.PortRows.Count; i++)
                {
                    PortRow row = portSet.PortRows[i];
                    string[] arr = [.. Enumerable.Repeat("", maxCount)];
                    if (!string.IsNullOrEmpty(portSet.PortName))
                    {
                        arr[0] = row.ColumnA ?? "";
                        arr[portNameIndex] = portSet.PortName;
                        arr[familyIndex] = row.ProtocolFamily;
                        arr[typeIndex] = row.ProtocolType;
                        arr[settingsIndex] = row.ProtocolSettings;
                        for (int j = 0; j < row.ProtocolSettingValues.Count; j++)
                        {
                            if (j < 10)
                            {
                                arr[setting0Index + j] = row.ProtocolSettingValues[j];
                            }
                        }
                        arr[nameIndex] = row.FunctionName;
                        arr[pinIndex] = row.FunctionPin;
                        arr[propertiesIndex] = row.FunctionProperties;
                        for (int j = 0; j < row.FunctionPropertyValues.Count; j++)
                        {
                            if (j < 10)
                            {
                                arr[property0Index + j] = row.FunctionPropertyValues[j];
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
