using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class TimeSetBasicSheet : IgxlSheet<TSet>
    {
        public override string SheetType => "DTTimesetBasicSheet";
        public override string IgxlSheetName { get; } = IgxlSheetNames.TimeSetsBasic;

        public Dictionary<string, double> CommentVariable = [];

        public string TimingMode { get; set; } = string.Empty;
        public string MasterTimeSet { get; set; } = string.Empty;
        public string TimeDomain { get; set; } = string.Empty;
        public string StrobeRefSetup { get; set; } = string.Empty;

        public TimeSetBasicSheet(TimeSetBasicSheet timeSetBasicSheet) : base(timeSetBasicSheet.Name)
        {
            Name = timeSetBasicSheet.Name;
            IgxlSheetName = timeSetBasicSheet.IgxlSheetName;
            FilePath = timeSetBasicSheet.FilePath;
            IgxlSheetContext = timeSetBasicSheet.IgxlSheetContext;
            SourceInfo = timeSetBasicSheet.SourceInfo.Copy();
            TimingMode = timeSetBasicSheet.TimingMode;
            MasterTimeSet = timeSetBasicSheet.MasterTimeSet;
            TimeDomain = timeSetBasicSheet.TimeDomain;
            StrobeRefSetup = timeSetBasicSheet.StrobeRefSetup;
            CommentVariable = timeSetBasicSheet.CommentVariable?.ToDictionary(entry => entry.Key, entry => entry.Value) ?? [];
            Rows = [.. timeSetBasicSheet.Rows.Select(x => x.Copy())];
        }

        public TimeSetBasicSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public TimeSetBasicSheet(string sheetName) : base(sheetName)
        {
        }

        public TimeSetBasicSheet(ExcelWorksheet excelWorksheet, string timingMode) : base(excelWorksheet)
        {
            TimingMode = timingMode;
        }

        public TimeSetBasicSheet(string sheetName, string timingMode) : base(sheetName)
        {
            TimingMode = timingMode;
        }

        public TimeSetBasicSheet(ExcelWorksheet excelWorksheet, string timingMode, string masterTimeSet, string timeDomain, string strobeRefSetup) : base(excelWorksheet)
        {
            TimingMode = timingMode;
            MasterTimeSet = masterTimeSet;
            TimeDomain = timeDomain;
            StrobeRefSetup = strobeRefSetup;
        }

        public TimeSetBasicSheet(string sheetName, string timingMode, string masterTimeSet, string timeDomain, string strobeRefSetup) : base(sheetName)
        {
            TimingMode = timingMode;
            MasterTimeSet = masterTimeSet;
            TimeDomain = timeDomain;
            StrobeRefSetup = strobeRefSetup;
        }

        public override void Write(string fileName, string version = "")
        {
            Write2P3(fileName, version);
        }

        protected override void WriteSheet(string fileName, string version, SheetInfo sheetInfo)
        {
            if (Rows.Count == 0)
            {
                return;
            }

            List<string> lines = PrintContent(version, sheetInfo);

            File.WriteAllLines(fileName, lines);
        }

        public void InsertTimeSet(int index, TSet tSet)
        {
            Rows.Insert(index, tSet);
        }

        public TimeSetBasicSheet Copy()
        {
            return new TimeSetBasicSheet(this);
        }

        protected List<string> PrintContent(string version, SheetInfo sheetInfo)
        {
            var lines = new List<string>();

            int maxCount = GetMaxCount(sheetInfo);
            int timeSetIndex = GetIndexFrom(sheetInfo, "Time Set");
            int periodIndex = GetIndexFrom(sheetInfo, "Cycle", "Period");
            int nameIndex = GetIndexFrom(sheetInfo, "Pin/Group", "Name");
            int clockPeriodIndex = GetIndexFrom(sheetInfo, "Pin/Group", "Clock Period");
            int setupIndex = GetIndexFrom(sheetInfo, "Pin/Group", "Setup");
            int srcIndex = GetIndexFrom(sheetInfo, "Data", "Src");
            int fmtIndex = GetIndexFrom(sheetInfo, "Data", "Fmt");
            int onIndex = GetIndexFrom(sheetInfo, "Drive", "On");
            int dataIndex = GetIndexFrom(sheetInfo, "Drive", "Data");
            int returnIndex = GetIndexFrom(sheetInfo, "Drive", "Return");
            int offIndex = GetIndexFrom(sheetInfo, "Drive", "Off");
            int compareModeIndex = GetIndexFrom(sheetInfo, "Compare", "Mode");
            int compareOpenIndex = GetIndexFrom(sheetInfo, "Compare", "Open");
            int compareCloseIndex = GetIndexFrom(sheetInfo, "Compare", "Close");
            int compareRefOffsetIndex = GetIndexFrom(sheetInfo, "Compare", "Ref Offset");
            int edgeModeIndex = GetIndexFrom(sheetInfo, "Edge Resolution", "Mode");
            int commentIndex = GetIndexFrom(sheetInfo, "Comment");

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            lines.Add(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];

                if (sheetInfo.Field != null)
                {
                    foreach (Field item in sheetInfo.Field)
                    {
                        if (item.rowIndex == i)
                        {
                            arr[item.columnIndex] = item.fieldName;
                            if (item.fieldName.EqualsIgnoreCase("Timing Mode:"))
                            {
                                if (string.IsNullOrEmpty(TimingMode))
                                {
                                    arr[item.columnIndex + 1] = "Single";
                                }
                                else
                                {
                                    arr[item.columnIndex + 1] = TimingMode;
                                }
                            }
                            if (item.fieldName.EqualsIgnoreCase("Master TimeSet Name:"))
                            {
                                arr[item.columnIndex + 1] = MasterTimeSet;
                            }

                            if (item.fieldName.EqualsIgnoreCase("Time Domain:"))
                            {
                                arr[item.columnIndex + 1] = TimeDomain;
                            }

                            if (item.fieldName.EqualsIgnoreCase("Strobe Ref Setup Name:"))
                            {
                                arr[item.columnIndex + 1] = StrobeRefSetup;
                            }
                        }
                    }
                }

                SetColumns(sheetInfo, i, arr);

                lines.Add(WriteHeader(arr));
            }
            #endregion

            #region data
            for (int index = 0; index < Rows.Count; index++)
            {
                TSet tSet = Rows[index];
                for (int i = 0; i < tSet.TimingRows.Count; i++)
                {
                    TimingRow row = tSet.TimingRows[i];
                    string[] arr = [.. Enumerable.Repeat("", maxCount)];
                    if (!string.IsNullOrEmpty(row.PinGrpName))
                    {
                        arr[0] = row.ColumnA ?? "";
                        arr[timeSetIndex] = tSet.Name;
                        arr[periodIndex] = tSet.CyclePeriod;
                        arr[nameIndex] = row.PinGrpName;
                        arr[clockPeriodIndex] = row.PinGrpClockPeriod;
                        arr[setupIndex] = row.PinGrpSetup;
                        arr[srcIndex] = row.DataSrc;
                        arr[fmtIndex] = row.DataFmt;
                        arr[onIndex] = row.DriveOn;
                        arr[dataIndex] = row.DriveData;
                        arr[returnIndex] = row.DriveReturn;
                        arr[offIndex] = row.DriveOff;
                        arr[compareModeIndex] = row.CompareMode;
                        arr[compareOpenIndex] = row.CompareOpen;
                        arr[compareCloseIndex] = row.CompareClose;
                        arr[compareRefOffsetIndex] = row.CompareRefOffset;
                        arr[edgeModeIndex] = row.EdgeMode;
                        arr[commentIndex] = row.Comment;
                    }
                    else
                    {
                        arr = ["\t"];
                    }
                    lines.Add(string.Join("\t", arr));
                }
            }
            #endregion
            return lines;
        }
    }
}
