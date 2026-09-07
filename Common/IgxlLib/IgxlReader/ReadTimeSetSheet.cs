using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Regs;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public partial class ReadTimeSetSheet : IgxlSheetReader<TimeSetBasicSheet>
    {
        private const int EndRowIndex = 7;

        [GeneratedRegex("^_", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex LeadingUnderscoreRegex();

        public override EnumSheetType SupportedType => EnumSheetType.DTTimesetBasicSheet;

        private List<string> _headList = [];
        private int _iStartRowIndex;
        private int _iStartColumnIndex;

        public override TimeSetBasicSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var timeSetBasicSheet = new TimeSetBasicSheet(excelWorksheet);
            TSet? tSet = null;
            _headList = [];
            // Get Source Sheet

            _iStartRowIndex = 3;
            _iStartColumnIndex = 2;

            int headRowIndex = _iStartRowIndex + 3;
            int nowRowIndex = _iStartRowIndex;
            int maxRowCount = excelWorksheet.Dimension.End.Row;
            int maxColumCount = excelWorksheet.Dimension.End.Column;

            GetBasicInfo(excelWorksheet, timeSetBasicSheet, ref nowRowIndex, maxColumCount);
            nowRowIndex += 2;
            for (int i = _iStartColumnIndex; i <= maxColumCount; i++)
            {
                string value = GetMergeCellValue(excelWorksheet, headRowIndex, i);
                string value2 = GetCellText(excelWorksheet, headRowIndex + 1, i);

                string head = Combination.CombineByUnderLine(value.Trim(), value2.Trim());

                _headList.Add(head);
            }

            nowRowIndex += 2;

            for (int i = nowRowIndex; i <= maxRowCount; i++)
            {
                for (int j = _iStartColumnIndex; j <= maxColumCount; j++)
                {
                    SetItemContent(excelWorksheet, i, out tSet);
                }

                timeSetBasicSheet.AddRow(tSet!);
            }

            return timeSetBasicSheet;
        }

        public override TimeSetBasicSheet ReadSheet(Stream stream, string sheetName)
        {
            var timeSetBasicSheet = new TimeSetBasicSheet(sheetName);
            const int maxColumnCount = 5;
            var lines = new List<string>();
            using (var reader = new StreamReader(stream))
            {
                string? line;
                while ((line = reader.ReadLine()) != null)
                {
                    lines.Add(line);
                }
            }

            var rows = new List<TimingRow>();
            ReadData(sheetName, timeSetBasicSheet, maxColumnCount, lines, rows);

            string temp = "";
            foreach (TimingRow row in rows)
            {
                if (string.IsNullOrEmpty(row.Name))
                {
                    row.Name = temp;
                }

                if (!string.IsNullOrEmpty(row.Name))
                {
                    temp = row.Name;
                }
            }

            IEnumerable<IGrouping<string, TimingRow>> groups = rows.ChunkBy(x => x.Name);
            foreach (IGrouping<string, TimingRow> group in groups)
            {
                var set = new TSet { Name = group.Key, CyclePeriod = group.ToList()[0].CyclePeriod };
                set.TimingRows.AddRange([.. group]);
                timeSetBasicSheet.AddRow(set);
            }
            return timeSetBasicSheet;
        }

        private static void ReadData(string sheetName, TimeSetBasicSheet timeSetBasicSheet, int maxColumnCount, List<string> lines, List<TimingRow> timingRows)
        {
            for (int i = 0; i < lines.Count; i++)
            {
                string line = lines[i];
                if (i + 1 > EndRowIndex)
                {
                    string[] arr = line.Split('\t');
                    if (arr.Length < 5)
                    {
                        if (!lines[i].Contains('=') && (lines[i].Contains('*') || lines[i].Contains('/')))
                        {
                            continue;
                        }

                        if (lines[i].Split(['\t'], StringSplitOptions.RemoveEmptyEntries).Length == 0)
                        {
                            continue;
                        }

                        if (lines[i].Contains('='))
                        {
                            string[] spt = lines[i].Split('=');
                            string varTok = spt[0].Trim();
                            varTok = LeadingUnderscoreRegex().Replace(varTok, "");
                            string valueTok = spt[1].Trim();
                            string value = valueTok.ConvertNumber();
                            bool isNumOk = double.TryParse(value, out double dValue);
                            if (isNumOk)
                            {
                                timeSetBasicSheet.CommentVariable.TryAdd(varTok, dValue);
                            }
                        }
                        else
                        {
                            string varTok = Reg.WhitespaceRegex.Replace(lines[i], "");
                            const double dValue = -1e9;
                            timeSetBasicSheet.CommentVariable.TryAdd(varTok, dValue);
                        }

                        continue;
                    }

                    TimingRow timingRow = GetTimeRow2P3(arr, sheetName, i + 1);
                    if (string.IsNullOrEmpty(timingRow.PinGrpName))
                    {
                        break;
                    }

                    timingRow.Name = arr[1];
                    timingRow.CyclePeriod = arr[2];
                    timingRows.Add(timingRow);
                }
                else
                {
                    #region Timing Mode & Master Timeset Name

                    string[] arr = line.Split('\t');
                    for (int col = 0; col <= maxColumnCount; col++)
                    {
                        if (col + 1 < arr.Length)
                        {
                            string value = arr[col];
                            if (CellDiff.IsLiked(value, "Timing Mode:"))
                            {
                                value = arr[col + 1];
                                timeSetBasicSheet.TimingMode = value;
                            }

                            if (CellDiff.IsLiked(value, "Master Time Set Name:"))
                            {
                                value = arr[col + 1];
                                timeSetBasicSheet.MasterTimeSet = value;
                            }
                        }
                    }
                    #endregion

                    #region Time Domain & Strobe Ref Setup Name

                    for (int col = 0; col <= maxColumnCount; col++)
                    {
                        if (col + 1 < arr.Length)
                        {
                            string value = arr[col];
                            if (CellDiff.IsLiked(value, "Time Domain:"))
                            {
                                value = arr[col + 1];
                                timeSetBasicSheet.TimeDomain = value;
                            }

                            if (CellDiff.IsLiked(value, "Strobe Ref Setup Name:"))
                            {
                                value = arr[col + 1];
                                timeSetBasicSheet.StrobeRefSetup = value;
                            }
                        }
                    }

                    #endregion
                }
            }
        }

        public static TimingRow GetTimeRow2P3(string[] arr, string sheetName, int rowNum)
        {
            var row = new TimingRow
            {
                SheetName = sheetName,
                RowNum = rowNum,
                ColumnA = arr[0],
                PinGrpName = arr[3],
                PinGrpClockPeriod = arr[4],
                PinGrpSetup = arr[5],
                DataSrc = arr[6],
                DataFmt = arr[7],
                DriveOn = arr[8],
                DriveData = arr[9],
                DriveReturn = arr[10],
                DriveOff = arr[11],
                CompareMode = arr[12],
                CompareOpen = arr[13],
                CompareClose = arr[14],
                CompareRefOffset = arr[15],
                EdgeMode = arr[16],
                Comment = arr.Length > 17 ? arr[17] : ""
            };
            return row;
        }

        private void GetBasicInfo(ExcelWorksheet excelWorksheet, TimeSetBasicSheet timeSetBasicSheet, ref int nowRowIndex, int maxColumCount)
        {
            // Get Basic Info

            #region Timing Mode & Master Timeset Name
            for (int i = _iStartColumnIndex; i <= maxColumCount; i++)
            {
                string value = GetMergeCellValue(excelWorksheet, nowRowIndex, i);
                if (CellDiff.IsLiked(value, "Timing Mode:"))
                {
                    value = GetMergeCellValue(excelWorksheet, nowRowIndex, i + 1);
                    timeSetBasicSheet.TimingMode = value;
                }

                if (CellDiff.IsLiked(value, "Master TimeSet Name:"))
                {
                    value = GetMergeCellValue(excelWorksheet, nowRowIndex, i + 1);
                    timeSetBasicSheet.MasterTimeSet = value;
                }
            }
            #endregion

            #region Time Domain & Strobe Ref Setup Name
            nowRowIndex++;
            for (int i = _iStartColumnIndex; i <= maxColumCount; i++)
            {
                string value = GetMergeCellValue(excelWorksheet, nowRowIndex, i);
                if (CellDiff.IsLiked(value, "Time Domain:"))
                {
                    value = GetMergeCellValue(excelWorksheet, nowRowIndex, i + 1);
                    timeSetBasicSheet.TimeDomain = value;
                }

                if (CellDiff.IsLiked(value, "Strobe Ref Setup Name:"))
                {
                    value = GetMergeCellValue(excelWorksheet, nowRowIndex, i + 1);
                    timeSetBasicSheet.StrobeRefSetup = value;
                }
            }
            #endregion
        }

        private void SetItemContent(ExcelWorksheet excelWorksheet, int pIRowIndex, out TSet tSet)
        {
            tSet = new TSet();
            var timingRow = new TimingRow();

            for (int i = _iStartColumnIndex; i < _iStartColumnIndex + _headList.Count; i++)
            {
                string head = _headList[i - _iStartColumnIndex];
                string content = GetCellText(excelWorksheet, pIRowIndex, i);
                switch (CellDiff.FormatStringForCompare(head))
                {
                    case "TIME SET":
                        tSet.Name = content;
                        break;
                    case "CYCLE_PERIOD":
                        tSet.CyclePeriod = content;
                        break;
                    case "PIN_GROUP_NAME":
                    case "PIN/GROUP_NAME":
                        timingRow.PinGrpName = content;
                        break;
                    case "CLOCK PERIOD":
                        timingRow.PinGrpClockPeriod = content;
                        break;
                    case "SETUP":
                        timingRow.PinGrpSetup = content;
                        break;
                    case "DATA_SRC":
                        timingRow.DataSrc = content;
                        break;
                    case "FMT":
                        timingRow.DataFmt = content;
                        break;
                    case "DRIVE_ON":
                        timingRow.DriveOn = content;
                        break;
                    case "DATA":
                        timingRow.DriveData = content;
                        break;
                    case "RETURN":
                        timingRow.DriveReturn = content;
                        break;
                    case "OFF":
                        timingRow.DriveOff = content;
                        break;
                    case "COMPARE_MODE":
                        timingRow.CompareMode = content;
                        break;
                    case "OPEN":
                        timingRow.CompareOpen = content;
                        break;
                    case "CLOSE":
                        timingRow.CompareClose = content;
                        break;
                    case "REF OFFSET":
                        timingRow.CompareRefOffset = content;
                        break;
                    case "EDGE RESOLUTION_MODE":
                        timingRow.EdgeMode = content;
                        break;
                    case "COMMENT":
                        timingRow.Comment = content;
                        break;
                }
            }

            tSet.TimingRows.Add(timingRow);
        }
    }
}
