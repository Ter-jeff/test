using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;
using IgxlLib.Utility.Tester;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public partial class ReadChanMapSheet : IgxlSheetReader<ChannelMapSheet>
    {
        public const int DeviceUnderTestPinNameIndex = 2;
        public const int DeviceUnderTestPackagePinIndex = 3;
        public const int DeviceUnderTestPinTypeIndex = 4;

        private const string RegWalkRoundChannel = @"(?<mainCh>\d+\.\d+)e\+(?<SubChan>\d+)";
        private const int StartRowIndex = 6;

        [GeneratedRegex(@"^site\s*\d+", RegexOptions.Compiled)]
        private static partial Regex SiteHeaderRegex();
        [GeneratedRegex(@"^comment\s*", RegexOptions.Compiled)]
        private static partial Regex CommentHeaderRegex();
        [GeneratedRegex(RegWalkRoundChannel, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex WalkRoundChannelRegex();
        [GeneratedRegex(@"^S\:", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex AliasSitePrefixRegex();

        private static readonly Regex _siteHeaderRegex = SiteHeaderRegex();
        private static readonly Regex _commentHeaderRegex = CommentHeaderRegex();
        private static readonly Regex _walkRoundChannelRegex = WalkRoundChannelRegex();
        private static readonly Regex _aliasSitePrefixRegex = AliasSitePrefixRegex();

        public override EnumSheetType SupportedType => EnumSheetType.DTChanMap;

        public override ChannelMapSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            return GetSheet(excelWorksheet, null);
        }

        public override ChannelMapSheet ReadSheet(Stream stream, string sheetName)
        {
            var sheet = new ChannelMapSheet(sheetName);
            var lines = new List<string>();
            using (var sr = new StreamReader(stream))
            {
                string? line;
                while ((line = sr.ReadLine()) != null)
                {
                    lines.Add(line);
                }
            }

            var siteIndex = new List<int>();
            for (int index = 0; index < lines.Count; index++)
            {
                if (lines[index].ContainsIgnoreCase("Site"))
                {
                    siteIndex = GetSiteColumnIndex(lines[index]);
                    break;
                }
            }

            sheet.SiteNum = siteIndex.Count;
            sheet.IsPogo = GetViewMode(lines);

            for (int index = 0; index < lines.Count; index++)
            {
                string line = lines[index];
                string[] arr = line.Split('\t');
                if (index + 1 > StartRowIndex)
                {
                    var channelMapRow = new ChannelMapRow
                    {
                        SheetName = sheetName,
                        RowNum = index + 1,
                        DeviceUnderTestPinName = arr[DeviceUnderTestPinNameIndex - 1],
                        DeviceUnderTestPackagePin = arr[DeviceUnderTestPackagePinIndex - 1],
                        Type = GetPinType(arr[DeviceUnderTestPinTypeIndex - 1])
                    };
                    foreach (int columnIndex in siteIndex)
                    {
                        string siteValue = arr[columnIndex];
                        if (_walkRoundChannelRegex.IsMatch(siteValue))
                        {
                            string mainChan = _walkRoundChannelRegex.Match(siteValue).Groups["mainCh"].Value;
                            string subChan = _walkRoundChannelRegex.Match(siteValue).Groups["SubChan"].Value;
                            siteValue = $"{mainChan.Replace(".", "")}.e{int.Parse(subChan) - 1}";
                        }

                        channelMapRow.Sites.Add(siteValue);
                    }

                    if (string.IsNullOrEmpty(channelMapRow.DeviceUnderTestPinName) &&
                        string.IsNullOrEmpty(channelMapRow.Type))
                    {
                        break;
                    }

                    sheet.AddRow(channelMapRow);
                }
            }

            return sheet;
        }

        public static ChannelMapSheet GetSheet(string fileName, Dictionary<string, TesterConfig>? testerConfigDic)
        {
            return GetSheet(IgxlSheetReaderHelpers.ConvertTxtToExcelSheet(fileName), testerConfigDic);
        }

        public static ChannelMapSheet GetSheet(ExcelWorksheet excelWorksheet, Dictionary<string, TesterConfig>? testerConfigDic)
        {
            int endRow = excelWorksheet.Dimension.End.Row;
            int endColumn = excelWorksheet.Dimension.End.Column;
            var channelMapSheet = new ChannelMapSheet(excelWorksheet);

            #region parse the header columns
            int pinNameIdx = 0;
            int packagePinIdx = 0;
            int pinTypeIdx = 0;
            int commentIdx = 0;
            var siteColumnIndex = new List<int>();
            int headerRowIdx = 0;
            bool isAllHeadersFound = false;

            for (int i = 0; i <= endRow; i++)
            {
                for (int j = 0; j <= endColumn; j++)
                {
                    string cellValue = EpplusExtensions.GetCellValue(excelWorksheet, i, j).ToLower();
                    if (_siteHeaderRegex.IsMatch(cellValue))
                    {
                        siteColumnIndex.Add(j);
                    }
                    else if (cellValue.StartsWith("package pin"))
                    {
                        packagePinIdx = j;
                    }
                    else if (cellValue.StartsWith("pin name"))
                    {
                        pinNameIdx = j;
                    }
                    else if (cellValue.StartsWith("type"))
                    {
                        pinTypeIdx = j;
                    }
                    else if (_commentHeaderRegex.IsMatch(cellValue))
                    {
                        commentIdx = j;
                    }
                    if (pinNameIdx > 0 && pinTypeIdx > 0 && siteColumnIndex.Count > 0 && commentIdx > 0)
                    {
                        isAllHeadersFound = true;
                        break;
                    }
                }
                if (isAllHeadersFound)
                {
                    headerRowIdx = i;
                    break;
                }
            }
            #endregion

            channelMapSheet.SiteNum = siteColumnIndex.Count;
            channelMapSheet.IsPogo = GetViewMode(excelWorksheet);

            int startRow = headerRowIdx + 1;
            for (int i = startRow; i <= endRow; i++)
            {
                var channelMapRow = new ChannelMapRow
                {
                    DeviceUnderTestPinName = EpplusExtensions.GetCellValue(excelWorksheet, i, pinNameIdx),
                    DeviceUnderTestPackagePin = EpplusExtensions.GetCellValue(excelWorksheet, i, packagePinIdx),
                    Type = GetPinType(EpplusExtensions.GetCellValue(excelWorksheet, i, pinTypeIdx))
                };

                foreach (int columnIndex in siteColumnIndex)
                {
                    string siteValue = EpplusExtensions.GetCellValue(excelWorksheet, i, columnIndex);
                    if (_walkRoundChannelRegex.IsMatch(siteValue))
                    {
                        string mainChan = _walkRoundChannelRegex.Match(siteValue).Groups["mainCh"].Value;
                        string subChan = _walkRoundChannelRegex.Match(siteValue).Groups["SubChan"].Value;
                        _ = double.TryParse(mainChan, out double value);
                        siteValue = $"{value.ToString(CultureInfo.InvariantCulture).Replace(".", "")}.e{int.Parse(subChan) - 1}";
                    }
                    channelMapRow.Sites.Add(siteValue);
                }

                if (string.IsNullOrEmpty(channelMapRow.DeviceUnderTestPinName) && string.IsNullOrEmpty(channelMapRow.Type))
                {
                    break;
                }

                channelMapSheet.AddRow(channelMapRow);
            }

            foreach (ChannelMapRow row in channelMapSheet.Rows)
            {
                // S:VDD_ANA_S_UVI80 => equals to VDD_ANA_S_UVI80
                var sties = new List<string>();
                int cnt = 0;
                foreach (string site in row.Sites)
                {
                    sties.Add(_aliasSitePrefixRegex.IsMatch(site) ? GetAliasSiteInfo(channelMapSheet, row.Sites, cnt) : site);
                    cnt++;
                }
                if (testerConfigDic != null)
                {
                    row.InstrumentType = TesterConfigManager.GetToolTypeByChannelAssignments(testerConfigDic, sties, excelWorksheet.Name);
                }
            }
            return channelMapSheet;
        }

        private static bool GetViewMode(ExcelWorksheet excelWorksheet)
        {
            for (int i = 1; i <= 7; i++)
            {
                for (int j = 1; j <= 10; j++)
                {
                    string text = EpplusExtensions.GetCellValue(excelWorksheet, i, j);
                    if (text.EqualsIgnoreCase("View Mode:"))
                    {
                        string mode = EpplusExtensions.GetCellValue(excelWorksheet, i, j + 1);
                        if (mode.EqualsIgnoreCase("Pogo"))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        private static string GetPinType(string type)
        {
            if (type.EqualsIgnoreCase("IO"))
            {
                return "I/O";
            }

            return type;
        }

        private static string GetAliasSiteInfo(ChannelMapSheet channelMapSheet, List<string> rowSiteData, int siteIndex)
        {
            string site = rowSiteData[siteIndex].Split(':')[1];
            for (int i = 0; i < channelMapSheet.Rows.Count; ++i)
            {
                ChannelMapRow mapRow = channelMapSheet.Rows[i];
                if (mapRow.DeviceUnderTestPinName == site)
                {
                    return mapRow.Sites[siteIndex];
                }
            }
            return "";
        }

        private static List<int> GetSiteColumnIndex(string line)
        {
            var siteColumnIndex = new List<int>();
            string[] arr = line.Split('\t');
            for (int col = 0; col < arr.Length; col++)
            {
                if (_siteHeaderRegex.IsMatch(arr[col].ToLower()))
                {
                    siteColumnIndex.Add(col);
                }
            }

            return siteColumnIndex;
        }

        private static bool GetViewMode(List<string> lines)
        {
            for (int i = 0; i < 7; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    string text = lines[i];
                    string[] arr = text.Split('\t');
                    if (j + 1 < arr.Length)
                    {
                        if (arr[j].EqualsIgnoreCase("View Mode:"))
                        {
                            string mode = arr[j + 1];
                            if (mode.EqualsIgnoreCase("Pogo"))
                            {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }
    }
}
