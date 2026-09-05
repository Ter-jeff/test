using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using TagDiff.Core.Common;

namespace TagDiffCore.Utility
{
    public static partial class SheetProvider
    {
        [GeneratedRegex(@"version=(?<version>([\d]|[\.])*):", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex _regexSheetVersion();

        public static (int startRow, int endRow, int startCol, int endCol) GetDimensions(List<string[]> data, string key, bool includeBackup)
        {
            if (data == null || data.Count == 0)
            {
                return (0, 0, 0, 0);
            }

            if (string.IsNullOrEmpty(key))
            {
                return (0, data.Count - 1, 0, data.Max(x => x.Length) - 1);
            }

            int startRow = -1;
            int startCol = -1;
            for (int r = 0; r < data.Count; r++)
            {
                string[] row = data[r];
                if (row == null)
                {
                    continue;
                }

                for (int c = 0; c < row.Length; c++)
                {
                    if (row[c].EqualsIgnoreCase(key))
                    {
                        startRow = r;
                        startCol = c;
                        break;
                    }
                }
                if (startRow != -1)
                {
                    break;
                }
            }

            if (startRow == -1)
            {
                throw new Exception($"Header {key} not found!");
            }

            int endCol = data[startRow] != null ? data[startRow].Length - 1 : startCol;
            string[] headerRow = data[startRow];
            for (int c = startCol + 1; headerRow != null && c < headerRow.Length; c++)
            {
                if (string.IsNullOrWhiteSpace(headerRow[c]))
                {
                    endCol = c - 1;
                    break;
                }
            }

            int endRow = data.Count - 1;
            if (!includeBackup)
            {
                for (int r = startRow; r < data.Count; r++)
                {
                    string[] row = data[r];
                    if (row == null || row.All(string.IsNullOrWhiteSpace))
                    {
                        endRow = r - 1;
                        break;
                    }
                }
            }

            return (startRow, endRow, startCol, endCol);
        }

        public static Location? GetFirstLocation(Dictionary<string, List<Location>> dic, string itemKey)
        {
            if (dic.TryGetValue(itemKey, out List<Location>? locations) && locations != null)
            {
                for (int i = 0; i < locations.Count; i++)
                {
                    Location l = locations[i];
                    if (l != null && !l.IsUsed)
                    {
                        return l;
                    }
                }
            }

            return null;
        }

        public static bool IsEmptyRow(List<string[]> sheet, int row)
        {
            ArgumentNullException.ThrowIfNull(sheet);

            if (row < 0 || row >= sheet.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(row));
            }

            string[] rowList = sheet[row];
            if (rowList == null)
            {
                return true;
            }

            return rowList.All(string.IsNullOrEmpty);
        }

        public static string GetSheetVersion(List<string[]> sheet, string sheetName = "Unknown")
        {
            ArgumentNullException.ThrowIfNull(sheet);

            if (TryGetSheetVersion(sheet, out string version))
            {
                return version;
            }

            throw new InvalidOperationException($"Can not find the sheet version of '{sheetName}'!");
        }

        private static bool TryGetSheetVersion(List<string[]> sheet, out string version)
        {
            version = "";
            ArgumentNullException.ThrowIfNull(sheet);

            foreach (string[] row in sheet)
            {
                if (row == null)
                {
                    continue;
                }

                for (int i = 0; i < row.Length; i++)
                {
                    string cell = row[i] ?? string.Empty;
                    Match match = _regexSheetVersion().Match(cell);
                    if (match.Success)
                    {
                        version = match.Groups["version"].Value;
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
