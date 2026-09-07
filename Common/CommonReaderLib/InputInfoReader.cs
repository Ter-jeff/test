using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using OfficeOpenXml;

namespace CommonReaderLib
{
    public class InputInfoReader(string inputInfo) : MySheetReader<InputInfoSheet>
    {
        private const string HeaderItem = "Item";
        private const string HeaderValue = "Value";
        private const string HeaderDescription = "Description";

        private readonly string _inputInfo = inputInfo;
        private string _sheetName = string.Empty;

        private int _indexItem = -1;
        private int _indexValue = -1;
        private int _indexDescription = -1;

        public override InputInfoSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;

            var inputInfoSheet = new InputInfoSheet(_sheetName);

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition(HeaderItem))
            {
                return null;
            }

            if (!GetHeaderIndex())
            {
                return null;
            }

            inputInfoSheet = ReadSheet(inputInfoSheet);

            inputInfoSheet.ModifyPath(_inputInfo);

            return inputInfoSheet;
        }

        private InputInfoSheet ReadSheet(InputInfoSheet inputInfoSheet)
        {
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                InputInfoRow row = new InputInfoRow(_sheetName)
                {
                    RowNum = i
                };
                if (_indexItem != -1)
                {
                    row.Item = ExcelWorksheet.GetCellValue(i, _indexItem).Trim();
                }

                if (_indexValue != -1)
                {
                    row.Value = ExcelWorksheet.GetCellValue(i, _indexValue).Trim();
                }

                if (_indexDescription != -1)
                {
                    row.Description = ExcelWorksheet.GetCellValue(i, _indexDescription).Trim();
                }

                inputInfoSheet.Rows.Add(row);
            }
            inputInfoSheet.IndexItem = _indexItem;
            inputInfoSheet.IndexValue = _indexValue;
            inputInfoSheet.IndexDescription = _indexDescription;
            return inputInfoSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(HeaderItem))
                {
                    _indexItem = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(HeaderValue))
                {
                    _indexValue = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(HeaderDescription))
                {
                    _indexDescription = i;
                }
            }

            return true;
        }

        private void Reset()
        {
            StartCol = -1;
            StartRow = -1;
            EndCol = -1;
            EndRow = -1;
            _indexItem = -1;
            _indexValue = -1;
            _indexDescription = -1;
        }
    }

    public class InputInfoSheet : MySheet
    {
        private static readonly HashSet<string> _checkFileItems = new(StringExtensions.IgnoreCase)
        {
            "testplanfile", "voltagetablefile", "scghfile", "bincutmodesequence",
            "postbincut", "bincut", "powerbinning", "efusetestplan", "fusechecktable",
            "eqn", "dramtype", "ttr", "binoutreport", "charplan", "patlistcsvfile",
            "compilepat", "patterninfofile"
        };

        private static readonly HashSet<string> _checkFolderItems = new(StringExtensions.IgnoreCase)
        {
            "testplanfile", "repo", "cslibrarypath", "baslibraryfolder", "patternfolder",
            "timesetfolder", "custom"
        };

        private static readonly HashSet<string> _modifyFileItems = new(StringExtensions.IgnoreCase)
        {
            "testplanfile", "voltagetablefiles", "scghfile", "bincutmodesequence",
            "postbincut", "bincuts", "powerbinnings", "efusetestplan", "fusechecktable",
            "eqn", "dramtype", "ttrs", "binoutreport", "charplan", "patlistcsvfile",
            "compilepat", "patterninfofile"
        };

        private static readonly HashSet<string> _modifyFolderItems = new(StringExtensions.IgnoreCase)
        {
            "repo", "cslibrarypath", "baslibraryfolder", "patternfolder",
            "timesetfolder", "custom", "outputfolder"
        };

        public List<InputInfoRow> Rows { get; set; }

        public int IndexItem = -1;
        public int IndexValue = -1;
        public int IndexDescription = -1;

        public InputInfoSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }

        public void Check()
        {
            foreach ((InputInfoRow row, string[] paths, bool shouldCheckFile, bool shouldCheckFolder) in GetPathRows(_checkFileItems, _checkFolderItems))
            {
                foreach (string path in paths)
                {
                    bool exists = (shouldCheckFolder && Directory.Exists(path)) || (shouldCheckFile && File.Exists(path));

                    if (!exists)
                    {
                        AddError(PreActionErrorType.E_MissingFile_01, SheetName, row.RowNum, 1, [row.Item, path]);
                    }
                }
            }
        }

        public void ModifyPath(string inputInfo)
        {
            string currentDir = Path.GetDirectoryName(inputInfo) ?? "";

            foreach ((InputInfoRow row, string[] paths, _, _) in GetPathRows(_modifyFileItems, _modifyFolderItems))
            {
                var result = new List<string>();

                foreach (string rawPath in paths)
                {
                    string normalizedPath = rawPath.Replace('\\', '/');
                    if (IsAbsolutePath(normalizedPath))
                    {
                        result.Add(Path.GetFullPath(normalizedPath));
                        continue;
                    }

                    string combinedPath = Path.Combine(currentDir, normalizedPath);
                    string fullPath = Path.GetFullPath(combinedPath);

                    result.Add(fullPath);
                }

                row.Value = string.Join(",", result);
            }
        }

        private IEnumerable<(InputInfoRow Row, string[] Paths, bool ShouldCheckFile, bool ShouldCheckFolder)> GetPathRows(HashSet<string> fileItems, HashSet<string> folderItems)
        {
            foreach (InputInfoRow row in Rows)
            {
                if (string.IsNullOrWhiteSpace(row.Item) || string.IsNullOrWhiteSpace(row.Value))
                {
                    continue;
                }

                string item = row.Item.Trim();
                bool shouldCheckFile = fileItems.Contains(item);
                bool shouldCheckFolder = folderItems.Contains(item);

                if (!shouldCheckFile && !shouldCheckFolder)
                {
                    continue;
                }

                string[] paths = [.. row.Value.Split([','], StringSplitOptions.RemoveEmptyEntries).Select(p => p.Trim())];
                yield return (row, paths, shouldCheckFile, shouldCheckFolder);
            }
        }

        private static bool IsAbsolutePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            return Path.IsPathRooted(path) && !path.StartsWith('.') && !path.StartsWith("..");
        }
    }

    public class InputInfoRow(string sourceSheetName) : MyRow
    {
        public string SourceSheetName { set; get; } = sourceSheetName;
        public string Item { set; get; } = string.Empty;
        public string Value { set; get; } = string.Empty;
        public string Description { set; get; } = string.Empty;
    }
}
