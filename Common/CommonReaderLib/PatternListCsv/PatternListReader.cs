using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Static;

using OfficeOpenXml;

namespace CommonReaderLib.PatternListCsv
{
    public class PatternListReader : MySheetReader<PatternListSheet>
    {
        private const string HeaderNumber = "#";
        private const string HeaderPattern = "Pattern";
        private const string HeaderUseNotUse = "USE/No Use";
        private const string HeaderTimeSetLatest = "Timeset Latest";
        private const string HeaderFileVersions = "File Versions";

        private int _indexFileVersions = -1;
        private int _indexNumber = -1;
        private int _indexPattern = -1;
        private int _indexTimeSetLatest = -1;
        private int _indexUseNotUse = -1;

        private PatternListSheet _sheet;

        public override PatternListSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            string sheetName = excelWorksheet.Name;

            _sheet = new PatternListSheet(sheetName);

            ExcelWorksheet = excelWorksheet;

            if (!GetDimensions())
            {
                _sheet.AddDimensionError();
                return _sheet;
            }

            if (!GetFirstHeaderPosition())
            {
                _sheet.AddFirstHeaderError(HeaderNumber);
                return _sheet;
            }

            GetHeaderIndex();

            return ReadSheet(sheetName);
        }

        private PatternListSheet ReadSheet(string sheetName)
        {
            var sheet = new PatternListSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new PatternListRow(sheetName);
                row.RowNum = i;
                if (_indexNumber != -1)
                {
                    row.Number = ExcelWorksheet.GetMergeCellValue(i, _indexNumber).Trim();
                }

                if (_indexPattern != -1)
                {
                    row.Pattern = ExcelWorksheet.GetMergeCellValue(i, _indexPattern).Trim();
                }

                if (_indexUseNotUse != -1)
                {
                    row.UseNotUse = ExcelWorksheet.GetMergeCellValue(i, _indexUseNotUse).Trim();
                }

                if (_indexTimeSetLatest != -1)
                {
                    row.TimeSetLatest = ExcelWorksheet.GetMergeCellValue(i, _indexTimeSetLatest).Trim();
                }

                if (_indexFileVersions != -1)
                {
                    row.FileVersions = ExcelWorksheet.GetMergeCellValue(i, _indexFileVersions).Trim();
                }

                if (!string.IsNullOrEmpty(row.Pattern))
                {
                    sheet.Rows.Add(row);
                }
            }

            sheet.IndexNumber = _indexNumber;
            sheet.IndexPattern = _indexPattern;
            sheet.IndexUseNotUse = _indexUseNotUse;
            sheet.IndexTimeSetLatest = _indexTimeSetLatest;
            sheet.IndexFileVersions = _indexFileVersions;

            return sheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(HeaderNumber))
                {
                    _indexNumber = i;
                    continue;
                }

                if (header.EqualsIgnoreCase(HeaderPattern))
                {
                    _indexPattern = i;
                    continue;
                }

                if (header.EqualsIgnoreCase(HeaderUseNotUse))
                {
                    _indexUseNotUse = i;
                    continue;
                }

                if (header.EqualsIgnoreCase(HeaderTimeSetLatest))
                {
                    _indexTimeSetLatest = i;
                    continue;
                }

                if (header.EqualsIgnoreCase(HeaderFileVersions))
                {
                    _indexFileVersions = i;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(HeaderNumber))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }

        public static List<PatternListRow> ReadPatListCsv(string fileName)
        {
            var fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
            var fileReader = new StreamReader(fs);
            try
            {
                var headerOrder = new Dictionary<string, int>();
                var patternListRows = new List<PatternListRow>();
                string? line = fileReader.ReadLine();
                int rowIndex = 0;
                int index = 0;
                if (line != null)
                {
                    foreach (string str in line.Split(','))
                    {
                        headerOrder.Add(str.Replace("\"", ""), index);
                        index++;
                    }

                    while ((line = fileReader.ReadLine()) != null)
                    {
                        rowIndex++;
                        line = line.Replace("\"", "");
                        if (string.IsNullOrEmpty(line.Trim()))
                        {
                            Response.Report(
                                "Blank Row " + rowIndex + " In Pattern List csv " + fileName + " is skipped.",
                                EnumMessageLevel.Warning, 100);
                            continue;
                        }

                        var patternListRow = new PatternListRow(Path.GetFileNameWithoutExtension(fileName));
                        patternListRow.RowNum = rowIndex + 1;
                        string[] arr = new string[index];
                        List<string> lineData = line.Split(',').ToList();
                        for (int i = 0; i < lineData.Count; i++)
                        {
                            if (i < index)
                            {
                                arr[i] = lineData[i];
                            }
                        }

                        if (string.IsNullOrEmpty(line))
                        {
                            Response.Report(
                                "Blank Row " + rowIndex + " In Pattern List csv " + fileName + " is skipped.",
                                EnumMessageLevel.Warning, 100);
                            continue;
                        }

                        if (headerOrder.ContainsKey(HeaderPattern))
                        {
                            patternListRow.Pattern = arr[headerOrder[HeaderPattern]];
                        }

                        if (string.IsNullOrEmpty(patternListRow.Pattern.Trim()))
                        {
                            Response.Report(
                                "Because Pattern is Blank, Row " + rowIndex + " Content " + line +
                                " In Pattern List csv " + fileName + " is skipped.", EnumMessageLevel.Warning, 100);
                            continue;
                        }

                        if (headerOrder.ContainsKey(HeaderNumber))
                        {
                            patternListRow.Number = arr[headerOrder[HeaderNumber]];
                        }

                        if (headerOrder.ContainsKey(HeaderUseNotUse))
                        {
                            patternListRow.UseNotUse = arr[headerOrder[HeaderUseNotUse]];
                        }

                        if (headerOrder.ContainsKey(HeaderTimeSetLatest))
                        {
                            patternListRow.TimeSetLatest =
                                Path.GetFileNameWithoutExtension(arr[headerOrder[HeaderTimeSetLatest]]);
                        }

                        if (headerOrder.ContainsKey(HeaderFileVersions))
                        {
                            patternListRow.FileVersions = arr[headerOrder[HeaderFileVersions]];
                        }

                        patternListRows.Add(patternListRow);
                    }
                }

                return patternListRows;
            }
            catch (Exception e)
            {
                throw new Exception("Reading pattern list failed, may be caused by wrong format of pattern list. " +
                                    e.Message);
            }
            finally
            {
                fileReader.Close();
                fs.Close();
            }
        }
    }
}
