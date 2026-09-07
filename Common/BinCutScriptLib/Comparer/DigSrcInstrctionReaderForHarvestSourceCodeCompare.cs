using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Printer;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace BinCutScriptLib.Comparer
{
    public class DigSrcInstrctionReaderForHarvestSourceCodeCompare : MySheetReader<DigSrcInstructionSheet>
    {
        private const string ConHeaderPatternName = "PatternName";
        //private const string ConHeaderBitNumber = "Dig SRC bit number";
        //private const string ConHeaderBitDescription = "Bit description";
        //private const string ConHeaderFunction = "Function";
        //private const string ConHeaderDirective = "Directive";

        private int _indexPatternName = -1;
        //private int _indexBitNumber = -1;
        //private int _indexBitDescription = -1;
        //private int _indexFunction = -1;
        //private int _indexDirective = -1;

        private readonly Dictionary<string, int> _patternKeyWords = [];

        public override DigSrcInstructionSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            DigSrcInstructionSheet harvPmodeTableModeSheet = ReadSheet(sheetName);

            return harvPmodeTableModeSheet;
        }

        private DigSrcInstructionSheet ReadSheet(string sheetName)
        {
            var digSrcInstructionSheet = new DigSrcInstructionSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new DigSrcInstructionSheetRow(sheetName);
                if (!ExcelWorksheet.GetCellValue(i, _indexPatternName).Trim().EqualsIgnoreCase("END"))
                {
                    row.PatternName = ExcelWorksheet.GetCellValue(i, _indexPatternName).ToUpper().Trim();
                    digSrcInstructionSheet.Rows.Add(row);
                }
                else
                {
                    break;
                }
            }
            foreach (KeyValuePair<string, int> patternKeyWord in _patternKeyWords)
            {
                var column = new DigSrcInstructionSheetColumn(sheetName)
                {
                    KeyWord = patternKeyWord.Key,
                    BitList = [],
                    ColumnNum = patternKeyWord.Value
                };
                for (int j = StartRow + 1; j <= EndRow; j++)
                {
                    if (!string.IsNullOrEmpty(ExcelWorksheet.GetCellValue(j, patternKeyWord.Value).ToUpper().Trim()))
                    {
                        column.BitList.Add(ExcelWorksheet.GetCellValue(j, patternKeyWord.Value).ToUpper().Trim());
                    }
                }
                digSrcInstructionSheet.Columns.Add(column);
            }

            return digSrcInstructionSheet;
        }

        private bool GetFirstHeaderPosition()
        {
            //int rowNum = EndRowNumber > 10 ? 10 : EndRowNumber;
            //int colNum = EndColNumber > 10 ? 10 : EndColNumber;
            for (int i = 1; i <= EndRow; i++)
            //for (int j = 1; j <= colNum; j++)
            {
                if (ExcelWorksheet.GetCellValue(i, 1).Trim().EqualsIgnoreCase(ConHeaderPatternName))
                {
                    StartRow = i;
                    return true;
                }
            }
            return false;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();

                if (header.EqualsIgnoreCase("End"))
                {
                    break;
                }

                if (header.EqualsIgnoreCase(ConHeaderPatternName))
                {
                    _indexPatternName = i;
                    continue;
                }
                //if (header.Equals(ConHeaderBitNumber, StringComparison.OrdinalIgnoreCase))
                //{
                //    _indexBitNumber = i;
                //    continue;
                //}
                //if (header.Equals(ConHeaderBitDescription, StringComparison.OrdinalIgnoreCase))
                //{
                //    _indexBitDescription = i;
                //    continue;
                //}
                //if (header.Equals(ConHeaderFunction, StringComparison.OrdinalIgnoreCase))
                //{
                //    _indexFunction = i;
                //    continue;
                //}
                //if (header.Equals(ConHeaderDirective, StringComparison.OrdinalIgnoreCase))
                //{
                //    _indexDirective = i;
                //    continue;
                //}
                _patternKeyWords.TryAdd(header, i);
            }
        }
    }

    public class DigSrcInstructionSheet : MySheet
    {
        public List<DigSrcInstructionSheetColumn> Columns { set; get; }
        public List<DigSrcInstructionSheetRow> Rows { set; get; }

        #region Constructor
        public DigSrcInstructionSheet(string sheetName)
        {
            SheetName = sheetName;
            Columns = [];
            Rows = [];
        }
        #endregion

        public List<bool> GetBits(string pattern, string key, Dictionary<string, string> harvFlags, StreamWriter streamWriter, HarvestSourceCodetLine harvestSourceCodetLine)
        {
            HashSet<int> printedRows = [];
            var bits = new List<bool>();
            if (!Rows.Exists(x => x.PatternName.EqualsIgnoreCase(pattern)))
            {
                return bits;
            }

            foreach (DigSrcInstructionSheetColumn column in Columns)
            {
                string patText = "^" + column.KeyWord.Trim().Replace("*", ".*");
                if (!Regex.IsMatch(key, patText))
                {
                    continue;
                }

                for (int i = 0; i < column.BitList.Count; i++)
                {
                    string bit = column.BitList[i];
                    if (bit == "1" || bit == "0")
                    {
                        bits.Add(bit == "1");
                        continue;
                    }
                    char[] bitChar = bit.ToCharArray();
                    List<string> bitTmp = GetBitTmps(bitChar);
                    for (int k = 0; k < bitTmp.Count; k++)
                    {
                        if (harvFlags.ContainsKey(bitTmp[k]))
                        {
                            bitTmp[k] = harvFlags.First(x => x.Key == bitTmp[k]).Value;
                        }
                    }
                    bit = string.Concat(bitTmp);
                    bit = HandleBit(streamWriter, harvestSourceCodetLine, printedRows, bit);

                    bool finalOperationResult;
                    try
                    {
                        bool matchInput = ContainsInvalidCharacters(bit);

                        if (string.IsNullOrWhiteSpace(bit))
                        {
                            throw new ArgumentException("Input cannot be null or empty.");
                        }
                        else if (matchInput)
                        {
                            if (!printedRows.Contains(harvestSourceCodetLine.LineNo))
                            {
                                BinCutPrint.PrintHarvFlagNotFind(streamWriter, harvestSourceCodetLine);
                                printedRows.Add(harvestSourceCodetLine.LineNo);
                            }
                            finalOperationResult = false;
                        }
                        else
                        {
                            bit = bit.Replace("&&", " AND ")
                                     .Replace("||", " OR ")
                                     .Replace("!T", "F")
                                     .Replace("!F", "T")
                                     .Replace("F", "0>1")
                                     .Replace("T", "1>0");

                            var dt = new DataTable();
                            object finalResult = dt.Compute(bit, null);
                            finalOperationResult = Convert.ToBoolean(finalResult);
                        }
                    }
                    catch (Exception)
                    {
                        if (!printedRows.Contains(harvestSourceCodetLine.LineNo))
                        {
                            BinCutPrint.PrintHarvFlagNotFind(streamWriter, harvestSourceCodetLine);
                            printedRows.Add(harvestSourceCodetLine.LineNo);
                            break;
                        }
                        finalOperationResult = false;
                    }

                    bits.Add(finalOperationResult);
                }
            }
            return bits;
        }

        private static string HandleBit(StreamWriter streamWriter, HarvestSourceCodetLine harvestSourceCodetLine, HashSet<int> printedRows, string bit)
        {
            while (true)
            {
                string[] sptConditions = bit.Split(['(', ')'], StringSplitOptions.RemoveEmptyEntries);
                if (sptConditions.Length == 1)
                {
                    break;
                }

                for (int l = 0; l < sptConditions.Length; l++)
                {
                    string sptCondition = sptConditions[l];

                    if (sptCondition.Split(['|', '&', ' ', '!'], StringSplitOptions.RemoveEmptyEntries).Length >= 2)
                    {
                        bool operationResult;
                        try
                        {
                            bool matchInput = ContainsInvalidCharacters(sptCondition);

                            if (string.IsNullOrWhiteSpace(sptCondition))
                            {
                                throw new ArgumentException("Input cannot be null or empty.");
                            }
                            else if (matchInput)
                            {
                                if (!printedRows.Contains(harvestSourceCodetLine.LineNo))
                                {
                                    if (!printedRows.Contains(harvestSourceCodetLine.LineNo))
                                    {
                                        BinCutPrint.PrintHarvFlagNotFind(streamWriter, harvestSourceCodetLine);
                                        printedRows.Add(harvestSourceCodetLine.LineNo);
                                        break;
                                    }
                                    printedRows.Add(harvestSourceCodetLine.LineNo);
                                }
                                operationResult = false;
                            }
                            else
                            {
                                sptCondition = sptCondition.Replace("&&", " AND ")
                                                           .Replace("||", " OR ")
                                                           .Replace("!T", "F")
                                                           .Replace("!F", "T")
                                                           .Replace("F", "0>1")
                                                           .Replace("T", "1>0");

                                var dt = new DataTable();
                                object result = dt.Compute(sptCondition, null);
                                operationResult = Convert.ToBoolean(result);
                            }
                        }
                        catch (Exception)
                        {
                            if (!printedRows.Contains(harvestSourceCodetLine.LineNo))
                            {
                                if (!printedRows.Contains(harvestSourceCodetLine.LineNo))
                                {
                                    BinCutPrint.PrintHarvFlagNotFind(streamWriter, harvestSourceCodetLine);
                                    printedRows.Add(harvestSourceCodetLine.LineNo);
                                }
                                printedRows.Add(harvestSourceCodetLine.LineNo);
                            }
                            operationResult = false;
                        }

                        sptCondition = operationResult ? "T" : "F";
                        bit = bit.Replace("(" + sptConditions[l] + ")", sptCondition);
                    }
                }
            }

            return bit;
        }

        private static List<string> GetBitTmps(char[] bitChar)
        {
            var bitTmp = new List<string>();
            string flagTmp = "";
            for (int j = 0; j < bitChar.Length; j++)
            {
                if (IsSymbol(bitChar[j]) && string.IsNullOrEmpty(flagTmp))
                {
                    bitTmp.Add(bitChar[j].ToString());
                    continue;
                }
                if (IsSymbol(bitChar[j]) && !string.IsNullOrEmpty(flagTmp))
                {
                    bitTmp.Add(flagTmp);
                    bitTmp.Add(bitChar[j].ToString());
                    flagTmp = "";
                    continue;
                }
                flagTmp += bitChar[j];
                if (j == bitChar.Length - 1)
                {
                    bitTmp.Add(flagTmp);
                }
            }

            return bitTmp;
        }

        public static bool ContainsInvalidCharacters(string input)
        {
            string validPattern = @"(&&|\|\||!T|!F|T|F)";

            string cleanedInput = Regex.Replace(input, validPattern, "");

            return !string.IsNullOrWhiteSpace(cleanedInput);
        }

        private static bool IsSymbol(char c)
        {
            return c.Equals('!') || c.Equals('(') || c.Equals(')') || c.Equals('&') || c.Equals('|') || c.Equals(' ');
        }
    }

    public class DigSrcInstructionSheetColumn : MyColumn
    {
        public string KeyWord { set; get; } = string.Empty;
        public List<string> BitList { set; get; } = [];

        #region Constructor
        public DigSrcInstructionSheetColumn(string sheetName)
        {
            SheetName = sheetName;
        }
        #endregion
    }

    public class DigSrcInstructionSheetRow : MyRow
    {
        public string PatternName { set; get; } = string.Empty;

        #region Constructor
        public DigSrcInstructionSheetRow(string sheetName)
        {
            SheetName = sheetName;
        }
        #endregion
    }
}
