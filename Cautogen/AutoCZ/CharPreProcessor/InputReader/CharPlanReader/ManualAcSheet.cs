using System;
using System.Collections.Generic;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader
{
    public class ManualAcSheet
    {
        public Dictionary<string, Dictionary<string, string>> AcCateSymbolValue = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> CategorySet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> SymbolSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public ManualAcSheet(ExcelWorksheet sh)
        {
            if (sh != null)
            {
                AcCateSymbolValue = _Read(sh);
            }
        }

        public void Add(string category, string symbol, string value)
        {
            if (string.IsNullOrEmpty(category) || string.IsNullOrEmpty(symbol) || string.IsNullOrEmpty(value))
            {
                return;
            }

            if (!AcCateSymbolValue.ContainsKey(category))
            {
                AcCateSymbolValue.Add(category, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
            }

            if (!AcCateSymbolValue[category].ContainsKey(symbol))
            {
                AcCateSymbolValue[category].Add(symbol, value);
            }
            else
            {
                AcCateSymbolValue[category][symbol] = value;
            }

            CategorySet.Add(category);
            SymbolSet.Add(symbol);
        }

        public void Write(ExcelWorksheet sh)
        {
            Tuple<Dictionary<string, int>, Dictionary<string, int>> headerTuple = _WriteCateSymbol(sh);
            Dictionary<string, int> rowDict = headerTuple.Item1;
            Dictionary<string, int> colDict = headerTuple.Item2;

            foreach (KeyValuePair<string, Dictionary<string, string>> cateDict in AcCateSymbolValue)
            {
                foreach (KeyValuePair<string, string> symbolValue in cateDict.Value)
                {
                    sh.Cells[rowDict[symbolValue.Key], colDict[cateDict.Key]].Value = symbolValue.Value;
                }
            }
        }

        private Tuple<Dictionary<string, int>, Dictionary<string, int>> _WriteCateSymbol(ExcelWorksheet sh)
        {
            sh.Cells[1, 1].Value = "Symbol";
            int startCateCol = 2;
            int startSymbolRow = 2;
            var rowDict = new Dictionary<string, int>();
            var colDict = new Dictionary<string, int>();
            foreach (string cate in CategorySet)
            {
                sh.Cells[1, startCateCol].Value = cate;
                colDict.Add(cate, startCateCol);
                startCateCol++;
            }

            foreach (string symbol in SymbolSet)
            {
                sh.Cells[startSymbolRow, 1].Value = symbol;
                rowDict.Add(symbol, startSymbolRow);
                startSymbolRow++;
            }
            return Tuple.Create(rowDict, colDict);
        }

        private Dictionary<string, Dictionary<string, string>> _Read(ExcelWorksheet sh)
        {
            if (sh == null)
            {
                return null;
            }

            var manualAc = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
            const int categoryRow = 1;
            const int symbolCol = 1;

            for (int colindex = 2; colindex <= sh.Dimension.Columns; colindex++)
            {
                string category = "";
                for (int rowindex = categoryRow; rowindex <= sh.Dimension.Rows; rowindex++)
                {
                    if (string.IsNullOrEmpty(sh.Cells[categoryRow, colindex].Text))
                    {
                        continue;
                    }

                    if (rowindex == categoryRow)
                    {
                        category = sh.Cells[categoryRow, colindex].Text.Trim();
                        if (!string.IsNullOrEmpty(category))
                        {
                            manualAc.Add(category, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
                            CategorySet.Add(category);
                        }
                    }
                    else
                    {
                        string symbol = sh.Cells[rowindex, symbolCol].Text.Trim();
                        string value = sh.Cells[rowindex, colindex].Text.Trim();
                        value = string.IsNullOrEmpty(value) ? "" : value;

                        try
                        {
                            if (manualAc[category].ContainsKey(symbol))
                            {
                                manualAc[category][symbol] = value;
                            }
                            else
                            {
                                manualAc[category].Add(symbol, value);
                            }

                            if (!SymbolSet.Contains(symbol))
                            {
                                SymbolSet.Add(symbol);
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
            return manualAc;
        }
    }
}
