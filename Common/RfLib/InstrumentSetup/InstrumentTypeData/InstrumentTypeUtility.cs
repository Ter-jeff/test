using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text.RegularExpressions;

using OfficeOpenXml;

namespace RfLib.InstrumentSetup.InstrumentTypeData
{
    [ExcludeFromCodeCoverage]
    internal partial class InstrumentTypeUtility
    {
        [GeneratedRegex(@"(?<Value>[+-]?(\d+[.])?\d+)\s?\*?(?<Unit>\w+)*")]
        private static partial Regex MyRegex();
        [GeneratedRegex("hz", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();

        #region Operation on excel
        public static string GetMergerdCellValue(ExcelWorksheet excelWorksheet, int rowNumber, int columnNumber)
        {
            string range = excelWorksheet.MergedCells[rowNumber, columnNumber];
            return string.IsNullOrEmpty(range) ?
                GetCellValue(excelWorksheet, rowNumber, columnNumber) :
                GetCellValue(excelWorksheet, new ExcelAddress(range).Start.Row, new ExcelAddress(range).Start.Column);
        }

        public static string GetCellValue(ExcelWorksheet excelWorksheet, int row, int column)
        {
            if (row <= 0 || column <= 0)
            {
                return "";
            }

            if (excelWorksheet.Cells[row, column].Value != null)
            {
                return excelWorksheet.Cells[row, column].Value.ToString()!.Trim();
            }

            return "";
        }
        #endregion

        public static List<InstrumentTypePara> GetInstrumentParameterHeader(ExcelWorksheet excelWorksheet, Dictionary<string, int> typeIndexList, int endColNumber)
        {
            var allInstrumentTypeParas = new List<InstrumentTypePara>();
            InstrumentTypePara instrumentTypePara;
            for (int i = 0; i < typeIndexList.Count; i++)
            {
                instrumentTypePara = new InstrumentTypePara();

                int startCol = typeIndexList.ElementAt(i).Value;
                instrumentTypePara.InstrumentType = typeIndexList.ElementAt(i).Key;
                int endCol;
                if (i + 1 < typeIndexList.Count)
                {
                    endCol = typeIndexList.ElementAt(i + 1).Value;
                }
                else
                {
                    endCol = endColNumber + 1;
                }

                for (int j = startCol; j < endCol; j++)
                {
                    string para = GetCellValue(excelWorksheet, 2, j);
                    if (!string.IsNullOrEmpty(para))
                    {
                        instrumentTypePara.DicPara.Add(para.ToLower(), j);
                    }
                }
                allInstrumentTypeParas.Add(instrumentTypePara);
            }
            return allInstrumentTypeParas;
        }

        public static Dictionary<string, int> GetInstrumentType(ExcelWorksheet excelWorksheet, int endColNumber)
        {
            var typeList = new Dictionary<string, int>();
            for (int i = 1; i < endColNumber; i++)
            {
                string type = GetCellValue(excelWorksheet, 1, i);
                if (!string.IsNullOrEmpty(type))
                {
                    typeList.Add(type, i);
                }
            }
            return typeList;
        }

        public static double ConvertToFreqValue(string freq)
        {
            string value = MyRegex().Match(freq).Groups["Value"].ToString();
            string unit = MyRegex().Match(freq).Groups["Unit"].ToString();

            double outputValue = 0;
            if (!double.TryParse(value, out double lDValue))
            {
                return 0;
            }

            switch (MyRegex1().Replace(unit, "").ToUpper())
            {
                case "K":
                    outputValue = lDValue * 1e3;
                    break;
                case "M":
                    outputValue = lDValue * 1e6;
                    break;
                case "G":
                    outputValue = lDValue * 1e9;
                    break;
                case "":
                    outputValue = lDValue;
                    break;
            }

            return Math.Abs(outputValue);
        }
    }
}
