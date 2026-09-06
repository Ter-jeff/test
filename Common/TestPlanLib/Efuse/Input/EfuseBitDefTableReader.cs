using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.Efuse.Input
{
    public partial class EfuseBitDefTableReader
    {
        [GeneratedRegex("bank_", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("efuse bit def", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(@"(?<udr>UDR\w+)\s*\t*CMP", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();

        public static List<BitDefTable> Read(ExcelWorksheet excelWorksheet)
        {
            return Read(excelWorksheet.ConvertToLines(), excelWorksheet.Name);
        }

        private static List<BitDefTable> Read(List<string> lines, string sheetName)
        {
            try
            {
                List<BitDefTable> binningTables = ReadSheet(lines, sheetName);
                return binningTables;
            }
            catch (Exception e)
            {
                string msg = "Find exception when reading " + sheetName + "!!! \n ErrMsg: " + e;
                var errMsg = new Exception(msg);
                throw errMsg;
            }
        }

        private static List<BitDefTable> ReadSheet(List<string> lines, string sheetName)
        {
            var bitDefTables = new List<BitDefTable>();
            var bitDefTable = new BitDefTable();
            bitDefTable.Rows.Clear();
            int index = 0;
            while (index < lines.Count)
            {
                ReadData(lines, sheetName, ref bitDefTable, ref index);

                index = ReadData1(lines, sheetName, bitDefTables, bitDefTable, index);
            }
            return bitDefTables;
        }

        private static int ReadData1(List<string> lines, string sheetName, List<BitDefTable> bitDefTables, BitDefTable bitDefTable, int index)
        {
            for (; index < lines.Count; index++)
            {
                string line = lines[index];
                if (line.Split(['\t'], StringSplitOptions.RemoveEmptyEntries).Length == 0)
                {
                    continue;
                }

                if (line.ContainsIgnoreCase("ACCESS MODE") || line.Trim().EqualsIgnoreCase("end"))
                {
                    bitDefTables.Add(bitDefTable);
                    break;
                }
                string[] spt = line.Split(['\t'], StringSplitOptions.None);
                var vLine = spt.Select(token => token.Trim()).ToList();
                var bitDefRow = new BitDefRow { Line = line, RowData = vLine, RowNum = index + 1, SheetName = sheetName };
                bitDefTable.Rows.Add(bitDefRow);
            }

            return index;
        }

        private static void ReadData(List<string> lines, string sheetName, ref BitDefTable bitDefTable, ref int index)
        {
            for (; index < lines.Count; index++)
            {
                string line = lines[index];
                if (line.ContainsIgnoreCase("ACCESS MODE"))
                {
                    bitDefTable = new BitDefTable { SheetName = sheetName, HeaderRowNum = index };
                    string[] lineSpt = line.Split(['\t'], StringSplitOptions.None);
                    bitDefTable.AccessMode = lineSpt.First();
                    bitDefTable.BitMode = bitDefTable.AccessMode.Split(['(', ')', ' '], StringSplitOptions.RemoveEmptyEntries).First(x => x.ContainsIgnoreCase("-BIT"));
                }
                else if (line.ContainsIgnoreCase("EFUSE BIT DEF"))
                {
                    bitDefTable.HeaderRowNum = index + 1;
                    HandleEfuseBitDef(bitDefTable, line);
                    index++;
                    break;
                }
            }
        }

        private static void HandleEfuseBitDef(BitDefTable bitDefTable, string line)
        {
            bitDefTable.Header = line;
            string[] arr = line.Split(['\t'], StringSplitOptions.RemoveEmptyEntries);
            bitDefTable.MaxColIdx = arr.Length;

            for (int i = 0; i < arr.Length; i++)
            {
                string cell = arr[i].ToUpper();
                // 1. Handle the complex "N/A" sequence (Na1 -> Na4)
                if (cell.Contains("N/A") || cell.Contains("NA") ||
                    cell.Contains("USI MSB") || cell.Contains("USI LSB") ||
                    cell.Contains("USO MSB") || cell.Contains("USO LSB"))
                {
                    if (bitDefTable.Na1Idx == -1)
                    {
                        bitDefTable.Na1Idx = i;
                    }
                    else if (bitDefTable.Na2Idx == -1)
                    {
                        bitDefTable.Na2Idx = i;
                    }
                    else if (bitDefTable.Na3Idx == -1)
                    {
                        bitDefTable.Na3Idx = i;
                    }
                    else if (bitDefTable.Na4Idx == -1)
                    {
                        bitDefTable.Na4Idx = i;
                    }
                    // Note: If the specific USI/USO keywords need unique indices 
                    // regardless of NA order, move them to the section below.
                }

                // 2. Handle standard keyword mapping
                HandleEfuseBitDefIndex(bitDefTable, arr, i, cell);
            }
        }

        private static void HandleEfuseBitDefIndex(BitDefTable bitDefTable, string[] arr, int i, string cell)
        {
            if (cell.Contains("EFUSE BIT DEF"))
            {
                bitDefTable.BankEfuseBitDefIdx = i;
                bitDefTable.BlockName = GetBitDefTableBlockName(arr[i]);
            }
            else if (cell.Contains("MSB BIT"))
            {
                bitDefTable.MsbBitIdx = i;
            }
            else if (cell.Contains("LSB BIT"))
            {
                bitDefTable.LsbBitIdx = i;
            }
            else if (cell.Contains("BIT WIDTH"))
            {
                bitDefTable.BitWidthIdx = i;
            }
            else if (cell.Contains("PROGRAMMING STAGE"))
            {
                bitDefTable.ProgrammingStageIdx = i;
            }
            else if (cell.Contains("LOW LIMIT"))
            {
                bitDefTable.LowLimitIdx = i;
            }
            else if (cell.Contains("HIGH LIMIT"))
            {
                bitDefTable.HighLimitIdx = i;
            }
            else if (cell.Contains("RESOLUTION"))
            {
                bitDefTable.ResolutionIdx = i;
            }
            else if (cell.Contains("ALGORITHM"))
            {
                bitDefTable.AlgorithmIdx = i;
            }
            else if (cell.Contains("DESCRIPTION"))
            {
                bitDefTable.DescriptionIdx = i;
            }
            else if (cell.Contains("DEFAULT OR REAL"))
            {
                bitDefTable.DefaultOrRealIdx = i;
            }
            else if (cell.Contains("DEFAULT VALUE"))
            {
                bitDefTable.DefaultValueIdx = i;
            }
            else if (cell.Contains("USI MSB-BIT CYCLE"))
            {
                bitDefTable.UsiMsbBitCycleIdx = i;
            }
            else if (cell.Contains("USI LSB-BIT CYCLE"))
            {
                bitDefTable.UsiLsbBitCycleIdx = i;
            }
            else if (cell.Contains("USO MSB-BIT CYCLE"))
            {
                bitDefTable.UsoMsbBitCycleIdx = i;
            }
            else if (cell.Contains("USO LSB-BIT CYCLE"))
            {
                bitDefTable.UsoLsbBitCycleIdx = i;
            }
        }

        private static string GetBitDefTableBlockName(string block)
        {
            string blockName = "";

            if (block.ContainsIgnoreCase("Bit Def"))
            {
                blockName = MyRegex().Replace(MyRegex1().Split(block)[0].Split('(')[0], "").Trim();
                if (MyRegex2().IsMatch(blockName))
                {
                    string udRnameOrg =
                        MyRegex2().Match(blockName).Groups["udr"].ToString();
                    blockName = udRnameOrg.Replace("UDR", "CMP");
                }
            }
            return blockName;
        }
    }
}
