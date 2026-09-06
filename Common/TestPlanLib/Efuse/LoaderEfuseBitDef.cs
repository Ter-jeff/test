using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using LogLib.Utility;

using OfficeOpenXml;

namespace TestPlanLib.Efuse
{
    public partial class LoaderEfuseBitDef(string inPath, string stdf)
    {
        public const int EfuseBefMaxCol = 16;

        [GeneratedRegex("calcBits|ignorebits|one_complement_target", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("Bit Def", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex("^end", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex(@"^end[\w]", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex3();
        [GeneratedRegex("Bit Def", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex4();
        [GeneratedRegex("MSB BIT", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex5();
        [GeneratedRegex("LSB BIT", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex6();
        [GeneratedRegex("Bit Width", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex7();
        [GeneratedRegex("programming stage", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex8();

        [GeneratedRegex("Low Limit", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex9();

        [GeneratedRegex("High Limit", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex10();

        [GeneratedRegex("Resolution", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex11();

        [GeneratedRegex("Algorithm", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex12();

        [GeneratedRegex("Description", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex13();

        [GeneratedRegex("Default or Real", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex14();

        [GeneratedRegex("Default Value", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex15();

        [GeneratedRegex("Difference", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex16();

        [GeneratedRegex("HIP_Name", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex17();

        [GeneratedRegex("EQ", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex18();

        [GeneratedRegex("Access Mode", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex19();

        [GeneratedRegex(@"(?<udr>UDR\w+)\s*\t*CMP", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex20();

        [GeneratedRegex("bank_", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex21();

        [GeneratedRegex("efuse bit def", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex22();

        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();
        private static readonly Regex _regex3 = MyRegex2();
        private static readonly Regex _regex4 = MyRegex3();
        private static readonly Regex _regex5 = MyRegex1();
        private static readonly Regex _regex6 = MyRegex2();
        private static readonly Regex _regex7 = MyRegex3();
        private static readonly Regex _regex8 = MyRegex4();
        private static readonly Regex _regex9 = MyRegex5();
        private static readonly Regex _regex10 = MyRegex6();
        private static readonly Regex _regex11 = MyRegex7();
        private static readonly Regex _regex12 = MyRegex8();
        private static readonly Regex _regex13 = MyRegex9();
        private static readonly Regex _regex14 = MyRegex10();
        private static readonly Regex _regex15 = MyRegex11();
        private static readonly Regex _regex16 = MyRegex12();
        private static readonly Regex _regex17 = MyRegex13();
        private static readonly Regex _regex18 = MyRegex14();
        private static readonly Regex _regex19 = MyRegex15();
        private static readonly Regex _regex20 = MyRegex16();
        private static readonly Regex _regex21 = MyRegex17();
        private static readonly Regex _regex22 = MyRegex18();
        private static readonly Regex _regex23 = MyRegex19();
        private static readonly Regex _regex24 = MyRegex20();
        private static readonly Regex _regex25 = MyRegex21();
        private static readonly Regex _regex26 = MyRegex4();
        private static readonly Regex _regex27 = MyRegex22();

        public string InPath = inPath;
        public EfuseBitDefTable BitDefTable = new();
        public string StdfFile = stdf;

        public void Parse(EfuseScriptConfig efuseScriptConfig)
        {
            try
            {
                BitDefTable.Titles.Clear();
                BitDefTable.Rows.Clear();
                BitDefTable.CrcItem.Clear();

                if (Path.GetExtension(InPath).EqualsIgnoreCase(".txt"))
                {
                    ReadBdfTxt(InPath, efuseScriptConfig);
                }
                else
                {
                    ReadBdFxls(InPath, efuseScriptConfig);
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        public static void UpdateCrcDescription(CrcItem crcItem, EfuseScriptConfig efuseScriptConfig)
        {
            if (!_regex.IsMatch(crcItem.Description))
            {
                if (efuseScriptConfig.CrCdescription.TryGetValue(crcItem.BlockName, out string? value))
                {
                    crcItem.Description = value;
                }
            }
            else
            {
                efuseScriptConfig.CrCdescription[crcItem.BlockName] = crcItem.Description;
            }
        }

        //This function is to read BDF txt file
        private void ReadBdfTxt(string path, EfuseScriptConfig efuseScriptConfig)
        {
            string? line;
            try
            {
                using StreamReader sr = new StreamReader(path);
                //STEP1. GetTitle and specific field index(eg. CPGB/CP2GB/FT1GB/FT2GB/ATE_FQAGB)
                //-----------------------------------------------------------
                while (!sr.EndOfStream)
                {
                    //Line#3>   Domain	ID	Mode	EQN	C	M	CPIDSMax	CPVmax	CPVmin	CPGB	CP2GB	FT1GB	FT2GB	FTIDS	SLTGB	ATE_FQAGB	HTOL_RO_GB	SLT_FQA_GB	PMUMax	PMUMin	CPHV	FTHV	QAHV	Comment	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail
                    line = sr.ReadLine();
                    if (_regex2.IsMatch(line!))
                    {
                        ReadHeader(line!);
                        break;
                    }
                }

                sr.DiscardBufferedData();
                sr.BaseStream.Seek(0, SeekOrigin.Begin);
                //STEP2. Get BitDef table
                //-----------------------------------------------------------

                string curBlock = "";
                string curAccessMode = "";
                while (!sr.EndOfStream)
                {
                    line = sr.ReadLine()!.ToUpper();
                    //CPUSRAM	MCS601	MC601	800	715.625	50	75	50	40.625	4	40	12.5	75	12.5				10	84.375	0	0.621875	0.525	

                    //3b. Check if AccessMode need to update, if update=> just continue;
                    if (IsUpdateAccessMode(ref curAccessMode, line))
                    {
                        continue;
                    }

                    //3c. another fuse block, just continue;
                    if (IsUpdateBlockName(ref curBlock, line))
                    {
                        continue;
                    }

                    //3d. empty, just continue
                    //if (line.Split(new char[] { '\t' }, StringSplitOptions.RemoveEmptyEntries).Length == 0)
                    //    continue;
                    if (!CheckValidLine(line))
                    {
                        continue;
                    }

                    //3a. end line
                    if (_regex3.IsMatch(line))
                    {
                        if (!_regex4.IsMatch(line))
                        {
                            //End	End	End
                            break;
                        }
                    }
                    //3e. Access BDF Content

                    List<string> vLine = ReadContent(line, curBlock, curAccessMode);

                    if (!string.IsNullOrEmpty(vLine[0].Trim()))
                    {
                        BitDefTable.Rows.Add(vLine);
                    }

                    if (!BitDefTable.BlockList.Contains(curBlock))
                    {
                        BitDefTable.BlockList.Add(curBlock);
                    }

                    UpdateCrcInfo(vLine, efuseScriptConfig);
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        //This function is to read BDF xls file
        private void ReadBdFxls(string path, EfuseScriptConfig efuseScriptConfig)
        {
            string line;
            using FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using ExcelPackage ep = new ExcelPackage(fs);
            ExcelWorksheet ws = ep.Workbook.Worksheets["EFUSE_BitDef_Table"];
            int startrow = 1;
            int endrow = ws.Dimension.End.Row;
            while (startrow != endrow)
            {
                //Line#3>   Domain	ID	Mode	EQN	C	M	CPIDSMax	CPVmax	CPVmin	CPGB	CP2GB	FT1GB	FT2GB	FTIDS	SLTGB	ATE_FQAGB	HTOL_RO_GB	SLT_FQA_GB	PMUMax	PMUMin	CPHV	FTHV	QAHV	Comment	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	Binning Fail	LVCC Fail	
                line = ws.GetCellLine(startrow);
                startrow++;
                if (_regex5.IsMatch(line))
                {
                    ReadHeader(line);
                    break;
                }
            }

            //STEP2. Get BitDef table
            //-----------------------------------------------------------

            string curBlock = "";
            string curAccessMode = "";
            startrow = 1;
            try
            {
                while (startrow != endrow)
                {
                    line = ws.GetCellLine(startrow).ToUpper();
                    //CPUSRAM	MCS601	MC601	800	715.625	50	75	50	40.625	4	40	12.5	75	12.5				10	84.375	0	0.621875	0.525	
                    startrow++;

                    //3b. Check if AccessMode need to update, if update=> just continue;
                    if (IsUpdateAccessMode(ref curAccessMode, line))
                    {
                        continue;
                    }

                    //3c. another fuse block, just continue;
                    if (IsUpdateBlockName(ref curBlock, line))
                    {
                        continue;
                    }

                    //3d. empty, just continue
                    //if (line.Split(new char[] {'\t'}, StringSplitOptions.RemoveEmptyEntries).Length == 0)
                    //    continue;
                    if (!CheckValidLine(line))
                    {
                        continue;
                    }

                    //3a. end line
                    if (_regex6.IsMatch(line))
                    {
                        if (!_regex7.IsMatch(line))
                        {
                            //End	End	End
                            break;
                        }
                    }
                    //3e. Access BDF Content

                    List<string> vLine = ReadContent(line, curBlock, curAccessMode);

                    BitDefTable.Rows.Add(vLine);
                    if (!BitDefTable.BlockList.Contains(curBlock))
                    {
                        BitDefTable.BlockList.Add(curBlock);
                    }

                    UpdateCrcInfo(vLine, efuseScriptConfig);
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
            ////////////////////////////////////////////

        }

        public static string[] ReArrangeLin(string[] line, int length)
        {
            if (line.GetLength(1) < length)
            {
                string[] strout = [];
                for (int i = 0; i < length; i++)
                {
                    strout[i] = line[i];
                }

                return strout;
            }

            return line;
        }

        private void ReadHeader(string line)
        {
            int lineLen = EfuseBefMaxCol;
            if (line.Contains("hip_name", StringComparison.OrdinalIgnoreCase))
            {
                lineLen++;
            }

            if (line.Contains("eq", StringComparison.OrdinalIgnoreCase))
            {
                lineLen++;
            }

            string[] lineSpt = line.Split(['\t'], StringSplitOptions.None);
            BitDefTable.Titles.Clear();
            if (Array.IndexOf(lineSpt, "Difference") != -1)
            {
                lineLen = Array.IndexOf(lineSpt, "Difference");
                BitDefTable.DifferentIdx = lineLen;
            }
            //int lineLen = (EFUSE_BEF_MAX_COL < lineSpt.Length) ? EFUSE_BEF_MAX_COL : lineSpt.Length;
            //lineSpt = ReArrangeLin(lineSpt, lineLen);
            Array.Resize(ref lineSpt, lineLen);
            for (int i = 0; i < lineLen; i++)
            {
                BitDefTable.Titles.Add(lineSpt[i]);

                if (_regex8.IsMatch(lineSpt[i]))
                {
                    BitDefTable.NameIdx = i;
                }
                else if (_regex9.IsMatch(lineSpt[i]))
                {
                    BitDefTable.MsbBitIdx = i;
                }
                else if (_regex10.IsMatch(lineSpt[i]))
                {
                    BitDefTable.LsbBitIdx = i;
                }
                else if (_regex11.IsMatch(lineSpt[i]))
                {
                    BitDefTable.BitWidthIdx = i;
                }
                else if (_regex12.IsMatch(lineSpt[i]))
                {
                    BitDefTable.PrgStageIdx = i;
                }
                else if (_regex13.IsMatch(lineSpt[i]))
                {
                    BitDefTable.LowLimitIdx = i;
                }
                else if (_regex14.IsMatch(lineSpt[i]))
                {
                    BitDefTable.HighLimitIdx = i;
                }
                else if (_regex15.IsMatch(lineSpt[i]))
                {
                    BitDefTable.ResolutionIdx = i;
                }
                else if (_regex16.IsMatch(lineSpt[i]))
                {
                    BitDefTable.AlgorithmIdx = i;
                }
                else if (_regex17.IsMatch(lineSpt[i]))
                {
                    BitDefTable.DescriptionIdx = i;
                }
                else if (_regex18.IsMatch(lineSpt[i]))
                {
                    BitDefTable.DefaultOrRealIdx = i;
                }
                else if (_regex19.IsMatch(lineSpt[i]))
                {
                    BitDefTable.DefaultValueIdx = i;
                }
                else if (_regex20.IsMatch(lineSpt[i]))
                {
                    BitDefTable.DifferentIdx = i;
                }
                else if (_regex21.IsMatch(lineSpt[i]))
                {
                    BitDefTable.HipNameIdx = i;
                }
                else if (_regex22.IsMatch(lineSpt[i]))
                {
                    BitDefTable.HipEquationIdx = i;
                }
            }
            if (BitDefTable.DifferentIdx == -1)
            {
                BitDefTable.DifferentIdx = BitDefTable.Titles.Count;
            }

            BitDefTable.Titles.Add("Block");

            BitDefTable.BlockIdx = BitDefTable.Titles.Count - 1;
            BitDefTable.Titles.Add("AccessMode");
            BitDefTable.AccessModeIdx = BitDefTable.Titles.Count - 1;
            if (line.Contains("hip_name", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(StdfFile))
            {
                BitDefTable.HiplimithIdx = BitDefTable.Titles.Count;
                BitDefTable.Titles.Add("HIPHighL");
                BitDefTable.HiplimitlIdx = BitDefTable.Titles.Count;
                BitDefTable.Titles.Add("HIPLowL");
            }
        }

        // check whether access mode need to update
        private static bool IsUpdateAccessMode(ref string curAccessMode, string line)
        {
            List<string> mode = [.. line.Split(['\t'], StringSplitOptions.RemoveEmptyEntries)];
            if (mode.Count > 0 && _regex23.IsMatch(mode[0]))
            {
                curAccessMode = mode[0];
                return true;
            }
            return false;
        }

        // check whether block name need to update
        private static bool IsUpdateBlockName(ref string curBlock, string line)
        {
            List<string> block = [.. line.Split(['\t'], StringSplitOptions.RemoveEmptyEntries)];
            if (block.Count > 0 && _regex26.IsMatch(block[0]))
            {
                if (block.Count != 0)
                {
                    curBlock = _regex25.Replace(_regex27.Split(block[0])[0].Split('(')[0], "").Trim();
                    if (_regex24.IsMatch(curBlock))
                    {
                        string udRnameOrg = _regex24.Match(curBlock).Groups["udr"].ToString();
                        curBlock = udRnameOrg.Replace("UDR", "CMP");
                    }
                }
                return true;
            }
            return false;
        }

        private List<string> ReadContent(string line, string curBlock, string curAccessMode)
        {
            string[] spt = line.Split(['\t'], StringSplitOptions.None);

            List<string> vLine = [];
            //foreach (string token in spt)
            //    vLine.Add(token.Trim());
            for (int i = 0; i < BitDefTable.DifferentIdx; i++)
            {
                if (i < spt.Length)
                {
                    vLine.Add(spt[i].Trim());
                }
            }
            // temp generate for Cebu
            //if (!string.IsNullOrEmpty(temp_vdd_dcs_old) && !temp_vdd_dcs_old.ToLower().Contains("vdd_dcs_ddr"))
            //{
            //    temp_vdd_dcs_new = temp_vdd_dcs_old.ToLower().Replace("vdd_dcs", "vdd_dcs_ddr");
            //    if (vLine.Exists(p => p.ToLower().Contains("bincut") || p.ToLower().Contains("ids")))
            //        vLine[0] = temp_vdd_dcs_new.ToUpper();
            //}

            vLine.Add(curBlock);
            vLine.Add(curAccessMode);
            if (BitDefTable.HipNameIdx != -1 && !string.IsNullOrEmpty(vLine[BitDefTable.HipNameIdx]))
            {
                BitDefTable.HipList.Add(vLine[BitDefTable.HipNameIdx], vLine[BitDefTable.HipEquationIdx]);
            }

            return vLine;
        }

        private void UpdateCrcInfo(List<string> vLine, EfuseScriptConfig efuseScriptConfig)
        {
            if (vLine[BitDefTable.AlgorithmIdx].EqualsIgnoreCase("crc"))
            {
                var crc = new CrcItem()
                {
                    FieldName = vLine[BitDefTable.NameIdx],
                    Blockindx = BitDefTable.BlockIdx,
                    BlockName = vLine[BitDefTable.BlockIdx],
                    Description = vLine[BitDefTable.DescriptionIdx],
                    CrcMethod = vLine[BitDefTable.BitWidthIdx],
                    Job = vLine[BitDefTable.PrgStageIdx]
                };
                crc.CalcBitwidth(int.Parse(vLine[BitDefTable.MsbBitIdx]), int.Parse(vLine[BitDefTable.LsbBitIdx]));
                UpdateCrcDescription(crc, efuseScriptConfig);
                BitDefTable.CrcItem.Add(crc);
            }
        }

        private static bool CheckValidLine(string line)
        {
            string[] spt = line.Split(['\t'], StringSplitOptions.RemoveEmptyEntries);
            if (spt.Length <= 1)
            {
                return false;
            }

            if (spt.Length >= EfuseBefMaxCol)
            {
                for (int i = 0; i < 10; i++)
                {
                    if (spt[i].Length == 0)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public Dictionary<string, string> GetIds()
        {
            var dic = new Dictionary<string, string>();
            foreach (List<string> row in BitDefTable.Rows)
            {
                if (row[0].StartsWithIgnoreCase("IDS"))
                {
                    if (!dic.ContainsKey(row[0]))
                    {
                        dic.Add(row[0], row[3]);
                    }
                }
            }
            return dic;
        }

        public Dictionary<string, string> GetIdsForEfuseCrossCheck(string stage)
        {
            var efusebitDef = new Dictionary<string, string>();
            foreach (List<string> row in BitDefTable.Rows)
            {
                if (!row[BitDefTable.PrgStageIdx].EqualsIgnoreCase(stage) ||
                    !row[BitDefTable.AlgorithmIdx].EqualsIgnoreCase("ids"))
                {
                    continue;
                }

                efusebitDef.Add(row[BitDefTable.NameIdx].ToUpper(), row[BitDefTable.HighLimitIdx]);
            }
            return efusebitDef;
        }

        public Dictionary<string, List<string>> GetVdd()
        {
            var dic = new Dictionary<string, List<string>>();
            foreach (List<string> row in BitDefTable.Rows)
            {
                if (row[12].StartsWithIgnoreCase("vddbin"))
                {
                    dic.TryAdd(row[0], row);
                }
            }
            return dic;
        }

        public string GetBaseVoltage()
        {
            return BitDefTable.Rows.Find(y => y[12].EqualsIgnoreCase("base"))!.ElementAt(10);
        }
    }
}
