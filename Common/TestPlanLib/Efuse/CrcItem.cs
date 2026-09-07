using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using LogLib.Utility;

namespace TestPlanLib.Efuse
{
    public partial class CrcItem
    {
        [GeneratedRegex("calcBits|ignorebits", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"\[\d+:\d+\]", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(@"\[(?<value1>\d+):(?<value2>\d+)\]", RegexOptions.Compiled)]
        private static partial Regex MyRegex2();
        [GeneratedRegex("ignorebits", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex3();

        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();
        private static readonly Regex _regex3 = MyRegex2();
        private static readonly Regex _regex4 = MyRegex3();

        public string FieldName = "";
        public string BlockName = "";
        public int Blockindx = -1;
        public string RawBitstream = "";
        public string Description = "";
        public string CrcMethod = "";
        public int Msb;
        public int Lsb;
        public int Bitwidth;
        public string Job = "";
        public int X;
        public int Y;
        // Additional use for generate another file
        public string LotId = "";
        public string WaferId = "";
        public string DsscX = "";
        public string DsscY = "";
        public string Site = "";
        public string FilteredBitstream = "";
        public string CalcultaedCrc = "";
        public string ExpectedCrc = "";
        public int BinNum;
        //use for OTP addr
        public int Offset = 0;

        public static string CalcCrcWithOneComplement(string binaryValue, int bitWidth)
        {
            string output;
            try
            {
                long oneComplement = ~Convert.ToInt32(binaryValue, 2);
                string tmpStr = Convert.ToString(oneComplement, 2);
                tmpStr = tmpStr.Length > 32 ? tmpStr.Substring(tmpStr.Length - 32, 32) : tmpStr;
                output = HexConverted(tmpStr, bitWidth);
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                output = "Error occurred during calculate CRC one complement";
            }
            return output;
        }

        public string CalcCrc()
        {
            string output;
            try
            {
                //ConfigMask 
                List<bool> bitMask = ConfigMask(RawBitstream.Length);
                FilteredBitstream = FilterBitStream(RawBitstream, bitMask);
                long[] bitArray = Str2bitarray_bitinv(FilteredBitstream);
                output = HexConverted(Auto_CRC_BinStr(bitArray, bitArray.Length, Bitwidth), Bitwidth);
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                output = "Error occurred during calculate CRC";
            }
            return output;
        }

        // Setup Mask for filtering
        public List<bool> ConfigMask(int size = 240)
        {
            var mask = new List<bool>();
            for (int i = 0; i < size; i++)
            {
                mask.Add(false);
            }

            //determine "calc" or "ignore" to concern bit_inv or not
            //ignore :  inverse
            if (_regex.IsMatch(Description))
            {
                MatchCollection matech = _regex2.Matches(Description);
                //enable specify interval with true (default is false)
                foreach (object match in matech)
                {
                    int value1 = int.Parse(_regex3.Match(match.ToString()!).Groups["value1"].Value) - Offset;
                    int value2 = int.Parse(_regex3.Match(match.ToString()!).Groups["value2"].Value) - Offset;
                    int startindx, stopindx;
                    if (value1 > value2)
                    {
                        startindx = value2;
                        stopindx = value1;
                    }
                    else
                    {
                        startindx = value1;
                        stopindx = value2;
                    }
                    for (int i = startindx; i <= stopindx; i++)
                    {
                        if (size - i - 1 >= 0)
                        {
                            mask[size - i - 1] = true;
                        }
                    }
                }
                //if flag is ignore
                if (!_regex4.IsMatch(Description))
                {
                    return mask;
                }

                for (int i = 0; i < mask.Count; i++)
                {
                    mask[i] = !mask[i];
                }
            }
            else
            {
                int stopindx;
                if (Msb > Lsb)
                {
                    stopindx = Lsb;
                }
                else
                {
                    stopindx = Msb;
                }
                for (int i = 0; i <= stopindx - 1; i++)
                {
                    mask[size - i - 1] = true;
                }
            }

            return mask;
        }

        // Filter bitstream with mask 
        public static string FilterBitStream(string bitstring, List<bool> mask, bool valDef = true)
        {
            string result = "";
            for (int i = 0; i < mask.Count; i++)
            {
                bool flag = mask[i];
                if (flag)
                {
                    result += bitstring[i];
                }
            }
            return result;
        }

        // CRC16 main VBT Name :  "auto_ECID_CRC2HexStr"
        public string Auto_CRC_BinStr(long[] bitArray, long crcEndBit, int crcBitLength)
        {
            string result = "";
            byte[] crCarray = new byte[crcBitLength];
            for (int i = 0; i < crCarray.Length; i++)
            {
                crCarray[i] = 0;
            }

            for (long i = crcEndBit - 1; i > -1; i--)
            {
                if (crcBitLength == 16)
                {
                    CRC16_ComputeCRCforBit(ref crCarray, (byte)bitArray[i]);
                }
                else
                {
                    CRC32_ComputeCRCforBit(ref crCarray, (byte)bitArray[i]);
                }
            }
            for (int i = 0; i < crCarray.Length; i++)
            {
                result = crCarray[i] + result;
            }

            CalcultaedCrc = result;
            return result;
        }

        // Binary string convert to Hex format
        public static string HexConverted(string strBinary, int bytelength)
        {
            string hexstr;
            if (bytelength == 32)
            {
                hexstr = "X8";
            }
            else
            {
                hexstr = "X4";
            }

            string strHex = Convert.ToInt32(strBinary, 2).ToString(hexstr);
            return "0x" + strHex;
        }

        // Calculate CRC16 
        private static void CRC16_ComputeCRCforBit(ref byte[] crc, byte bit)
        {
            byte inv = (byte)(bit ^ crc[15]);
            crc[15] = crc[14];
            crc[14] = crc[13];
            crc[13] = (byte)(crc[12] ^ inv);
            crc[12] = (byte)(crc[11] ^ inv);
            crc[11] = (byte)(crc[10] ^ inv);
            crc[10] = (byte)(crc[9] ^ inv);
            crc[9] = crc[8];
            crc[8] = (byte)(crc[7] ^ inv);
            crc[7] = crc[6];
            crc[6] = (byte)(crc[5] ^ inv);
            crc[5] = (byte)(crc[4] ^ inv);
            crc[4] = crc[3];
            crc[3] = crc[2];
            crc[2] = (byte)(crc[1] ^ inv);
            crc[1] = crc[0];
            crc[0] = inv;
        }

        // Calculate CRC32
        private static void CRC32_ComputeCRCforBit(ref byte[] crc, byte bit)
        {
            byte inv = (byte)(bit ^ crc[31]);
            crc[31] = crc[30];
            crc[30] = crc[29];
            crc[29] = crc[28];
            crc[28] = crc[27];
            crc[27] = crc[26];
            crc[26] = (byte)(crc[25] ^ inv);
            crc[25] = crc[24];
            crc[24] = crc[23];
            crc[23] = (byte)(crc[22] ^ inv);
            crc[22] = (byte)(crc[21] ^ inv);
            crc[21] = crc[20];
            crc[20] = crc[19];
            crc[19] = crc[18];
            crc[18] = crc[17];
            crc[17] = crc[16];
            crc[16] = (byte)(crc[15] ^ inv);
            crc[15] = crc[14];
            crc[14] = crc[13];
            crc[13] = crc[12];
            crc[12] = (byte)(crc[11] ^ inv);
            crc[11] = (byte)(crc[10] ^ inv);
            crc[10] = (byte)(crc[9] ^ inv);
            crc[9] = crc[8];
            crc[8] = (byte)(crc[7] ^ inv);
            crc[7] = (byte)(crc[6] ^ inv);
            crc[6] = crc[5];
            crc[5] = (byte)(crc[4] ^ inv);
            crc[4] = (byte)(crc[3] ^ inv);
            crc[3] = crc[2];
            crc[2] = (byte)(crc[1] ^ inv);
            crc[1] = (byte)(crc[0] ^ inv);
            crc[0] = inv;
        }

        // Convert Binary String to integer array for CRC

        // Convert and invert Binary String to integer array for CRC
        private static long[] Str2bitarray_bitinv(string input)
        {
            var output = new List<long>();
            foreach (char str in input)
            {
                output.Insert(0, (long)char.GetNumericValue(str));
            }
            return [.. output];
        }

        public void CalcBitwidth(int msb, int lsb)
        {
            if (msb > lsb)
            {
                Msb = msb;
                Lsb = lsb;
            }
            else
            {
                Msb = lsb;
                Lsb = msb;
            }
            Bitwidth = Msb - Lsb + 1;
        }
    }
}
