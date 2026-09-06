using System;
using System.Linq;

using CommonLib.Datalog;
using CommonLib.Extension;

namespace EfuseCheckCmdLib.Datalog
{
    public class EfuseLine : LineBase
    {
        public EfuseLine(string line, int lineNo)
        {
            Line = line;
            LineNo = lineNo;
        }

        public EfuseRow? ToRow(bool hasReadWaferData)
        {
            try
            {
                return hasReadWaferData
                    ? ParseReadWaferLine()
                    : ParseEfuseLine();
            }
            catch
            {
                return null;
            }
        }

        private EfuseRow? ParseEfuseLine()
        {
            string[] arr = Line.Split(["(", ")", "Values", "(MSB)", ":", "(LSB)", "="], StringSplitOptions.None);
            if (arr.Length < 9 || arr.Length > 13)
            {
                return null;
            }

            try
            {
                string siteNum = arr.ElementAtOrDefault(1)?.Trim()!;
                string bankConfig = arr.ElementAtOrDefault(3)?.Trim()!;
                if (!int.TryParse(arr.ElementAtOrDefault(5)?.Trim(), out int msbValue) || !int.TryParse(arr.ElementAtOrDefault(6)?.Trim(), out int lsbValue))
                {
                    return null;
                }

                string[] testStageLotId = arr[8].Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (testStageLotId.Length < 2)
                {
                    return null;
                }

                string programmingStage = testStageLotId[0];
                string subConfig = testStageLotId[1];
                string data = "";
                string hexValue = "";
                string binValue = "";
                string decValue = "";
                string bits = "";
                string sortedBits = "";
                if (subConfig.ContainsIgnoreCase("lot_id"))
                {
                    HandleLotId(arr, msbValue, lsbValue, out data, out binValue, out bits, out sortedBits);
                }
                else if (arr.Length > 10)
                {
                    HandleLengthBigerThan10(arr, msbValue, lsbValue, out hexValue, ref binValue, ref decValue, out bits, out sortedBits);
                }
                else if (arr.Length == 10)
                {
                    bits = arr.ElementAtOrDefault(9)?.Trim()!;
                    sortedBits = msbValue > lsbValue ? arr.ElementAtOrDefault(9)?.Trim()! : new string(arr.ElementAtOrDefault(9)?.Trim().Reverse().ToArray());
                    binValue = "b" + bits;
                    decValue = EfuseCmdUtility.BinaryStringToDecimal(bits);
                    hexValue = "x" + EfuseCmdUtility.BinaryStringToHexString(bits);
                }
                _ = int.TryParse(siteNum, out int site);
                return new EfuseRow
                {
                    Site = site,
                    SubConfig = subConfig,
                    BankConfig = bankConfig,
                    MsbBit = msbValue,
                    LsbBit = lsbValue,
                    BitWidth = Math.Abs(lsbValue - msbValue) + 1,
                    ProgrammingStage = programmingStage,
                    HexValue = hexValue,
                    BinValue = binValue,
                    DecValue = decValue,
                    Bits = bits,
                    SortedBits = sortedBits,
                    Data = data,
                    Line = this,
                };
            }
            catch
            {
                return null;
            }
        }

        private EfuseRow? ParseReadWaferLine()
        {
            string[] arr = Line.Trim().Split([' '], StringSplitOptions.RemoveEmptyEntries);
            if (arr.Length < 4)
            {
                return null;
            }
            string bits;
            string siteNum = arr.ElementAtOrDefault(1)!;
            _ = int.TryParse(siteNum, out int site);
            string subConfig = arr.ElementAtOrDefault(2)!;

            if (subConfig.ContainsIgnoreCase("LotID"))
            {
                bits = ConvertStrToBinary(arr.ElementAtOrDefault(3)!, 36);
            }
            else if (subConfig.ContainsIgnoreCase("_X") || subConfig.ContainsIgnoreCase("_Y"))
            {
                bits = DecimalToBinary((int)double.Parse(arr.ElementAtOrDefault(5)!), 6);
            }
            else
            {
                bits = DecimalToBinary((int)double.Parse(arr.ElementAtOrDefault(3)!), 5);
            }

            return new EfuseRow
            {
                Site = site,
                SubConfig = subConfig,
                BankConfig = "",
                MsbBit = 0,
                LsbBit = 0,
                BitWidth = 0,
                ProgrammingStage = "",
                HexValue = "",
                BinValue = "",
                DecValue = "",
                Bits = bits,
                SortedBits = "",
                Data = "",
                Line = this,
            };

        }

        public static string ConvertStrToBinary(string value, int bitWidth)
        {
            ulong result = 0UL;

            if (value == "0")
            {
                return new string('0', bitWidth);
            }

            int charCount = value.Length;
            for (int i = 0; i < charCount; i++)
            {
                ulong charValue = value[i];

                if (charValue >= '0' && charValue <= '9')
                {
                    charValue -= 48;
                }
                else
                {
                    charValue -= 55;
                }

                result |= charValue << ((charCount - 1 - i) * 6);
            }
            return Convert.ToString((long)result, 2).PadLeft(bitWidth, '0');
        }

        public static string DecimalToBinary(int value, int bitWidth)
        {
            return Convert.ToString(value, 2).PadLeft(bitWidth, '0');
        }

        private static void HandleLengthBigerThan10(string[] arr, int msbValue, int lsbValue, out string hexValue, ref string binValue, ref string decValue, out string bits, out string sortedBits)
        {
            hexValue = arr.ElementAtOrDefault(9)?.Trim()!;
            if (arr.Length > 10)
            {
                if (msbValue == 0)
                {
                    binValue = arr.ElementAtOrDefault(10)?.Trim()!;
                    decValue = arr.ElementAtOrDefault(11)?.Trim()!;
                }
                else
                {
                    decValue = arr.ElementAtOrDefault(10)?.Trim()!;
                    binValue = arr.ElementAtOrDefault(11)?.Trim()!;
                }
            }
            bits = binValue.TrimStart('b');
            sortedBits = msbValue > lsbValue ? binValue.TrimStart('b') : new string([.. binValue.TrimStart('b').Reverse()]);
        }

        private static void HandleLotId(string[] arr, int msbValue, int lsbValue, out string data, out string binValue, out string bits, out string sortedBits)
        {
            data = arr.ElementAtOrDefault(9)?.Trim()!;
            binValue = arr.ElementAtOrDefault(Array.FindIndex(arr, str => str.Trim().StartsWith('b')))?.Trim()!;
            bits = binValue.TrimStart('b');
            sortedBits = msbValue > lsbValue ? binValue.TrimStart('b') : new string([.. binValue.TrimStart('b').Reverse()]);
        }
    }
}
