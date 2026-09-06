using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Automation.GenerateIgxl.HardIp.InputObject;

namespace Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility
{
    public class RegisterConvertor
    {
        public static int GetSendBitLength(string sendBitName, HardIpInfo patternInfo)
        {
            List<string> srcBitNameList = patternInfo.SendBitName.Split('+', '|').ToList();
            List<string> srcBitStrList = patternInfo.SendBitStr.Split('+', '|').ToList();

            for (int i = 0; i < srcBitNameList.Count; i++)
            {
                if (srcBitNameList[i].Equals(sendBitName, StringComparison.OrdinalIgnoreCase))
                {
                    string[] arr = srcBitStrList[i].Split('_');
                    if (arr.Length > 1)
                    {
                        return int.Parse(srcBitStrList[i].Split('_')[1]);
                    }
                    return 0;
                }
            }
            return 0;
        }

        public static string ConvertNumberToSrc(string sendBitName, string assignStr, HardIpInfo patternInfo)
        {
            int iValue = 0;
            string prefix = assignStr.Substring(0, 2).ToLower();
            string assignValue = assignStr.Remove(0, 2);
            int bitLength = GetSendBitLength(sendBitName, patternInfo);

            //if bit length=0, return false
            if (bitLength == 0)
            {
                return "";
            }

            //Gets int value
            try
            {
                switch (prefix)
                {
                    case "0b":
                        iValue = Convert.ToInt32(assignValue, 2);
                        break;
                    case "0x":
                        iValue = Convert.ToInt32(assignValue, 16);
                        break;
                    case "0d":
                        iValue = Convert.ToInt32(assignValue);
                        break;
                }
            }
            catch (Exception)
            {
                return "";
            }

            //convert to binary format
            assignValue = Reverse(Convert.ToString(iValue, 2));
            // Due to 2's complement, modify the bit size according to patInfo
            if (iValue < 0)
            {
                assignValue = assignValue.Substring(0, bitLength);
            }
            //if bit length<assign length, format error.
            if (assignValue.Length > bitLength)
            {
                return assignStr;
            }

            //if bit length>assign length, padding "0"
            assignStr = assignValue.PadRight(bitLength, '0');
            return assignStr;
        }

        private static string Reverse(string input)
        {
            int length = input.Length;
            var sb = new StringBuilder(length);
            for (int i = length - 1; i >= 0; i--)
            {
                sb.Append(input[i]);
            }

            return sb.ToString();
        }
    }
}
