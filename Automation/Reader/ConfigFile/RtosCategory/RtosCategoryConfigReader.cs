using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;



namespace Automation.Reader.ConfigFile.RtosCategory
{
    public class RtosCategoryConfigReader
    {
        private readonly bool _useCategoryOffset = false;

        public List<RtosConfigBitRow> Rows { set; get; } = new List<RtosConfigBitRow>();

        public int KeywordFieldNumber { set; get; } = 12;


        public string GetCategoryName(string source)
        {
            if (string.IsNullOrEmpty(source) || !Regex.IsMatch(source, "^[0-9A-Fa-f]+$"))
            {
                return "";
            }

            string bynaryStr = Convert.ToString(Convert.ToInt64(source, 16), 2);
            bynaryStr = bynaryStr.PadLeft(source.Length * 4, '0');
            string modeStr = "";
            foreach (RtosConfigBitRow row in Rows)
            {
                if (!row.Use)
                {
                    continue;
                }

                int maxBit = row.GetMaxBit();
                int bitNum = row.BitNum();
                row.GetMinBit();
                if (bynaryStr.Length < maxBit)
                {
                    return "";
                }

                string subStr = bynaryStr.Substring(bynaryStr.Length - maxBit - 1, bitNum);
                string modeNum = Convert.ToInt32(subStr, 2).ToString("x");  //4位2进制数转化为1位16进制数
                if (_useCategoryOffset)
                {
                    modeNum = row.PerformanceOffset(Convert.ToInt32(subStr, 2)).ToString("x");
                }

                modeStr += modeNum;
            }

            return modeStr;

        }
    }

    public class RtosConfigBitRow
    {
        #region Field

        private readonly string _bit;
        private readonly string _performanceMode;
        private readonly string _powerPins;
        private readonly string _use;
        private readonly string _offsetPmode;
        #endregion

        #region Properity
        public string PerformanceMode { set; get; }
        public List<string> PowerPins { get { return _powerPins.Split(',').ToList(); } }

        public bool Use
        {
            get
            {
                if (_use.Equals("Y", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (_use.Equals("N", StringComparison.OrdinalIgnoreCase))
                {
                }

                return false;
            }
        }

        #endregion

        #region Constructor
        public RtosConfigBitRow(string bit, string performanceMode, string powerPins, string use, string offsetPMode)
        {
            _bit = bit;
            _powerPins = powerPins;
            _performanceMode = performanceMode;
            _use = use;
            _offsetPmode = offsetPMode;
        }

        #endregion

        public int GetMaxBit()
        {
            TryGetBit(out int num1, out int num2);
            return Math.Max(num1, num2);
        }

        public int GetMinBit()
        {
            TryGetBit(out int num1, out int num2);
            return Math.Min(num1, num2);
        }

        public int BitNum()
        {
            int num = GetMaxBit() - GetMinBit() + 1;
            return num;
        }

        private void TryGetBit(out int bit1, out int bit2)
        {
            string regexPattern = @"\[(?<num1>\d+)\:(?<num2>\d+)\]";
            string num1 = Regex.Match(_bit, regexPattern, RegexOptions.IgnoreCase).Groups["num1"].ToString();
            string num2 = Regex.Match(_bit, regexPattern, RegexOptions.IgnoreCase).Groups["num2"].ToString();
            bit1 = int.Parse(num1);
            bit2 = int.Parse(num2);

        }

        public string GetMode(string bynaryStr)
        {
            int modeNum = Convert.ToInt32(bynaryStr, 2);
            string formatStr = Regex.Match(_performanceMode, @"\w+\d+(?<str>X+)", RegexOptions.IgnoreCase).Groups["str"].ToString();
            string format = Regex.Replace(formatStr, "x", "0", RegexOptions.IgnoreCase);
            string mode = PerformanceOffset(modeNum).ToString(format);
            string targetMode = _performanceMode.Replace(formatStr, mode);

            return targetMode;
        }

        public int PerformanceOffset(int orgModeNum)
        {
            if (!string.IsNullOrEmpty(_offsetPmode))
            {
                return orgModeNum + Convert.ToInt32(_offsetPmode, 2);
            }
            return orgModeNum;
        }
    }

}
