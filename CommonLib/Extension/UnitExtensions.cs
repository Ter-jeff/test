using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace CommonLib.Extension
{
    public static class UnitExtensions
    {
        private const string UNIT_AMPERE = "A";
        private const string UNIT_MINI_AMPERE = "mA";
        private const string UNIT_MICRO_AMPERE = "uA";
        private const string UNIT_NANO_AMPERE = "nA";
        private const string UNIT_PICO_AMPERE = "pA";
        private const string UNIT_FEMTO_AMPERE = "fA";

        private const string UNIT_VOLT = "V";
        private const string UNIT_MINI_VOLT = "mV";
        private const string UNIT_MICRO_VOLT = "uV";

        private const string UNIT_HZ = "Hz";
        private const string UNIT_KHZ = "KHz";
        private const string UNIT_MHZ = "MHz";
        private const string UNIT_GHZ = "GHz";

        private static readonly Regex _regex = new Regex(@"^([+-]?\d+\.?\d*)\*?([a-zA-Z]|%|\/)*$", RegexOptions.Compiled);
        private static readonly Regex _regex2 = new Regex(@"^([+-]?\d+\.?\d*[Ee][+-]?\d+)\*?([a-zA-Z]|%|\/)*$", RegexOptions.Compiled);
        private static readonly Regex _regex9 = new Regex("^m.*", RegexOptions.Compiled);
        private static readonly Regex _regex10 = new Regex("^u.*", RegexOptions.Compiled);
        private static readonly Regex _regex11 = new Regex("^n.*", RegexOptions.Compiled);
        private static readonly Regex _regex12 = new Regex("^p.*", RegexOptions.Compiled);
        private static readonly Regex _regex13 = new Regex("^f.*", RegexOptions.Compiled);
        private static readonly Regex _regex14 = new Regex("^k.*", RegexOptions.Compiled);
        private static readonly Regex _regex15 = new Regex("^M.*", RegexOptions.Compiled);
        private static readonly Regex _regex16 = new Regex("^G.*", RegexOptions.Compiled);
        private static readonly Regex _regex17 = new Regex(@"\w*\/\w*", RegexOptions.Compiled);
        private static readonly Regex _regex3 = new Regex(@"(?<value>[+-]?\d*[.]?\d+)\s*(?<unit>\w*)\s*?", RegexOptions.Compiled);


        private static readonly Regex _regexNumber = new Regex("^([+-]?\\d+\\.?\\d*)");

        public static bool IsNumber(this string text)
        {
            Match match = _regexNumber.Match(text);
            if (match.Success)
            {
                return true;
            }

            return false;
        }

        public static bool TryConvertToFreq(this string source, out string outputValue)
        {
            string value = _regex3.Match(source).Groups["value"].ToString();
            string unit = _regex3.Match(source).Groups["unit"].ToString();
            outputValue = string.Empty;
            if (!double.TryParse(value, out double number))
            {
                return false;
            }

            if (unit.Equals(UNIT_HZ, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = source;
            }
            else if (unit.Equals(UNIT_KHZ, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = (number * 1e3).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UNIT_MHZ, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = (number * 1e6).ToString(CultureInfo.InvariantCulture);
            }
            else if (string.IsNullOrEmpty(unit))
            {
                outputValue = value;
            }
            else
            {
                outputValue = source;
                return false;
            }

            return true;
        }

        public static bool TryConvertToVolt(this string source, out string outputValue)
        {
            string value = _regex3.Match(source).Groups["value"].ToString();
            string unit = _regex3.Match(source).Groups["unit"].ToString();
            outputValue = string.Empty;
            if (!double.TryParse(value, out double number))
            {
                return false;
            }

            if (unit.Equals(UNIT_VOLT, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = value;
            }
            else if (unit.Equals(UNIT_MINI_VOLT, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(number / 1e3, 6).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UNIT_MICRO_VOLT, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(number / 1e6, 9).ToString(CultureInfo.InvariantCulture);
            }
            else if (string.IsNullOrEmpty(unit))
            {
                outputValue = value;
            }
            else
            {
                outputValue = source;
                return false;
            }

            return true;
        }

        public static bool TryToConvertToAmpere(this string source, out string outputValue)
        {
            string value = _regex3.Match(source).Groups["value"].ToString();
            string unit = _regex3.Match(source).Groups["unit"].ToString();
            outputValue = string.Empty;
            if (!double.TryParse(value, out double number))
            {
                return false;
            }

            if (unit.Equals(UNIT_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = value;
            }
            else if (unit.Equals(UNIT_MINI_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(number / 1e3, 6).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UNIT_MICRO_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(number / 1e6, 9).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UNIT_NANO_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(number / 1e9, 9).ToString(CultureInfo.InvariantCulture);
            }
            else if (string.IsNullOrEmpty(unit))
            {
                outputValue = value;
            }
            else
            {
                outputValue = source;
                return false;
            }

            return true;
        }

        public static bool TryCombineUnit(this string source, string unit, out string outputValue)
        {
            outputValue = string.Empty;
            if (!double.TryParse(source, out double number))
            {
                return false;
            }

            if (unit.Equals(UNIT_VOLT, StringComparison.OrdinalIgnoreCase) ||
                unit.Equals(UNIT_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = source;
            }
            else if (unit.Equals(UNIT_MINI_VOLT, StringComparison.OrdinalIgnoreCase) ||
                     unit.Equals(UNIT_MINI_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(number / 1e3, 6).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UNIT_MICRO_VOLT, StringComparison.OrdinalIgnoreCase) ||
                     unit.Equals(UNIT_MICRO_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(number / 1e6, 9).ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                return false;
            }

            return true;
        }

        public static bool TryCombineVolt(this string source, string unit, out string outputValue)
        {
            outputValue = string.Empty;
            if (!double.TryParse(source, out double number))
            {
                return false;
            }
            if (unit.Equals(UNIT_VOLT, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = source;
            }
            else if (unit.Equals(UNIT_MINI_VOLT, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(number / 1e3, 6).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UNIT_MICRO_VOLT, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(number / 1e6, 9).ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                outputValue = source;
                return false;
            }

            return true;
        }

        public static bool TryCombineHz(this string source, string unit, out string outputValue)
        {
            outputValue = string.Empty;
            if (!double.TryParse(source, out double lDValue))
            {
                return false;
            }
            if (unit.Equals(UNIT_HZ, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = source;
            }
            else if (unit.Equals(UNIT_KHZ, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = (lDValue * 1e3).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UNIT_MHZ, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = (lDValue * 1e6).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UNIT_GHZ, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = (lDValue * 1e9).ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                outputValue = source;
                return false;
            }

            return true;
        }

        public static bool TryCombineAmpere(this string source, string unit, out string outputValue)
        {
            outputValue = string.Empty;
            if (!double.TryParse(source, out double lDValue))
            {
                return false;
            }
            if (unit.Equals(UNIT_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = source;
            }
            else if (unit.Equals(UNIT_MINI_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(lDValue / 1e3, 6).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UNIT_MICRO_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(lDValue / 1e6, 9).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UNIT_NANO_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(lDValue / 1e9, 12).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UNIT_PICO_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(lDValue / 1e9, 15).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UNIT_FEMTO_AMPERE, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(lDValue / 1e9, 18).ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                outputValue = source;
                return false;
            }


            return true;
        }

        public static string ConvertNumber(this string source, bool nonScience = false)
        {
            return source.ConvertNumber(out _, out _, nonScience);
        }

        public static string ConvertNumberFixedPoint(this string source)
        {
            return source.ConvertNumberFixedPoint(out _, out _);
        }

        public static string ConvertNumber(this string source, out string unit, out string scale, bool nonScience = false)
        {
            unit = "";
            scale = "";
            if (source.Contains("10^"))
            {
                source = source.Replace("*10^", "E");
            }

            bool isNumber = _regex.IsMatch(source);
            bool isExponents = _regex2.IsMatch(source);
            if (isNumber || isExponents)
            {
                string number = isNumber ? _regex.Match(source).Groups[1].Value : _regex2.Match(source).Groups[1].Value;

                unit = source.Replace(number, "").Trim().TrimStart('*');
                double rate = 1;
                if (_regex17.IsMatch(unit))
                {
                    string[] arr = unit.Split('/');
                    double rate1 = 1;
                    double rate2 = 1;
                    string scale1 = "";
                    string scale2 = "";
                    GetScale(arr[0], ref scale1, ref rate1);
                    GetScale(arr[1], ref scale2, ref rate2);
                    rate = rate1 / rate2;
                    string numerator = RemoveScale(arr[0], scale1);
                    string denominator = RemoveScale(arr[1], scale2);
                    unit = numerator + "/" + denominator;
                }
                else
                {
                    GetScale(unit, ref scale, ref rate);
                }

                unit = RemoveScale(unit, scale);
                if (double.TryParse(number, out double value))
                {
                    return nonScience ?
                        (value * rate).ToString("0." + new string('#', 339)) :
                        (value * rate).ToString("G");
                }
            }

            return source;
        }

        public static string ConvertNumberFixedPoint(this string source, out string unit, out string scale)
        {
            unit = "";
            scale = "";
            if (source.Contains("10^"))
            {
                source = source.Replace("*10^", "E");
            }

            bool isNumber = _regex.IsMatch(source);
            bool isExponents = _regex2.IsMatch(source);
            if (isNumber || isExponents)
            {
                string number = isNumber ? _regex.Match(source).Groups[1].Value : _regex2.Match(source).Groups[1].Value;

                unit = source.Replace(number, "").Trim().TrimStart('*');
                double rate = 1;
                if (_regex17.IsMatch(unit))
                {
                    string[] arr = unit.Split('/');
                    double rate1 = 1;
                    double rate2 = 1;
                    string scale1 = "";
                    string scale2 = "";
                    GetScale(arr[0], ref scale1, ref rate1);
                    GetScale(arr[1], ref scale2, ref rate2);
                    rate = rate1 / rate2;
                    string numerator = RemoveScale(arr[0], scale1);
                    string denominator = RemoveScale(arr[1], scale2);
                    unit = numerator + "/" + denominator;
                }
                else
                {
                    GetScale(unit, ref scale, ref rate);
                }
                unit = RemoveScale(unit, scale);
                if (double.TryParse(number, out double value))
                {
                    string formatted = (value * rate).ToString("F99");
                    return formatted.TrimEnd('0').TrimEnd('.');
                }
            }

            return source;
        }

        private static string RemoveScale(string unit, string scale)
        {
            unit = Regex.Replace(unit, "^" + scale, "", RegexOptions.IgnoreCase);
            unit = Regex.Replace(unit, "Ohm", "Ohm", RegexOptions.IgnoreCase);
            unit = Regex.Replace(unit, "R", "Ohm", RegexOptions.IgnoreCase);
            unit = Regex.Replace(unit, "Hz", "Hz", RegexOptions.IgnoreCase);
            return unit;
        }

        private static void GetScale(string unit, ref string scale, ref double rate)
        {
            if (_regex9.IsMatch(unit))
            {
                rate = 1 / (double)1000;
                scale = "m";
            }
            else if (_regex10.IsMatch(unit))
            {
                rate = 1 / (double)1000000;
                scale = "u";
            }
            else if (_regex11.IsMatch(unit))
            {
                rate = 1 / (double)1000000000;
                scale = "n";
            }
            else if (_regex12.IsMatch(unit))
            {
                rate = 1 / (double)1000000000000;
                scale = "p";
            }
            else if (_regex13.IsMatch(unit))
            {
                rate = 1 / (double)1000000000000000;
                scale = "f";
            }
            else if (_regex14.IsMatch(unit.ToLower()))
            {
                rate = 1000;
                scale = "K";
            }
            else if (_regex15.IsMatch(unit))
            {
                rate = 1000000;
                scale = "M";
            }
            else if (_regex16.IsMatch(unit))
            {
                rate = 1000000000;
                scale = "G";
            }
        }

        public static (string number, string unit) ConvertNumberAndUnit(this string source)
        {
            bool isNumber = _regex.IsMatch(source);
            if (!isNumber)
            {
                return ("", "");
            }
            string number = _regex.Match(source).Groups[1].Value;
            string unit = source.Replace(number, "").Trim().TrimStart('*');
            return (number, unit);
        }
    }
}
