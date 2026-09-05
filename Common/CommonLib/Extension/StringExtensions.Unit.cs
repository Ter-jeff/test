using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace CommonLib.Extension
{
    public static partial class StringExtensions
    {
        private const string ScalePercent = "%";
        private const string UnitAmpere = "A";
        private const string UnitMilliAmpere = "mA";
        private const string UnitMicroAmpere = "uA";
        private const string UnitNanoAmpere = "nA";
        private const string UnitPicoAmpere = "pA";
        private const string UnitFemtoAmpere = "fA";

        private const string UnitVolt = "V";
        private const string UnitMilliVolt = "mV";
        private const string UnitMicroVolt = "uV";

        private const string UnitHz = "Hz";
        private const string UnitKhz = "KHz";
        private const string UnitMhz = "MHz";
        private const string UnitGhz = "GHz";

        [GeneratedRegex(@"^([+-]?\d+\.?\d*)\*?([a-zA-Z]|%|\/)*$", RegexOptions.Compiled)]
        private static partial Regex MyRegexPlainNumber();
        [GeneratedRegex(@"^([+-]?\d+\.?\d*[Ee][+-]?\d+)\*?([a-zA-Z]|%|\/)*$", RegexOptions.Compiled)]
        private static partial Regex MyRegexScientific();
        [GeneratedRegex("^m.*", RegexOptions.Compiled)]
        private static partial Regex MyRegexMilliPrefix();
        [GeneratedRegex("^u.*", RegexOptions.Compiled)]
        private static partial Regex MyRegexMicroPrefix();
        [GeneratedRegex("^n.*", RegexOptions.Compiled)]
        private static partial Regex MyRegexNanoPrefix();
        [GeneratedRegex("^p.*", RegexOptions.Compiled)]
        private static partial Regex MyRegexPicoPrefix();
        [GeneratedRegex("^f.*", RegexOptions.Compiled)]
        private static partial Regex MyRegexFemtoPrefix();
        [GeneratedRegex("^k.*", RegexOptions.Compiled)]
        private static partial Regex MyRegexKiloPrefix();
        [GeneratedRegex("^M.*", RegexOptions.Compiled)]
        private static partial Regex MyRegexMegaPrefix();
        [GeneratedRegex("^G.*", RegexOptions.Compiled)]
        private static partial Regex MyRegexGigaPrefix();
        [GeneratedRegex(@"\w*\/\w*", RegexOptions.Compiled)]
        private static partial Regex MyRegexHasSlash();
        [GeneratedRegex(@"^(?<value>[+-]?\d*[.]?\d+)\s*(?<unit>\w*)\s*$", RegexOptions.Compiled)]
        private static partial Regex MyRegexValueUnit();
        [GeneratedRegex("^([+-]?\\d+\\.?\\d*)")]
        private static partial Regex MyRegexLeadingNumber();
        [GeneratedRegex("Ohm", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegexOhmWord();
        [GeneratedRegex("R", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegexOhmSymbol();
        [GeneratedRegex("Hz", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegexHzWord();

        private static readonly Regex _regexPlainNumber = MyRegexPlainNumber();
        private static readonly Regex _regexScientific = MyRegexScientific();
        private static readonly Regex _regexMilliPrefix = MyRegexMilliPrefix();
        private static readonly Regex _regexMicroPrefix = MyRegexMicroPrefix();
        private static readonly Regex _regexNanoPrefix = MyRegexNanoPrefix();
        private static readonly Regex _regexPicoPrefix = MyRegexPicoPrefix();
        private static readonly Regex _regexFemtoPrefix = MyRegexFemtoPrefix();
        private static readonly Regex _regexKiloPrefix = MyRegexKiloPrefix();
        private static readonly Regex _regexMegaPrefix = MyRegexMegaPrefix();
        private static readonly Regex _regexGigaPrefix = MyRegexGigaPrefix();
        private static readonly Regex _regexHasSlash = MyRegexHasSlash();
        private static readonly Regex _regexValueUnit = MyRegexValueUnit();
        private static readonly Regex _regexLeadingNumber = MyRegexLeadingNumber();

        public static bool IsNumber(this string text)
        {
            Match match = _regexLeadingNumber.Match(text);
            if (match.Success)
            {
                return true;
            }

            return false;
        }

        private static (bool Success, double Number, string Unit) ParseSource(string source)
        {
            Match match = _regexValueUnit.Match(source);
            if (!match.Success)
            {
                return (false, 0, string.Empty);
            }

            string valueStr = match.Groups["value"].Value;
            string unitStr = match.Groups["unit"].Value;

            if (!double.TryParse(valueStr, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
            {
                return (false, 0, string.Empty);
            }

            return (true, number, unitStr);
        }

        private static bool TryConvertUnit(string source, string baseUnit, Dictionary<string, double> unitMultipliers, out string outputValue)
        {
            (bool success, double number, string unit) = ParseSource(source);

            if (!success)
            {
                outputValue = string.Empty;
                return false;
            }

            if (string.IsNullOrEmpty(unit) || unit.EqualsIgnoreCase(baseUnit))
            {
                outputValue = number.ToString("R", CultureInfo.InvariantCulture);
                return true;
            }

            if (unitMultipliers.TryGetValue(unit, out double multiplier))
            {
                outputValue = (number * multiplier).ToString("R", CultureInfo.InvariantCulture);
                return true;
            }

            outputValue = string.Empty;
            return false;
        }

        public static bool TryConvertToFreq(this string source, out string outputValue)
        {
            var mappings = new Dictionary<string, double>(IgnoreCase)
            {
                { UnitKhz, 1e3 },
                { UnitMhz, 1e6 }
            };
            return TryConvertUnit(source, UnitHz, mappings, out outputValue);
        }

        public static bool TryConvertToVolt(this string source, out string outputValue)
        {
            var mappings = new Dictionary<string, double>(IgnoreCase)
        {
            { UnitMilliVolt, 1e-3 },
            { UnitMicroVolt, 1e-6 }
        };
            return TryConvertUnit(source, UnitVolt, mappings, out outputValue);
        }

        public static bool TryConvertToAmpere(this string source, out string outputValue)
        {
            var mappings = new Dictionary<string, double>(IgnoreCase)
            {
                { UnitMilliAmpere, 1e-3 },
                { UnitMicroAmpere, 1e-6 },
                { UnitNanoAmpere, 1e-9 }
            };
            return TryConvertUnit(source, UnitAmpere, mappings, out outputValue);
        }

        public static bool TryCombineVolt(this string source, string unit, out string outputValue)
        {
            outputValue = string.Empty;
            if (!double.TryParse(source, out double number))
            {
                return false;
            }
            if (unit.EqualsIgnoreCase(UnitVolt))
            {
                outputValue = source;
            }
            else if (unit.EqualsIgnoreCase(UnitMilliVolt))
            {
                outputValue = Math.Round(number / 1e3, 6).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.EqualsIgnoreCase(UnitMicroVolt))
            {
                outputValue = Math.Round(number / 1e6, 9).ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                return false;
            }

            return true;
        }

        public static bool TryCombineHz(this string source, string unit, out string outputValue)
        {
            outputValue = string.Empty;
            if (!double.TryParse(source, out double number))
            {
                return false;
            }
            if (unit.EqualsIgnoreCase(UnitHz))
            {
                outputValue = source;
            }
            else if (unit.EqualsIgnoreCase(UnitKhz))
            {
                outputValue = (number * 1e3).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.EqualsIgnoreCase(UnitMhz))
            {
                outputValue = (number * 1e6).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.EqualsIgnoreCase(UnitGhz))
            {
                outputValue = (number * 1e9).ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                return false;
            }

            return true;
        }

        public static bool TryCombineAmpere(this string source, string unit, out string outputValue)
        {
            outputValue = string.Empty;
            if (!double.TryParse(source, out double number))
            {
                return false;
            }
            if (unit.EqualsIgnoreCase(UnitAmpere))
            {
                outputValue = source;
            }
            else if (unit.EqualsIgnoreCase(UnitMilliAmpere))
            {
                outputValue = Math.Round(number / 1e3, 6).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.EqualsIgnoreCase(UnitMicroAmpere))
            {
                outputValue = Math.Round(number / 1e6, 9).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.EqualsIgnoreCase(UnitNanoAmpere))
            {
                outputValue = Math.Round(number / 1e9, 12).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.EqualsIgnoreCase(UnitPicoAmpere))
            {
                outputValue = Math.Round(number / 1e12, 15).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.EqualsIgnoreCase(UnitFemtoAmpere))
            {
                outputValue = Math.Round(number / 1e15, 15).ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                return false;
            }

            return true;
        }

        public static string ConvertNumber(this string source, bool nonScience = false)
        {
            return source.ConvertNumber(out _, out _, nonScience);
        }

        public static string ConvertNumber(this string source, out string unit, out string scale, bool nonScience = false)
        {
            unit = "";
            scale = "";
            if (source.Contains("10^"))
            {
                source = source.Replace("*10^", "E");
            }

            bool isNumber = _regexPlainNumber.IsMatch(source);
            bool isExponents = _regexScientific.IsMatch(source);
            if (isNumber || isExponents)
            {
                string number = isNumber ? _regexPlainNumber.Match(source).Groups[1].Value : _regexScientific.Match(source).Groups[1].Value;

                unit = source.Replace(number, "").Trim().TrimStart('*');
                double rate = 1;
                if (_regexHasSlash.IsMatch(unit))
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

                    if (nonScience)
                    {
                        return (value * rate).ToString("0." + new string('#', 339));
                    }
                    if (isExponents)
                    {
                        return (value * rate).ToString("G2");
                    }
                    return (value * rate).ToString("G15");
                }
            }

            return source;
        }

        private static string RemoveScale(string unit, string scale)
        {
            unit = Regex.Replace(unit, "^" + scale, "", RegexOptions.IgnoreCase);
            unit = MyRegexOhmWord().Replace(unit, "Ohm");
            unit = MyRegexOhmSymbol().Replace(unit, "Ohm");
            unit = MyRegexHzWord().Replace(unit, "Hz");
            return unit;
        }

        private static void GetScale(string unit, ref string scale, ref double rate)
        {
            if (_regexMilliPrefix.IsMatch(unit))
            {
                rate = 1 / (double)1000;
                scale = "m";
            }
            else if (_regexMicroPrefix.IsMatch(unit))
            {
                rate = 1 / (double)1000000;
                scale = "u";
            }
            else if (_regexNanoPrefix.IsMatch(unit))
            {
                rate = 1 / (double)1000000000;
                scale = "n";
            }
            else if (_regexPicoPrefix.IsMatch(unit))
            {
                rate = 1 / (double)1000000000000;
                scale = "p";
            }
            else if (_regexFemtoPrefix.IsMatch(unit))
            {
                rate = 1 / (double)1000000000000000;
                scale = "f";
            }
            else if (_regexKiloPrefix.IsMatch(unit.ToLower()))
            {
                rate = 1000;
                scale = "K";
            }
            else if (_regexMegaPrefix.IsMatch(unit))
            {
                rate = 1000000;
                scale = "M";
            }
            else if (_regexGigaPrefix.IsMatch(unit))
            {
                rate = 1000000000;
                scale = "G";
            }
            else if (unit.StartsWith(ScalePercent))
            {
                rate = 1 / (double)100;
                scale = ScalePercent;
            }
        }

        public static (string number, string unit) ConvertNumberAndUnit(this string source)
        {
            bool isNumber = _regexPlainNumber.IsMatch(source);
            if (!isNumber)
            {
                return ("", "");
            }
            string number = _regexPlainNumber.Match(source).Groups[1].Value;
            string unit = source.Replace(number, "").Trim().TrimStart('*');
            return (number, unit);
        }
    }
}
