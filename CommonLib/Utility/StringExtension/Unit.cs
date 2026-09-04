using System;
using System.Text.RegularExpressions;

namespace CommonLib.Utility.StringExtension
{
    public static class Unit
    {
        private const string ScalePercent = "%";
        private const string ScaleNano = "n"; //10e-9
        private const string ScaleMicro = "u"; //10e-6
        private const string ScaleMilli = "m"; //10e-3
        private const string ScaleKilo = "K"; //10e+3
        private const string ScaleMega = "M"; //10e+6
        private const string ScaleGiga = "G"; //10e+9
        private static readonly Regex _regex = new Regex(@"^[+|-]?s*\d+(\.\d+)?$", RegexOptions.Compiled);
        private static readonly Regex _regex2 = new Regex(@"^(?<num>[+|-]?s*\d+(\.\d+)?)", RegexOptions.Compiled);
        private static readonly Regex _regex3 = new Regex(@"^(?<notion>E[+|-]?\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex4 = new Regex(@"^\*", RegexOptions.Compiled);

        public static string ConvertUnit(this string text)
        {
            return text.ConvertUnit(out _, out _);
        }

        public static string ConvertUnit(this string text, out string unit)
        {
            return text.ConvertUnit(out unit, out _);
        }

        public static string ConvertUnit(this string text, out string unit, out string scale)
        {
            unit = "";
            scale = "";
            text = text.Replace(" ", "");
            if (text.Contains("10^"))
            {
                text = text.Replace("*10^", "E");
            }

            if (text == "")
            {
                return text;
            }

            if (_regex.IsMatch(text)) //number only
            {
                return text;
            }

            bool isScientific = false;
            if (_regex2.IsMatch(text))
            {
                string number = _regex2.Match(text).Groups["num"].ToString();
                unit = text.Replace(number, "");
                if (_regex3.IsMatch(unit))
                {
                    string scientific = _regex3.Match(unit).Groups["notion"].ToString();
                    number = number + scientific;
                    unit = unit.Replace(scientific, "");
                    isScientific = true;
                }

                unit = _regex4.Replace(unit, "");
                double rate = 1;
                if (unit.StartsWith(ScalePercent))
                {
                    rate = 1 / (double)100;
                    unit = unit.Substring(1, unit.Length - 1);
                    scale = ScalePercent;
                }
                else if (unit.StartsWith(ScaleMilli))
                {
                    rate = 1 / (double)1000;
                    unit = unit.Substring(1, unit.Length - 1);
                    scale = ScaleMilli;
                }
                else if (unit.StartsWith(ScaleMicro))
                {
                    rate = 1 / (double)1000000;
                    unit = unit.Substring(1, unit.Length - 1);
                    scale = ScaleMicro;
                }
                else if (unit.StartsWith(ScaleNano))
                {
                    rate = 1 / (double)1000000000;
                    unit = unit.Substring(1, unit.Length - 1);
                    scale = ScaleNano;
                }
                else if (unit.StartsWith(ScaleKilo, StringComparison.CurrentCultureIgnoreCase))
                {
                    rate = 1000;
                    unit = unit.Substring(1, unit.Length - 1);
                    scale = ScaleKilo;
                }
                else if (unit.StartsWith(ScaleMega))
                {
                    rate = 1000000;
                    unit = unit.Substring(1, unit.Length - 1);
                    scale = ScaleMega;
                }
                else if (unit.StartsWith(ScaleGiga))
                {
                    rate = 1000000000;
                    unit = unit.Substring(1, unit.Length - 1);
                    scale = ScaleGiga;
                }

                if (unit.Equals("HZ", StringComparison.CurrentCultureIgnoreCase))
                {
                    unit = "Hz";
                }
                else if (unit.Equals("OHM", StringComparison.CurrentCultureIgnoreCase))
                {
                    unit = "Ohm";
                }
                else if (unit.Equals("OHMS", StringComparison.CurrentCultureIgnoreCase))
                {
                    unit = "Ohms";
                }

                double value;
                if (double.TryParse(number, out value))
                {
                    if (isScientific)
                    {
                        return (value * rate).ToString("G2");
                    }

                    return (value * rate).ToString("G");
                }
            }
            return text;
        }

        public static double ToDouble(this string text)
        {
            double value;
            double.TryParse(ConvertUnit(text, out _, out _), out value);
            return value;
        }
    }
}
