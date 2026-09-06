using System.Text.RegularExpressions;

namespace RF_PatternTool
{
    public class LutItem
    {
        private List<LutSet> _luts;
        public string LutSetup;

        public LutItem()
        {
        }

        public LutItem(string value)
        {
            LutSetup = value.Split('=')[0].Trim();
            _luts = new List<LutSet>();
        }

        public void SetLut(string value, string fuseValue)
        {
            var lut = new LutSet(value, fuseValue);
            _luts.Add(lut);
        }

        public List<string> PrintLut()
        {
            var result = new List<string>();
            result.Add(LutSetup);
            foreach (LutSet lutSet in _luts)
            {
                var info = new List<string>();
                info.Add(lutSet.LSL.Trim());
                info.Add(lutSet.USL.Trim());
                info.Add(lutSet.FuseValue);
                result.Add(string.Join("\t", info));
            }
            return result;
        }

        private class LutSet
        {
            public string LSL { get; } = "NA";

            public string USL { get; } = "NA";

            public string FuseValue { get; } = "NA";

            public LutSet()
            {
            }

            public LutSet(string value, string fuse)
            {
                /*
                 * LUT,V<0.4,tx_5g_ppa_casc_bias_tune_1p0 =,0xD,,,,,
LUT,V>0.46,tx_5g_ppa_casc_bias_tune_1p0 =,0xE,,,,,
LUT,0.4<V<0.46,tx_5g_ppa_casc_bias_tune_1p0 =,0xF,,,,,
                 */

                FuseValue = fuse;
                string valueStr;
                for (int i = 0; i < value.Split('>').Count() - 1; i++)
                {
                    valueStr = ConvertValueUnit(value.Split('>')[i]).Trim();
                    if (!string.IsNullOrEmpty(valueStr))
                    {
                        USL = valueStr;
                    }

                    valueStr = ConvertValueUnit(value.Split('>')[i + 1]).Trim();
                    if (!string.IsNullOrEmpty(valueStr))
                    {
                        LSL = valueStr;
                    }
                }

                for (int i = 0; i < value.Split('<').Count() - 1; i++)
                {
                    valueStr = ConvertValueUnit(value.Split('<')[i]).Trim();
                    if (!string.IsNullOrEmpty(valueStr))
                    {
                        LSL = valueStr;
                    }

                    valueStr = ConvertValueUnit(value.Split('<')[i + 1]).Trim();
                    if (!string.IsNullOrEmpty(valueStr))
                    {
                        USL = valueStr;
                    }
                }
            }

            private static string ConvertValueUnit(string rawValue)
            {
                string regValue = @"^(?<value>-*\d+(\.\d+)*)";
                string regUnit = @"(?<unit>[umkMG])";
                string regSci = @"(?<unit>e-*\d+)";
                string result = Regex.Match(rawValue, regValue).Groups["value"].ToString();
                if (string.IsNullOrEmpty(result))
                {
                    return result;
                }

                string unit = Regex.Match(rawValue, regUnit, RegexOptions.IgnoreCase).Groups["unit"].ToString();
                string unit_after = "";
                if (!string.IsNullOrEmpty(unit))
                {

                    switch (unit)
                    {
                        case "u":
                            unit_after = "E-6";
                            break;
                        case "m":
                            unit_after = "E-3";
                            break;
                        case "k":
                            unit_after = "E3";
                            break;
                        case "M":
                            unit_after = "E6";
                            break;
                        case "G":
                            unit_after = "E9";
                            break;
                        default:
                            break;
                    }

                }
                else if (Regex.IsMatch(rawValue, regSci, RegexOptions.IgnoreCase))
                {
                    unit_after = Regex.Match(rawValue, regSci, RegexOptions.IgnoreCase).Groups["unit"].ToString().ToUpper();
                    ;
                }
                if (Regex.IsMatch(unit_after, regSci, RegexOptions.IgnoreCase))
                {
                    string regPower = Regex.Match(unit_after, @"e(?<Power>-*\d+)", RegexOptions.IgnoreCase).Groups["Power"].Value;
                    double calcValue = double.Parse(result) * Math.Pow(10, int.Parse(regPower));
                    return Convert.ToString(calcValue);
                }
                return result + unit_after;
            }
        }
    }
}
