using System;
using System.Linq;

using BinCutScriptLib.Static;

namespace BinCutScriptLib.Base.Line
{
    public class OffsetLine1 : PatternLine
    {
        public OffsetData GetOffsetData()
        {
            if (!Line.StartsWith("=>"))
            {
                string[] arr = Line.Split(',');
                var offsetData = new OffsetData();

                //Dynamic Offset,0,VDD_SOC,Voltage Change From 781.25mV To 762.5mV (Offset=-18.75mV)
                if (Line.StartsWith("Dynamic Offset"))
                {
                    offsetData.PinName = arr[2];
                    _ = int.TryParse(arr[1], out offsetData.Site);
                    string[] arrtmp = arr[3].Split(' ');
                    _ = double.TryParse(arrtmp[6].Replace("(Offset=", "").Replace("mV)", ""), out offsetData.OffSetValue);
                    _ = double.TryParse(arrtmp[5].Replace("mV", ""), out offsetData.VoltageValue);
                }
                else
                {
                    string[] patToken = arr[1].Split([' '], StringSplitOptions.RemoveEmptyEntries);
                    if (arr.Length <= 2)
                    {
                        return null!;
                    }

                    string offset = Reg.Regexoffset.Match(arr[1]).Groups["offset"].Value;
                    double value = 0;
                    if (int.TryParse(arr.First().Replace("Site: ", ""), out int site))
                    {
                    }
                    if (double.TryParse(offset, out double offsetValue))
                    {
                    }
                    if (arr.Length >= 3)
                    {
                        if (double.TryParse(arr[2].Split('=')[1], out value))
                        {
                        }
                    }

                    string powerName = patToken[0].Replace("_OFFSET", "");
                    offsetData.PinName = powerName;
                    offsetData.Site = site;
                    offsetData.OffSetValue = offsetValue;
                    offsetData.VoltageValue = value;
                }

                return offsetData;
            }
            else
            {
                string[] patToken = Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                string offset = Reg.RegexOldRegexoffset.Match(Line).Groups["offset"].Value;
                if (double.TryParse(offset, out double offsetValue))
                {
                }
                offsetValue *= -1000;
                string powerName = patToken[1];

                var offsetData = new OffsetData
                {
                    PinName = powerName,
                    OffSetValue = offsetValue
                };
                return offsetData;
            }
        }
    }

    public class OffsetData
    {
        public string PinName = string.Empty;
        public double OffSetValue;
        public double VoltageValue;
        public int Site;
    }
}
