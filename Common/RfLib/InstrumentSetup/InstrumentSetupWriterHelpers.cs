using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using RfLib.InstrumentSetup.InstrumentTypeData;

namespace RfLib.InstrumentSetup
{
    internal static partial class InstrumentSetupWriterHelpers
    {
        private const string LXPlusSource = "LXPlus Source";
        private const string LXPlusAnalyzer = "LXPlus Analyzer";
        private const string LitePointSource = "LitePoint Source";
        private const string LitePointAnalyzer = "LitePoint Analyzer";
        private const string UltraWave24CWSource = "UltraWave24 CW Source";
        private const string UltraWave24MultiToneSource = "UltraWave24 MultiTone Source";
        private const string UltraWave24ModulatedSource = "UltraWave24 Modulated Source";
        private const string UltraWave24Receiver = "UltraWave24 Receiver";
        private const string UltraPac80Source = "UltraPac80 Source";
        private const string UltraPac80Capture = "UltraPac80 Capture";
        private const string Gigadig = "Gigadig";
        private const string UVI80Vdm = "UVI80 Vdm";
        private const string UVI80Source = "UVI80 Source";
        private const string UVI80Meter = "UVI80 Meter";
        private const string UVS256Meter = "UVS256 Meter";
        private const string UP1600 = "UP1600";
        private const string UP1600SingleMode = "UP1600 (Single mode)";
        private const string UP1600DualMode = "UP1600 (Dual mode)";
        private const string UP1600QuadMode = "UP1600 (Quad mode)";
        private const string MW7GSource = "MW7G Source";
        private const string MW7GAnalyzer = "MW7G Analyzer";
        // new setup for LookBack #33,#34
        private const string SwitchControl = "SWITCH CONTROL";

        [GeneratedRegex("::")]
        private static partial Regex MyRegex();

        public static void ClassClone(List<InstrumentSetupRow> instrumentSetupRows, InstrumentSetupRow instrumentSetupRow)
        {
            var copyRow = (InstrumentSetupRow)instrumentSetupRow.Clone();
            instrumentSetupRows.Add(copyRow);
        }

        internal static string GetInstrSuffix(string type)
        {
            return type switch
            {
                UltraWave24CWSource or UltraWave24MultiToneSource or UltraWave24ModulatedSource or UltraWave24Receiver => "_UW",
                //for issue #53
                UP1600 => "_UP1600",
                UVI80Meter => "_UVI80",// for issue #53
                                       //case UltraPac80Source:
                                       //    return "_SRC";
                UltraPac80Capture => "_CAP",
                _ => "",
            };
        }

        internal static string GetSubSettingName(string type)
        {
            return type switch
            {
                LXPlusSource => "LXS",
                LXPlusAnalyzer => "LXM",
                LitePointSource => "LPS",
                LitePointAnalyzer => "LPA",
                UltraWave24CWSource => "UWS",
                UltraWave24MultiToneSource => "UWS",
                UltraWave24ModulatedSource => "UWS",
                UltraWave24Receiver => "UWM",
                UltraPac80Source => "UPS",
                UltraPac80Capture => "UPC",
                Gigadig => "Giga",
                UVI80Vdm => "UVI",
                UVI80Meter => "UVI",
                UVS256Meter => "UVS",
                UVI80Source => "UVI",
                UP1600 => "UP",
                UP1600SingleMode => "UP",
                UP1600DualMode => "UP",
                UP1600QuadMode => "UP",
                MW7GSource => "MW7S",
                MW7GAnalyzer => "MW7A",
                SwitchControl => "SwitchControl",// new setup for LookBack #33,#34
                _ => "",
            };
        }

        internal static string ProcessPinName(string pin, string extra, string pinsuffix)
        {
            //var diffPins = Regex.Split(pin, "::").Select(p => p + extra + pinsuffix);
            IEnumerable<string> diffPins = MyRegex().Split(pin).Select(p => p.EndsWithIgnoreCase($"{pinsuffix}") ?
            p
            : p + (extra ?? string.Empty) + (pinsuffix ?? string.Empty));
            return string.Join("::", diffPins);
        }

        internal static void WriteModeValue(InstrumentSetupRow instrumentSetupRow, int rowNumber, InstrumentTypePara instrumentTypePara)
        {
            if (instrumentTypePara.DicPara.ContainsKey("mode"))
            {
                if (instrumentSetupRow.MeasSeqType.EqualsIgnoreCase("I"))
                {
                    instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][instrumentTypePara.DicPara["mode"] - 1] = "Current";
                }

                if (instrumentSetupRow.MeasSeqType.EqualsIgnoreCase("V"))
                {
                    instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][instrumentTypePara.DicPara["mode"] - 1] = "Voltage";
                }
            }
        }

        internal static void WriteParaValues(InstrumentSetupRow instrumentSetupRow, int rowNumber, Dictionary<string, string> dicInfoPara, InstrumentTypePara instrumentTypePara)
        {
            foreach (KeyValuePair<string, string> para in dicInfoPara)
            {
                if (instrumentTypePara.DicPara.TryGetValue(para.Key, out int index))
                {
                    if (instrumentTypePara.DicPara.ContainsKey(para.Key))
                    {
                        instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][index - 1] = para.Value;
                    }
                }
                else
                {
                    //hardcode due to key is not match to the instrument setup
                    if (para.Key.EqualsIgnoreCase("srcpin") && instrumentTypePara.DicPara.TryGetValue("reserved1", out int index1))
                    {
                        instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][index1 - 1] = para.Key + "=" + para.Value;
                    }
                    else if (para.Key.EqualsIgnoreCase("att") && instrumentTypePara.DicPara.TryGetValue("reserved2", out int index2))
                    {
                        instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][index2 - 1] = para.Key + "=" + para.Value;
                    }
                }
            }
        }

        internal static void WritePathCtrValue(InstrumentSetupRow instrumentSetupRow, int rowNumber, Dictionary<string, string> dicInfoPara, InstrumentTypePara instrumentTypePara)
        {
            if (instrumentTypePara.DicPara.TryGetValue("pathctr", out int value))
            {
                instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][value - 1] = instrumentSetupRow.Pins;
                instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][instrumentTypePara.DicPara["expectatt"] - 1] = dicInfoPara["expected_attenuation"];
            }
        }
    }
}
