using System.Collections.Generic;
using System.Globalization;

using Automation.GenerateIgxl.Basic.Business.GenAc.AcInput.BassData;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

namespace Automation.GenerateIgxl.Basic.Business.GenAc
{
    public class BasicInitial
    {
        public const string Common = "Common";
        public const string TckFreqVar = "TCK_Freq";
        public const string ShiftInFreqVar = "ShiftIn_Freq";
        public const string T0 = "T0";
        public const string ShmooFreqVar = "_Shmoo_Freq";
        public const string CycleS = "Cycle_S";
        public const string ClockS = "Clock_S";
        public const string ClockE = "Clock_E";
        public const string Strobe = "Strobe";

        public const string Typ = "Typ";
        public const string Min = "Min";
        public const string Max = "Max";

        public const string AcSpecDefault = "-1";

        public AcInputSheet InitalAcSymbols()
        {
            string freq24Mhz = 24e6.ToString();
            var acInput = new AcInputSheet();

            //for Time Set use
            acInput.AddRow(AddAcSymbolWithSameValue(TckFreqVar, freq24Mhz));         //freq24Mhz

            //CharSetUp2D, SCAN Time Set
            acInput.AddRow(AddAcSymbolWithSameValue(ShiftInFreqVar, AcSpecDefault)); // freq20Mhz
            //Add nWire Freq Var & shmoo for 2DCharSetUp use
            List<ProtocolAwarePin> nwirePins = NwireSingleton.Instance().SettingInfo.NwirePins;
            bool isUfOnly = !TestPlanStatic.Equipments.Exists(x => x.Equals(EnumEquipment.UltraFlexPlus));
            foreach (ProtocolAwarePin pin in nwirePins)
            {
                acInput.AddRow(AddAcSymbolWithSameValue(pin.CreateAcSymbol(), "=_" + pin.CreateFreqSpecName(isUfOnly ? EnumEquipment.UltraFlex : EnumEquipment.UltraFlexPlus)));
                acInput.AddRow(AddAcSymbolWithSameValue(pin.OutClk + ShmooFreqVar, pin.Freq.ToString(CultureInfo.InvariantCulture)));
            }
            return acInput;
        }

        private AcInputRow AddAcSymbolWithSameValue(string symbol, string value)
        {
            var row = new AcInputRow(symbol, value, Typ, Typ, value, value, value);
            return row;
        }
    }
}
