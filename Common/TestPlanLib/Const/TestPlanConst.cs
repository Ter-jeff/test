
namespace TestPlanLib.Const
{
    public static class TestPlanConst
    {
        //nWire
        public const string OutClkPinPattern = @"^Output\s*Clock\s*Pin$";
        public const string OutClkPinVoltagePattern = @"^Output\s*Clock\s*Voltage$";
        public const string FrequencyPattern = @"^Default\s*Frequency$";
        public const string PowerUpSequencePattern = @"^Power\s*Up\s*Sequence$";
        public const string PowerDownSequencePattern = @"^Power\s*Down\s*Sequence$";
        public const string RelayControlPattern = @"^Relay\s*Control$";
        public const string RefClkPinPattern = @"^Reference\s*Clock\s*Pin$";
        public const string RefClkVoltagePattern = @"^Reference\s*Clock\s*Voltage$";
        public const string Protocol = "Protocol";
        public const string DcCategory = @"^DC\s*Category$";

        public const string OutFirst = "OutFirst";
        public const string OutSecond = "OutSecond";
        public const string PinRegPattern = "(?<" + OutFirst + @">\w+)[:]{2}(?<" + OutSecond + @">\w+)";

        public const string Value = "Value";
        public const string Unit = "Unit";
        public const string UnitRegPattern = "(?<" + Value + @">\d+([.]\d+)?)\s*(?<" + Unit + @">\w+)";

        public const string ValtRowPinNameFlag = "_Valt";
    }
}
