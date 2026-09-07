namespace Automation.Const
{
    public static class CommonConst
    {
        public const string ScalePico = "p";
        public const string ScaleNano = "n";
        public const string ScaleMicro = "u";
        public const string ScaleMilli = "m";
        public const string ScalePercent = "%";
        public const string ScaleKilo = "K";
        public const string ScaleMega = "M";
        public const string ScaleGiga = "G";
        public const string ScaleTera = "T";
        public const string ScaleNo = "No";

        public const string Cz2Only = "CZ2_Only";
        public const string Nv = "NV";
        public const string Lv = "LV";
        public const string Hv = "HV";

        public const string AllPower = "All_Power";

        public const string IoPin = "I/O";
        public const string InitPins = "Init_Pins";
        public const string InitHiPins = "InitHi_Pins";
        public const string InitLoPins = "InitLo_Pins";
        public const string DcvsPower = "DCVS_POWER";
        public const string DcviPower = "DCVI_POWER";

        public static string SpiRomPwr { get; set; } = "SPI_PWR";
        public const string RtosRomPwr = "RTOS_PWR";
        public const string SpiRomPins = "SPIROM_PINS";
        public const string SpiRomTimeDomain = "SPIROM_Tdomain";

        public const string NWirePosGroup = "nWireContiPos";
        public const string NWireNegGroup = "nWireContiNeg";

        public const string TimingFileCategoryMappingSheetName = "TimingFileCategoryMapping";

        public const string OutFirst = "OutFirst";
        public const string OutSecond = "OutSecond";
        public const string PinRegPattern = "(?<" + OutFirst + @">\w+)[:]{2}(?<" + OutSecond + @">\w+)";

        public const string Value = "Value";
        public const string Unit = "Unit";
        public const string UnitRegPattern = "(?<" + Value + @">\d+([.]\d+)?)\s*(?<" + Unit + @">\w+)";

        public const string Typ = "Typ";
        public const string Min = "Min";
        public const string Max = "Max";

        public const string All = "All";

        public const string MultipleConst = "Multiple_";

        public const string AutogenDefault = "Autogen Default";

        public const char Underline = '_';
        public const string Tab = "\t";
        public const string Enter = "\r\n";
    }
}
