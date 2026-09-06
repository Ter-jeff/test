using System.Text.RegularExpressions;

using CommonReaderLib;

using EfuseCheckCmdLib.Base;

using TestPlanLib.Efuse;

namespace EfuseCheckCmdLib.BDF
{
    public partial class BitDefRow : MyRow
    {
        public const string Regbin = "(?<format>^0*[bB])[01]+";
        public const string Reghex = "(?<format>^0*[xX])[a-fA-F0-9]+";
        public const string RegDec = @"(?<value>\d+(\.\d+)*)(?<unit>[numkKMgG]*)\w*";

        [GeneratedRegex(Reghex)]
        private static partial Regex HexRegex();

        [GeneratedRegex(Regbin)]
        private static partial Regex BinRegex();

        //from output
        public string RealStage = "";
        public int JobStage;
        public string Name = "";
        public string Resolution = "";
        public bool IsDefault;
        public string DefaultReal = "";
        public string Msb = "";
        public string Lsb = "";
        public string LowLimit = "";
        public string HighLimit = "";
        public double RealValue;
        public string DefaultValue = "";
        public string Block = "";
        public string Algorithm = "";
        public string HiPparName = "";
        public string HiPparEq = "";
        public double HipHighLimit;
        public double HipLowLimit;
        public ValueType Type;

        public bool IsCrc { get; internal set; }
        public CrcItem Crc { get; internal set; } = null!;

        public ValueType JudgeValueType(string value)
        {
            if (IsDefault)
            {
                ValueType type = ValueType.Dec;
                if (HexRegex().IsMatch(value))
                {
                    return ValueType.Hex;
                }

                if (BinRegex().IsMatch(value))
                {
                    return ValueType.Bin;
                }
                return type;
            }
            return ValueType.Real;
        }
    }
}
