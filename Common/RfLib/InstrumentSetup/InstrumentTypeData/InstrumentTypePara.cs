using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace RfLib.InstrumentSetup.InstrumentTypeData
{
    [ExcludeFromCodeCoverage]
    public class InstrumentTypePara
    {
        public string InstrumentType { get; set; } = string.Empty;
        public Dictionary<string, int> DicPara = [];
    }
}
