
namespace Automation.GenerateIgxl.Basic.Business.GenAc.AcInput.BassData
{
    public class AcInputRow
    {
        public string Symbol { get; set; }
        public string Value { get; set; }
        public string Name { get; set; }
        public string Val { get; set; }
        public string Typ { get; set; }
        public string Min { get; set; }
        public string Max { get; set; }

        public AcInputRow(string symbol, string value, string name, string val, string typ, string min, string max)
        {
            Symbol = symbol;
            Value = value;
            Name = name;
            Val = val;
            Typ = typ;
            Min = min;
            Max = max;
        }
    }
}
