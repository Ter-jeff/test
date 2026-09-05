using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.GenerateIgxl.BinCut.Base
{
    public class BinCutBinningItem
    {
        public string BinName { set; get; }
        public string Block { set; get; }
        public string PerformanceMode { set; get; }
        public string FlagName { set; get; }
        public string Level { get; }

        public BinCutBinningItem(string binningName, string block, string performanceMode, string flag, EnumColumnName columnName)
        {
            BinName = binningName;
            Block = block;
            PerformanceMode = performanceMode;
            FlagName = flag;
            //HBV only
            if (columnName == EnumColumnName.TD)
            {
                Level = "TD HBV";
            }
            else if (columnName == EnumColumnName.Mbist)
            {
                Level = "Mbist HBV";
            }
            else if (columnName == EnumColumnName.ILB)
            {
                Level = "ILB HBV";
            }
            else if (columnName == EnumColumnName.ELB)
            {
                Level = "ELB HBV";
            }
            else
            {
                Level = "SPI HBV";
            }
        }
    }
}
