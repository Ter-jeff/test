namespace TestPlanLib.BinNumberLegacy
{
    public class BinNumDefPara
    {
        private string _block = "";
        private string _condition = "";
        public string Block
        {
            set { _block = value; }
            get { return _block.Replace("_", "").Replace(" ", ""); }
        }
        public string Condition
        {
            set { _condition = value; }
            get { return _condition.Replace("_", "").Replace(" ", ""); }
        }

        public BinNumDefPara(EnumBinNumDefBlock enumBinNumDefBlock, string condition)
        {
            Block = enumBinNumDefBlock.ToString();
            Condition = condition;
        }
    }
}
