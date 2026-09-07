namespace BinCutScriptLib.Base
{
    public class IdsData
    {
        //ex: IDS_VDD_CPU/IDS_VDD_FIXED/IDS_VDD_LOW
        //實際量測結果
        public string IdsName = string.Empty;
        public double RealVal;
        //無條件捨去後結果
        public double EfuseVal;
        //精度 from SetWriteDecimal of datalog 
        public double IdsLsb;
        //精度 from SetWriteDecimal of datalog 
        public EnumIdsType IdsType = EnumIdsType.None;

        public IdsData() { }

        public IdsData(IdsData idsData)
        {
            if (idsData == null)
            {
                return;
            }

            IdsName = idsData.IdsName;
            RealVal = idsData.RealVal;
            EfuseVal = idsData.EfuseVal;
            IdsLsb = idsData.IdsLsb;
            IdsType = idsData.IdsType;
        }

        public IdsData Copy()
        {
            return new IdsData(this);
        }
    }
}
