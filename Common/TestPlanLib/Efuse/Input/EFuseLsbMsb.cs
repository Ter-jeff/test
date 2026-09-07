namespace TestPlanLib.Efuse.Input
{
    public class EFuseLsbMsb
    {
        private const char SplitColon = ':';
        private string _lsb = "";
        private string _msb = "";
        private bool _sortFlg = true;

        #region Enter Method
        public bool SetLsbmsbData(string pStrLsbmsb)
        {
            pStrLsbmsb = pStrLsbmsb.Replace("[", "");
            pStrLsbmsb = pStrLsbmsb.Replace("]", "");

            if (pStrLsbmsb.Trim().StartsWith(':'))
            {
                pStrLsbmsb = pStrLsbmsb[1..];
            }
            string[] lStrLsbmsbList = pStrLsbmsb.Split(SplitColon);
            if (lStrLsbmsbList.Length == 2)
            {
                if (lStrLsbmsbList[0] == "N/A" || lStrLsbmsbList[1] == "N/A")
                {
                    _lsb = "0";
                    _msb = "0";
                    return false;
                }
                if (int.Parse(lStrLsbmsbList[0]) >= int.Parse(lStrLsbmsbList[1]))
                {
                    _lsb = lStrLsbmsbList[1].Trim();
                    _msb = lStrLsbmsbList[0].Trim();
                    _sortFlg = false;
                }
                else
                {
                    _lsb = lStrLsbmsbList[0].Trim();
                    _msb = lStrLsbmsbList[1].Trim();
                    _sortFlg = true;
                }
                return true;
            }
            else
            {
                _lsb = pStrLsbmsb.Trim();
                _msb = pStrLsbmsb.Trim();
                return false;
            }
        }
        #endregion

        #region public property
        public string GetLsb()
        {
            return _lsb;
        }

        public string GetMsb()
        {
            return _msb;
        }

        public bool GetSortFlg()
        {
            return _sortFlg;
        }
        #endregion
    }
}
