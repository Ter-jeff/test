namespace Automation.GenerateIgxl.PostAction.GenTestNumber.Base
{
    internal class TestNumberBase
    {
        #region Field
        private const long Max = 999999999;
        public long StartNum { set; get; } = 0;
        public long Interval { set; get; } = 100;
        public long MaxNum { set; get; } = Max;
        public int CurrCnt { set; get; } = 0;

        #endregion
    }
}
