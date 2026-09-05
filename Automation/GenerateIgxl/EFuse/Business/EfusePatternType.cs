namespace Automation.GenerateIgxl.EFuse.Business
{
    public class EfusePatternType
    {
        public string PatJob = "";
        public bool IsDvrv;
        public bool IsEarly;
        public bool IsFinal = false;
        public DvRvType DvrvType = DvRvType.Not;
        public EfuseTestMode TestMode = EfuseTestMode.Unknow;
    }
}
