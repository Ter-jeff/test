namespace TestPlanLib.DataStruct
{
    public enum DcCategoryType
    {
        Default,
        Ids,
        Evs,
        MbistEfuse,
        PerformanceMode,
        Vmargin,
        Conti,
        HardIp,
        HardIpDc,
        Lcd,
        HardIpPerfMode,
        Rtos,
        MemoryBist
    }

    public class DcCategory(string categoryName, string block, string subCat, DcCategoryType dcCategoryType, string targetDcSpec = "")
    {
        public DcCategoryType Type { set; get; } = dcCategoryType;
        public string CategoryName { set; get; } = categoryName;
        public string Block { set; get; } = block;
        public string SubCategory { set; get; } = subCat;
        public string DcSpecSheet { set; get; } = targetDcSpec;
    }
}
