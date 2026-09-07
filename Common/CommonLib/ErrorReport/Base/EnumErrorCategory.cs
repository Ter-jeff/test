using System.ComponentModel;

namespace CommonLib.ErrorReport.Base
{
    public enum EnumErrorCategory
    {
        [Description("AI")] AutoAi,
        [Description("BA")] Basic,
        [Description("BC")] BinCut,
        [Description("BI")] Mbist,
        [Description("BO")] BinOutReport,
        [Description("CL")] ClockCheck,
        [Description("CZ")] Char,
        [Description("DC")] Conti,
        [Description("EF")] EFuse,
        [Description("EV")] Evs,
        [Description("ES")] EfuseCheck,
        [Description("FM")] FlowMain,
        [Description("HP")] HardIp,
        [Description("HT")] Htol,
        [Description("HV")] Harvest,
        [Description("ID")] Ids,
        [Description("PA")] PreAction,
        [Description("PO")] PostAction,
        [Description("RT")] Rtos,
        [Description("SC")] Scan,
        [Description("SN")] Ssn,
    }
}
