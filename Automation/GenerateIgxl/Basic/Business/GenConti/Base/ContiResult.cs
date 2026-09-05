using System.Collections.Generic;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base
{
    public class ContiResult
    {
        public const string FlowSheetName = "Flow_DC_Conti";
        public const string InstanceSheetName = "TestInst_DC_Conti";
        public const string CommonInstanceName = "TestInst_Common";

        public SubFlowSheet ContiFlowSheet { get; set; } = new SubFlowSheet(FlowSheetName);
        public InstanceSheet ConInstanceSheet { get; set; } = new InstanceSheet(InstanceSheetName);
        public List<BinTableRow> ContiBinTableRows { get; set; } = new List<BinTableRow>();
        public InstanceSheet CommonInstance { set; get; } = new InstanceSheet(CommonInstanceName);
        public List<ContiResult> ContiResultList { set; get; } = new List<ContiResult>();
    }
}
