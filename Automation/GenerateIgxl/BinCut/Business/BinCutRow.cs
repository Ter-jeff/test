using System.Collections.Generic;

using Automation.GenerateIgxl.BinCut.Base;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.BinCut.Business
{
    public class BinCutRow
    {
        //Input
        public BinCutSourceItem BinCutSourceItem { get; set; }
        public List<BinCutFinalInstanceRow> BinCutFinalInstanceRows { get; set; }
        //Output
        public List<InstanceRow> InstanceRowDetail { get; set; }
    }

    public class BinCutRowForSort
    {
        public BinCutSourceItem BinCutSourceRow;
        public BinCutFinalInstanceRow BinCutFinalInstanceRow;
    }

    public class BinCutInstanceForReName
    {
        public List<BinCutFinalInstanceRow> BinCutFinalInstanceRows;
        public InstanceRow InstanceRowDetail;
    }

    public class BinCutExtraPolationModeInfo
    {
        public bool IsOnlyExtraMode { get; set; }
        public bool IsExtraPolation { get; set; }
    }
}
