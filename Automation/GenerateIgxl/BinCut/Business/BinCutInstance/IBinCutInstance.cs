using System.Collections.Generic;


using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.BinCut.Business.BinCutInstance
{
    public interface IBinCutInstance
    {
        InstanceRow GenerateInstance();
        List<InstanceRow> GenerateInstanceByTestName();
        FlowRow GenerateFlowRow(bool isHvccOrPost, bool isTmps, bool isCsharp);
        FlowRow GetBinTableRow();
        FlowRow GetBinTableRowBv();
        string GetBinTableName();
    }
}
