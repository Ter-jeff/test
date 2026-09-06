using System.Text.RegularExpressions;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.InputManager.Data;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.GenerateIgxl.BinCut.Business.BinCutInstance
{
    public class BinCutInstanceInterfaceFactory
    {
        public IBinCutInstance GetBinCutInstanceInterface(BinCutFinalInstanceRow dataRow, BinCutSourceItem sourceRow, BinCutInputData binCutInputManager, bool isPost)
        {
            if (isPost)
            {
                if (Regex.IsMatch(dataRow.BinCutInstanceRow.FlowName, "ILB", RegexOptions.IgnoreCase))
                {
                    return new BinCutInstanceIlb(dataRow, sourceRow, binCutInputManager);
                }

                if (Regex.IsMatch(dataRow.BinCutInstanceRow.FlowName, "ELB", RegexOptions.IgnoreCase))
                {
                    return new BinCutInstanceElb(dataRow, sourceRow, binCutInputManager);
                }

                return new BinCutInstancePost(dataRow, sourceRow, binCutInputManager);
            }

            if (Regex.IsMatch(sourceRow.PerformanceMode, "TMPS", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(sourceRow.ColumnContent, "TEMP SENSOR", RegexOptions.IgnoreCase))
            {
                return new BinCutInstanceTmps(dataRow, sourceRow, binCutInputManager);
            }

            if (Regex.IsMatch(dataRow.BinCutInstanceRow.FlowName, "ILB", RegexOptions.IgnoreCase))
            {
                return new BinCutInstanceIlb(dataRow, sourceRow, binCutInputManager);
            }

            if (Regex.IsMatch(dataRow.BinCutInstanceRow.FlowName, "ELB", RegexOptions.IgnoreCase))
            {
                return new BinCutInstanceElb(dataRow, sourceRow, binCutInputManager);
            }

            if (sourceRow.ColumnName == EnumColumnName.TD)
            {
                return new BinCutInstanceTdDdr(dataRow, sourceRow, binCutInputManager);
            }

            if (sourceRow.ColumnName == EnumColumnName.Mbist)
            {
                return new BinCutInstanceBist(dataRow, sourceRow, binCutInputManager);
            }

            if (sourceRow.ColumnName == EnumColumnName.FUNC && sourceRow.GetDomainOfMode() == "DDR")
            {
                return new BinCutInstanceDdrLb(dataRow, sourceRow, binCutInputManager);
            }


            return new BinCutInstanceTdDdr(dataRow, sourceRow, binCutInputManager);
        }
    }
}
