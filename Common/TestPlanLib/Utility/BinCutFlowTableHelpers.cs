using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using TestPlanLib.BinCut;
using TestPlanLib.BinCut.Flow;

namespace TestPlanLib.Utility
{
    public static class BinCutFlowTableHelpers
    {
        public static List<BinCutOrderRow> GetHvBinCutOrderRows(List<BinCutOrderRow> binCutOrderRows, EnumBinCutTableBinType enumBinCutTableBinType)
        {
            var hvOrders = new List<BinCutOrderRow>();
            if (enumBinCutTableBinType == EnumBinCutTableBinType.Bin1)
            {
                hvOrders = [.. binCutOrderRows.Where(x => x.Bincut.EqualsIgnoreCase("HVCC"))];
            }
            else if (enumBinCutTableBinType == EnumBinCutTableBinType.BinX)
            {
                hvOrders = [.. binCutOrderRows.Where(x => x.Bincut.EqualsIgnoreCase("HVCC - BinX"))];
            }
            else if (enumBinCutTableBinType == EnumBinCutTableBinType.BinY)
            {
                hvOrders = [.. binCutOrderRows.Where(x => x.Bincut.EqualsIgnoreCase("HVCC - BinY"))];
            }

            if (hvOrders.Count == 0)
            {
                hvOrders = [.. binCutOrderRows.Where(x => x.Bincut.EqualsIgnoreCase("HVCC"))];
            }

            return hvOrders;
        }
    }
}
