using System.Collections.Generic;

using Automation.Singleton;

using IgxlLib.IgxlBase;

using TestPlanLib.BinNumber;
using TestPlanLib.Singleton;

namespace Automation.GenerateIgxl.Basic.Business.GenNwire.Business
{
    public class NwireBintable
    {
        public List<BinTableRow> GenerateFlow()
        {
            List<BinTableRow> nWireBinTableRows = new List<BinTableRow>();
            BinTableRow row = new BinTableRow
            {
                Name = NwireNaming.GetBinTableName(),
                ItemList = NwireSingleton.NwireFlag,
                Op = "AND"
            };
            row.Items.Add("T");

            string[] binNameSegments = row.Name.Split('_');
            string category1 = binNameSegments.Length > 1 ? binNameSegments[1] : "";
            string category2 = binNameSegments.Length > 2 ? binNameSegments[2] : "";
            BinNumResult binInfo = BinNumberSingleton.Instance.GetBinInfo("Nwire", category1, category2, row);

            row.Sort = binInfo.SoftBin.ToString("G15");
            row.Bin = binInfo.BinNumInfo.HardBin.ToString("G15");
            row.Result = binInfo.BinNumInfo.Status;

            nWireBinTableRows.Add(row);
            return nWireBinTableRows;
        }
    }
}
