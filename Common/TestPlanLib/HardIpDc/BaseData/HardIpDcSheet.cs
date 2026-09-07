using System.Collections.Generic;
using System.Linq;

namespace TestPlanLib.HardIpDc.BaseData
{
    public class HardIpDcSheet
    {
        public List<HardIpCategoryDef> Rows { set; get; } = [];

        public List<HardIpSpecValue> GetDistcRatioSpecValues()
        {
            List<HardIpSpecValue> allSpecValues = [];
            foreach (HardIpCategoryDef categoryDef in Rows)
            {
                foreach (HardIpDcRow dataRow in categoryDef.DataRows)
                {
                    allSpecValues.AddRange(categoryDef.GetSpecValueFromDef(dataRow));
                }
            }

            allSpecValues = [.. allSpecValues.GroupBy(p => new { p.PinName, p.Parameter }).Select(p => p.First())];
            return allSpecValues;
        }
    }
}
