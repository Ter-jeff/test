using System.Collections.Generic;
using System.Linq;

namespace Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure
{
    public class CharPlanSheet
    {
        /* Property */
        public string SheetName { get; set; }
        public bool IsHardIp { get; set; }

        public List<string> GetReferencePayloads()
        {
            return CharList.SelectMany(p => p.UsedPayloads).Distinct().ToList();
        }

        public List<CharPlanItem> CharList { get; set; }

        /* Constructor */
        public CharPlanSheet(string sheetName)
        {
            SheetName = sheetName;
            CharList = new List<CharPlanItem>();
        }
    }
}
