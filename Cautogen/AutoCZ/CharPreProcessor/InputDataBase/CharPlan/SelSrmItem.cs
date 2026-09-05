using System.Collections.Generic;
using System.Linq;

namespace Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan
{
    public class SelSrmItem
    {
        public string Stage;
        public string Block;
        public string Pattern { get; set; }
        public List<SelSrmRow> Rows = new List<SelSrmRow>();
        public List<string> LogicPins
        {
            get { return Rows.Select(p => p.LogicPins).ToList(); }
        }

        public List<string> SramPins
        {
            get { return Rows.Select(p => p.SramPins).ToList(); }
        }


        public List<string> Selsrm0Set
        {
            get { return Rows.Select(p => p.Selsrm0).ToList(); }
        }
        public List<string> Selsrm1Set
        {
            get { return Rows.Select(p => p.Selsrm1).ToList(); }
        }

    }

}
