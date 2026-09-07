using System.Collections.Generic;
using System.Linq;

namespace BinCutScriptLib.Base
{
    public class BvName
    {
        public List<int> Sites = [];
        public string Name = string.Empty;
        public string NameRemoveBv = string.Empty;
        public string Mode = string.Empty;
        public int Index = -1;

        public BvName()
        {
        }

        public BvName(BvName bvName)
        {
            if (bvName == null)
            {
                return;
            }

            Sites = bvName.Sites?.ToList() ?? [];
            Name = bvName.Name;
            NameRemoveBv = bvName.NameRemoveBv;
            Mode = bvName.Mode;
            Index = bvName.Index;
        }

        public BvName Copy()
        {
            return new BvName(this);
        }
    }
}
