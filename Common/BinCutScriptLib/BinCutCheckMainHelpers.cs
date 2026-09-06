using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Base;

using CommonLib.Extension;

namespace BinCutScriptLib
{
    internal static class BinCutCheckMainHelpers
    {
        internal static void AddBvName(BvName bvName, ref List<BvName> bvNames)
        {
            if (!bvNames.Exists(x => x.Name.EqualsIgnoreCase(bvName.Name)))
            {
                bvNames.Add(bvName);
            }
            else
            {
                BvName foundBvName = bvNames.Find(x => x.Name.EqualsIgnoreCase(bvName.Name))!;
                foundBvName.Sites = [.. foundBvName.Sites.Concat(bvName.Sites).Distinct().OrderBy(x => x)];
            }
        }
    }
}
