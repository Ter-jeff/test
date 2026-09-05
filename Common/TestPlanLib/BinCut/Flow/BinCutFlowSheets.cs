using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace TestPlanLib.BinCut.Flow
{
    public partial class BinCutFlowSheets : List<BinCutFlowTables>
    {
        [GeneratedRegex("/s{2,}", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        public List<string> GetFlowNames()
        {
            var flowNames = new List<string>();
            foreach (BinCutFlowTables binCutFlowSheet in this)
            {
                flowNames.AddRange(binCutFlowSheet.SelectMany(x => x.Rows).SelectMany(x => x.Atpg.Split(';').ToList()));
                flowNames.AddRange(binCutFlowSheet.SelectMany(x => x.Rows).SelectMany(x => x.Mbist.Split(';').ToList()));
                flowNames.AddRange(binCutFlowSheet.SelectMany(x => x.Rows).SelectMany(x => x.SpiRtos.Split(';').ToList()));
                flowNames = [.. flowNames.Where(x => x != "0" && !string.IsNullOrEmpty(x)).Distinct().Select(x => x.Trim()).Select(flowName => _regex.Replace(flowName, " "))];
            }
            return flowNames;
        }
    }
}
