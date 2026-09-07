using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;

namespace Automation.GenerateIgxl.BinCut.Business
{
    public class BinCutPatternSearch
    {
        private readonly List<BinCutFinalInstanceRow> _binCutFinalInstanceRows;
        private readonly BinCutSourceItem _sourceSheetRow;

        public BinCutPatternSearch(BinCutSourceItem sourceSheetRow, List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            _binCutFinalInstanceRows = binCutFinalInstanceRows;
            _sourceSheetRow = sourceSheetRow;
        }

        public List<BinCutFinalInstanceRow> GetPatterns()
        {
            return GetBinCutInstDataRowsByFlowName(_sourceSheetRow, _binCutFinalInstanceRows);
        }

        private List<BinCutFinalInstanceRow> GetBinCutInstDataRowsByFlowName(BinCutSourceItem binCutSourceItem, List<BinCutFinalInstanceRow> binCutFinalInstanceRow)
        {
            var targetRows = binCutFinalInstanceRow.Where(binCutSourceItem.JudgeIsTargetFlow).ToList();
            targetRows.ForEach(x => x.IsUsed = true);
            return targetRows.Select(p => (BinCutFinalInstanceRow)p.Clone(binCutSourceItem.Job)).ToList();
        }
    }
}
