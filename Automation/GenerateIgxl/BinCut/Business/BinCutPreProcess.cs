using System.Collections.Generic;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.Scan.NonBinCut;
using Automation.InputManager.Data;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.GenerateIgxl.BinCut.Business
{
    internal class BinCutPreProcess
    {
        private readonly BinCutInputData _binCutInputManager;

        public BinCutPreProcess(BinCutInputData binCutInputManager)
        {
            _binCutInputManager = binCutInputManager;
        }

        public BinCutFinalInstanceRows GetBinCutInstanceRows(List<BinCutInstanceSheet> binCutInstanceSheets, List<BinCutFinalInstanceRow> previousRows = null)
        {
            var instSheetPreProcessBinCut = new InstSheetPreProcessBinCut(_binCutInputManager.Config);
            BinCutFinalInstanceRows binCutFinalInstanceRows = instSheetPreProcessBinCut.InitialInstance(binCutInstanceSheets, _binCutInputManager.BinCutInstanceNamingSheet);
            return binCutFinalInstanceRows.RePatSetNameDuplicateRowBinCut(previousRows);
        }
    }
}
