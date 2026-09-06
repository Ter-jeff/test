using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlSheets;

using TestPlanLib.BinCut;

namespace Automation.GenerateIgxl.BinCut.Business
{
    public class BinCutPostFlowInstanceWriter : BinCutFlowInstanceWriter
    {
        private readonly BinCutInputData _binCutInputManager;
        private readonly List<BinCutFinalInstanceRow> _binCutFinalInstanceRows;

        public BinCutPostFlowInstanceWriter(BinCutInputData binCutInputManager, List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
            : base(binCutInputManager, binCutFinalInstanceRows)
        {
            _binCutInputManager = binCutInputManager;
            _binCutFinalInstanceRows = binCutFinalInstanceRows;
        }

        public void GetBinCutOutsideResult(Dictionary<string, List<BinCutSourceItem>> sourceRowDic, out BinCutPatternReport binCutPostPatternReport)
        {
            var instanceSheet = new InstanceSheet("TestInst_Vddbinning_OutsideBV");
            var binCutFlowSheets = new List<SubFlowSheet>();
            binCutPostPatternReport = new BinCutPatternReport();

            const bool isPost = true;
            var instanceRowList = new InstanceRows();
            foreach (KeyValuePair<string, List<BinCutSourceItem>> item in sourceRowDic)
            {
                var binCutInstanceAndSourceRow = new BinCutInstanceHvGenerator(item.Value, _binCutInputManager, _binCutFinalInstanceRows);
                binCutInstanceAndSourceRow.GenInstanceRows(isPost, BinCutInputData);
                binCutInstanceAndSourceRow.BinCutInstanceRowMergeByJob();
                binCutInstanceAndSourceRow.BinCutInstanceNameCheck(); //if the instance name is duplicated, the flag turn true
                List<BinCutRowForSort> reOrderResult = binCutInstanceAndSourceRow.ReArrangeByOrderOption(item.Key, false);
                instanceRowList.AddRange(binCutInstanceAndSourceRow.GetInstanceRows(item.Key, reOrderResult, out List<BinCutPatternRow> binCutPatternRowsByMode, isPost, false));
                bool isCSharp = instanceRowList.Any(x => x.VbtName.EndsWith("BinCutTest", StringComparison.CurrentCultureIgnoreCase));
                binCutFlowSheets.Add(binCutInstanceAndSourceRow.GenerateFlowRows(item.Key, reOrderResult, isPost, isCSharp));
                binCutPostPatternReport.Rows.AddRange(binCutPatternRowsByMode);

            }

            #region Flow_Vddbinning and instance
            foreach (SubFlowSheet sheet in binCutFlowSheets)
            {
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirBinCut, sheet);
            }

            var newInstanceRowList = new InstanceRows();
            newInstanceRowList.AddRange(instanceRowList.Where(x => !x.TestName.ContainsIgnoreCase("_QA_")).ToList());
            newInstanceRowList.AddRange(instanceRowList.Where(x => x.TestName.ContainsIgnoreCase("_QA_")).ToList());
            instanceSheet.Rows = newInstanceRowList;
            instanceSheet.RemoveDuplicateInstance();

            #region Header/Footer for instance
            instanceSheet.Rows.AddRange(GenHeaderFooterRows(sourceRowDic));
            #endregion

            TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirBinCut, instanceSheet);
            #endregion
        }
    }
}
