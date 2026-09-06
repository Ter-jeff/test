using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BistBira.NewLogicData;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using ScghLib.Enums;
using ScghLib.Reader;
using ScghLib.Utility;

using TestPlanLib.Basic;

namespace Automation.GenerateIgxl.BistBira.Base
{
    public class BistBurst
    {
        private readonly BistNaming _naming;
        private readonly List<PatSet> _patSets = new List<PatSet>();

        public BistBurst(BistNaming naming)
        {
            _naming = naming;
        }

        public BistProdFlowSheet BurstLabel(BistProdFlowSheet prodFlowSheet, Dictionary<string, PatternData> patternDatas, MbistPatSetType mbistPatSetType, MbistDataStore mbistDataStore)
        {
            var burstPattern = new List<string>();
            var burstFlowSheet = new BistProdFlowSheet { MbistSheet = prodFlowSheet.MbistSheet };
            for (int i = 1; i < prodFlowSheet.Rows.Count; i++)
            {
                if (prodFlowSheet.Rows[i].Label == prodFlowSheet.Rows[i - 1].Label &&
                    /*prodFlowSheet.DataRows[i].Action == prodFlowSheet.DataRows[i - 1].Action &&*/
                    prodFlowSheet.Rows[i].Voltage == prodFlowSheet.Rows[i - 1].Voltage &&
                    prodFlowSheet.Rows[i].FailBranch == prodFlowSheet.Rows[i - 1].FailBranch &&
                    prodFlowSheet.Rows[i].TimeSet == prodFlowSheet.Rows[i - 1].TimeSet)
                {
                    //combine pattern
                    burstPattern.Add(prodFlowSheet.Rows[i - 1].Pattern);
                }
                else
                {
                    CheckBurstMismatch(prodFlowSheet.Rows[i - 1], prodFlowSheet.Rows[i]);
                    string pFlowPattern = prodFlowSheet.Rows[i - 1].Pattern;
                    if (BistAction.GetActionType(prodFlowSheet.Rows[i - 1]) == BistActionType.RunPattern)
                    {
                        burstPattern.Add(pFlowPattern);

                        BistProdFlowRow burstRow = prodFlowSheet.Rows[i - 1].Copy();
                        burstFlowSheet.Rows.Add(burstRow);
                        burstFlowSheet.Rows.Last().Pattern = GenPatSet(burstFlowSheet.Rows.Last().Label, burstPattern,
                            burstFlowSheet.MbistSheet.SheetName, mbistPatSetType);
                        burstFlowSheet.Rows.Last().IsPatBurst = true;
                        burstFlowSheet.Rows.Last().BurstPatterns.AddRange(burstPattern);
                        // copy pattern obj for burst row
                        if (patternDatas != null && patternDatas.ContainsKey(pFlowPattern.ToLower()))
                        {
                            if (!patternDatas.ContainsKey(burstRow.Pattern.ToLower()))
                            {
                                patternDatas.Add(burstRow.Pattern.ToLower(), patternDatas[pFlowPattern.ToLower()]);
                            }
                        }
                    }
                    else
                    {
                        burstPattern.Add(prodFlowSheet.Rows[i - 1].Pattern);

                        burstFlowSheet.Rows.Add(prodFlowSheet.Rows[i - 1]);
                    }
                    burstPattern.Clear();
                }

                if (i == prodFlowSheet.Rows.Count - 1)
                {
                    burstFlowSheet.Rows.Add(prodFlowSheet.Rows.Last());
                }
            }

            string sheetName = prodFlowSheet.MbistSheet.SheetName;
            if (!mbistDataStore.DicPatSets.ContainsKey(sheetName))
            {
                mbistDataStore.DicPatSets.Add(sheetName, _patSets);
            }
            else
            {
                mbistDataStore.DicPatSets[sheetName].AddRange(_patSets);
            }

            return burstFlowSheet;
        }

        private void CheckBurstMismatch(BistProdFlowRow previousRow, BistProdFlowRow currentRow)
        {
            if (currentRow.Label == previousRow.Label)
            {
                if (currentRow.Voltage != previousRow.Voltage)
                {
                    Response.Report($"Column[Voltage]: Same label with previous row, but Voltage mismatch at row {currentRow.RowNum}.", EnumMessageLevel.Error);
                    ErrorReportManager.AddError(MbistErrorType.E_BurstInfoMismatch_01, currentRow.SheetName, currentRow.RowNum, 1, "Column[Voltage]: Same label with previous row, but Voltage mismatch.");
                }

                if (currentRow.FailBranch != previousRow.FailBranch)
                {
                    Response.Report($"Column[FailBranch]: Same label with previous row, but FailBranch mismatch at row {currentRow.RowNum}.", EnumMessageLevel.Error);
                    ErrorReportManager.AddError(MbistErrorType.E_BurstInfoMismatch_02, currentRow.SheetName, currentRow.RowNum, 1, "Column[FailBranch]: Same label with previous row, but FailBranch mismatch.");
                }

                if (currentRow.TimeSet != previousRow.TimeSet)
                {
                    Response.Report($"Column[TimeSet]: Same label with previous row, but TimeSet mismatch at row {currentRow.RowNum}.", EnumMessageLevel.Error);
                    ErrorReportManager.AddError(MbistErrorType.E_BurstInfoMismatch_03, currentRow.SheetName, currentRow.RowNum, 1, "Column[TimeSet]: Same label with previous row, but TimeSet mismatch.");
                }
            }
        }

        private string GenPatSet(string label, List<string> burstPatternList, string sheetName, MbistPatSetType mbistPatSetType)
        {
            bool biraFlag = false;
            var patSet = new PatSet { PatSetName = $"PatSet_{label}_{sheetName.Split('_').Last()}" };

            foreach (string pattern in burstPatternList)
            {
                var patSetRow = new PatSetRow { File = pattern, Burst = mbistPatSetType == MbistPatSetType.BurstNo ? "No" : "Yes" };
                patSet.AddRow(patSetRow);
                if (_naming.IsBira(pattern, label))
                {
                    biraFlag = true;
                }
            }
            string biraInfo = biraFlag ? "_BIRA_" : "";
            patSet.PatSetName = $"PatSet_{biraInfo}{label}_{sheetName.Split('_').Last()}";

            _patSets.Add(patSet);

            return patSet.PatSetName;
        }

        public bool HasPatSet
        {
            get { return _patSets.Count > 0; }
        }

        public void GenPatSetSheet(string sheetName, MbistDataStore mbistDataStore)
        {
            string module = _naming.GetModule(sheetName);
            string sName = "PatSet_" + sheetName;//+ "Mbist_Burst";
            List<PatSet> pSets = mbistDataStore.DicPatSets[sheetName];
            if (pSets.Any())
            {
                var patSetSheet = new PatSetSheet(sName);
                patSetSheet.AddRows(pSets);

                TestProgram.IgxlWorkBk.AddPatSetSheet(BistBiraMain.GetFolder(module, _naming), patSetSheet);
            }
        }
    }
}
