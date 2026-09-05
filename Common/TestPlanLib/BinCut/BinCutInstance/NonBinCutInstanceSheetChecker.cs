using System.Collections.Generic;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc;
using TestPlanLib.Utility;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public class NonBinCutInstanceSheetChecker : BinCutInstanceSheetChecker
    {
        public void WorkFlow(BinCutInstanceSheet binCutInstanceSheet, Dictionary<string, PatternData>? patternDatas = null, Dictionary<string, int>? timeSetDic = null, List<UfDigSrcSheet>? ufDigSrcSheets = null)
        {
            SetType();

            CheckDuplicatedInstance(binCutInstanceSheet);

            if (timeSetDic != null && timeSetDic.Count != 0)
            {
                CheckTimeSet(binCutInstanceSheet, patternDatas!, timeSetDic);
            }

            CheckPatternTimeSet(binCutInstanceSheet, patternDatas!);

            CheckUserDefinePatternSetName(binCutInstanceSheet);

            CheckVoltageType(binCutInstanceSheet);

            CheckUserFunction(binCutInstanceSheet, ufDigSrcSheets);
        }

        protected override void SetType()
        {
            Type = EnumInstanceSheetType.Scan;
        }

        protected void CheckVoltageType(BinCutInstanceSheet binCutInstanceSheet)
        {
            foreach (BinCutInstanceRow row in binCutInstanceSheet.Rows)
            {
                string typeByFlowName = BinCutInstanceRowUtility.GetTypeByFlowNameOrDcCategory(row.FlowName);
                string typeByDCcategory = BinCutInstanceRowUtility.GetTypeByFlowNameOrDcCategory(row.DCcategory);
                if (typeByDCcategory != "UnknowType" && typeByFlowName != "UnknowType")
                {
                    if (!typeByFlowName.EqualsIgnoreCase(typeByDCcategory))
                    {
                        if (Type == EnumInstanceSheetType.Bincut)
                        {
                            binCutInstanceSheet.AddError(BinCutErrorType.E_VoltageType_01, binCutInstanceSheet.SheetName, row.RowNum, 0, "Voltage type is different in the Sub Flow column and Voltage Category column !!!");
                        }
                        else
                        {
                            binCutInstanceSheet.AddError(ScanErrorType.E_VoltageType_01, binCutInstanceSheet.SheetName, row.RowNum, 0, "Voltage type is different in the Sub Flow column and Voltage Category column !!!");
                        }
                    }
                }
            }
        }
    }
}
