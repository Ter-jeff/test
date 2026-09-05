using System.Collections.Generic;

using TestPlanLib.Basic;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public class EfuseInstanceSheetChecker : BinCutInstanceSheetChecker
    {
        public void WorkFlow(BinCutInstanceSheet binCutInstanceSheet, Dictionary<string, PatternData>? patternDatas = null, Dictionary<string, int>? timeSetDic = null)
        {
            SetType();
            CheckDuplicatedInstance(binCutInstanceSheet);

            if (timeSetDic!.Count != 0)
            {
                CheckTimeSet(binCutInstanceSheet, patternDatas!, timeSetDic);
            }

            CheckPatternTimeSet(binCutInstanceSheet, patternDatas!);

            CheckUserDefinePatternSetName(binCutInstanceSheet);
        }

        protected override void SetType()
        {
            Type = EnumInstanceSheetType.Efuse;
        }
    }
}
