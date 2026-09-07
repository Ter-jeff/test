using System.IO;

using BinCutScriptLib.Base;

using IgxlLib.Enums;

namespace BinCutScriptLib.Algorithm.GradeSearch
{
    internal class GradeSearchHv(EnumJob enumJob, StreamWriter streamWriter, SearchLine searchLine) : GradeSearchBase(enumJob, streamWriter, searchLine)
    {
        public override bool GetInstances(ref OneTouchDown oneTouchDown, out OneGradeSearch oneGradeSearch, ref string tempName, out bool isSearch)
        {
            oneGradeSearch = new OneGradeSearch();
            var oneStep = new OneStep();
            bool isSelSram = false;
            isSearch = false;
            if (!SearchLine.GetReturn(oneTouchDown))
            {
                return false;
            }

            //Search until found Initial_Voltage
            bool? returnFlag = null;
            int oneTouchIndex = GetStartIndex(oneTouchDown, ref oneGradeSearch, ref oneStep, ref returnFlag);
            if (returnFlag != null)
            {
                return (bool)returnFlag;
            }

            bool flag = HvSearch(ref oneTouchDown, ref oneGradeSearch, ref oneStep, ref oneTouchIndex, ref isSelSram);
            PostPorcess(oneTouchDown, oneGradeSearch, oneTouchIndex, ref tempName, ref isSelSram, isSearch);

            return flag;
        }

        public bool GetInstancesCs(ref OneTouchDown oneTouchDown, out OneGradeSearch oneGradeSearch, ref string tempName)
        {
            bool isSelSram = false;
            oneGradeSearch = new OneGradeSearch();
            var oneStep = new OneStep();

            if (!SearchLine.GetReturn(oneTouchDown))
            {
                return false;
            }

            int oneTouchIndex = GetStartIndexCs(oneTouchDown, ref oneGradeSearch, out bool returnFlag);

            if (returnFlag)
            {
                return !returnFlag;
            }

            bool flag = HvSearchCs(ref oneTouchDown, ref oneGradeSearch, ref oneStep, ref oneTouchIndex, ref isSelSram);

            bool isSearch = false;
            PostPorcess(oneTouchDown, oneGradeSearch, oneTouchIndex, ref tempName, ref isSelSram, isSearch);

            return flag;
        }

        public int GetStartIndex(OneTouchDown oneTouchDown, ref OneGradeSearch oneGradeSearch, ref OneStep oneStep, ref bool? returnFlag)
        {
            int oneTouchIndex;
            for (oneTouchIndex = 0; oneTouchIndex < oneTouchDown.Lines.Count; oneTouchIndex++)
            {
                if (SearchLine.IsStopPoint(oneTouchDown.Lines[oneTouchIndex]))
                {
                    RemoveLines(oneTouchDown, oneTouchIndex);
                    returnFlag = false;
                    return oneTouchIndex;
                }

                if (IsCallInstanceLine(oneTouchDown, oneTouchIndex))
                {
                    oneGradeSearch.IsCallInstance = true;
                }

                if (SearchLine.IsStartPoint(oneTouchDown.Lines[oneTouchIndex]))
                {
                    return oneTouchIndex;
                }

                #region read data
                if (GetBvLine(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }

                if (GetDsscSelSram(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }

                if (GetOffSetLine(oneTouchDown, ref oneStep, oneTouchIndex))
                {
                    continue;
                }

                #endregion
            }
            returnFlag = false;
            return oneTouchIndex;
        }
    }
}
