using System.IO;
using System.Linq;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;

using IgxlLib.Enums;

namespace BinCutScriptLib.Algorithm
{
    internal class GetCp1LvTests(EnumJob enumJob, StreamWriter streamWriter) : SearchLine(enumJob, streamWriter)
    {
        public override bool GetInstances(ref OneTouchDown oneTouchDown, out OneGradeSearch oneGradeSearch, ref string tempName, out bool isSearch)
        {
            return new GradeSearchLv(Job, Sw, this).GetInstances(ref oneTouchDown, out oneGradeSearch, ref tempName, out isSearch);
        }

        public override bool GetInstancesCs(ref OneTouchDown oneTouchDown, out OneGradeSearch oneGradeSearch, ref string tempName, out bool isSearch)
        {
            return new GradeSearchLv(Job, Sw, this).GetInstancesCs(ref oneTouchDown, out oneGradeSearch, ref tempName, out isSearch);
        }

        public override bool IsStartPoint(BinCutLineBase binCutLineBase)
        {
            return binCutLineBase.IsCp1LvStartPoint();
        }

        public override bool IsStepStartCs(BinCutLineBase binCutLineBase)
        {
            return binCutLineBase.IsCp1LvStartPointCsharp();
        }

        public override bool IsStopPoint(BinCutLineBase binCutLineBase)
        {
            return binCutLineBase.IsCp1LvStopPoint();
        }

        public override bool IsStepEndCs(BinCutLineBase binCutLineBase)
        {
            return binCutLineBase.IsCp1LvStopPointCsharp();
        }

        public override bool GetReturn(OneTouchDown oneTouchDown)
        {
            if (oneTouchDown.Lines.Count == 0)
            {
                return false;
            }

            if (IsStopPoint(oneTouchDown.Lines.First()))
            {
                return false;
            }

            return true;
        }

        public override bool IsStopPointCs(BinCutLineBase binCutLineBase)
        {
            throw new System.NotImplementedException();
        }
    }
}
