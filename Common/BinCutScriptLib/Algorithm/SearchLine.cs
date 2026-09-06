using System.IO;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;

using IgxlLib.Enums;

namespace BinCutScriptLib.Algorithm
{
    internal abstract class SearchLine(EnumJob enumJob, StreamWriter streamWriter)
    {
        protected EnumJob Job = enumJob;
        protected readonly StreamWriter Sw = streamWriter;

        public abstract bool GetInstancesCs(ref OneTouchDown oneTouchDown, out OneGradeSearch oneGradeSearch, ref string tempName, out bool isSearch);

        public abstract bool GetInstances(ref OneTouchDown oneTouchDown, out OneGradeSearch oneGradeSearch, ref string tempName, out bool isSearch);

        public abstract bool IsStartPoint(BinCutLineBase binCutLineBase);

        public abstract bool IsStepStartCs(BinCutLineBase binCutLineBase);

        public abstract bool IsStopPoint(BinCutLineBase binCutLineBase);

        public abstract bool IsStopPointCs(BinCutLineBase binCutLineBase);

        public abstract bool IsStepEndCs(BinCutLineBase binCutLineBase);

        public abstract bool GetReturn(OneTouchDown oneTouchDown);
    }
}
