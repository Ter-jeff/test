namespace BinCutScriptLib.Base
{
    public class PatternInfo
    {
        public string PatternName = string.Empty;
        public int Bin;
        public bool IsFail;
        public int EqName;
        public double Cp;
        public int StepNum;

        public PatternInfo() { }

        public PatternInfo(PatternInfo patternInfo)
        {
            if (patternInfo == null)
            {
                return;
            }

            PatternName = patternInfo.PatternName;
            Bin = patternInfo.Bin;
            IsFail = patternInfo.IsFail;
            EqName = patternInfo.EqName;
            Cp = patternInfo.Cp;
            StepNum = patternInfo.StepNum;
        }

        public PatternInfo Copy()
        {
            return new PatternInfo(this);
        }
    }
}
