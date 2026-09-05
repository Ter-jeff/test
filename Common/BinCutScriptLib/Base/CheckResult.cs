namespace BinCutScriptLib.Base
{
    public class CheckResult
    {
        public bool IsBvPass = true;
        public bool IsInterpolationPass = true;
        public bool IsPowerBinningPass = true;
        public bool IsDsscPass = true;
        public bool IsLvResultPass = true;
        public bool IsBinoutStatusPass = true;
        public string CheckPass = "P";
        public int PatPassCnt;
        public int PatFailCnt;

        public CheckResult() { }

        public CheckResult(CheckResult checkResult)
        {
            if (checkResult == null)
            {
                return;
            }

            IsBvPass = checkResult.IsBvPass;
            IsInterpolationPass = checkResult.IsInterpolationPass;
            IsPowerBinningPass = checkResult.IsPowerBinningPass;
            IsDsscPass = checkResult.IsDsscPass;
            IsLvResultPass = checkResult.IsLvResultPass;
            IsBinoutStatusPass = checkResult.IsBinoutStatusPass;
            CheckPass = checkResult.CheckPass;
            PatPassCnt = checkResult.PatPassCnt;
            PatFailCnt = checkResult.PatFailCnt;
        }

        public CheckResult Copy()
        {
            return new CheckResult(this);
        }
    }
}
