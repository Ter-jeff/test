using System.Collections.Generic;

namespace Automation.Const
{
    public static class MeasType
    {

        public const string MeasI = "MeasI";
        public const string MeasV = "MeasV";
        public const string MeasN = "MeasN";
        public const string MeasR1 = "MeasR1";
        public const string MeasR2 = "MeasR2";
        public const string MeasIdiff = "MeasIdiff";
        public const string MeasVdiff = "MeasVdiff";
        public const string MeasVdiff2 = "MeasVdiff2";
        public const string MeasVocm = "MeasVocm";
        public const string MeasF = "MeasF";
        public const string MeasFdiff = "MeasFdiff";
        public const string MeasC = "MeasC";
        public const string MeasCalcLimit = "CalcLimit";
        public const string MeasCalc = "Calc";
        public const string MeasLimit = "Limit";
        public const string WiMeas = "WiMeas";
        public const string WiSrc = "WiSrc";
        public const string MeasVdm = "MeasVDM";
        public const string MeasWait = "MeasWait";
        public const string MeasDutyCycle = "MeasDutyCycle";
        public const string MeasD = "MeasD";

        public static readonly List<string> MeasTypes = new List<string>
        {
            MeasV,
            MeasI,
            MeasC,
            MeasF,
            MeasIdiff,
            MeasVdiff,
            MeasVdiff2,
            MeasFdiff,
            MeasVocm,
            MeasR1,
            MeasR2,
            MeasCalc,
            MeasLimit,
            MeasCalcLimit,
            WiMeas,
            WiSrc,
            MeasDutyCycle
        };
    }
}
