
using System.Diagnostics.CodeAnalysis;

namespace Automation.GenerateIgxl.Wireless.DVDC.WirelessConst
{
    [ExcludeFromCodeCoverage]
    public class WirelessConstData
    {
        //VBT arg of DVDC_Trim_Universal_func
        public const string TrimFuseName = "TrimFuseName";
        public const string TrimRegName = "TrimRegName";
        public const string TrimTarget = "TrimTarget";
        public const string TrimType = "TrimType";
        //Plan Header for Wireless FW
        public const string RfCalc = "RF Test Type";
        public const string RfInterpose = @"RF\s*Interpose";
        public const string FwReadBackSequence = "ReadBack Sequence";
        public const string DigSrcpin = "DigSrcPin";
        public const string DigSrcEquation = "DigSrcEquation";
        public const string DigSrcAssignment = "DigSrcAssignment";
        public const string TrimCalcEqn = "BestCodeCalcFunc";
        public const string UpdateBstc2FuseDataBaseFlag = "UpdateBSTC2FuseDataBase_Flag";

        public const string Slope = "Slope";
        public const string SweepRange = "SweepRange";
        public const string SweepRange2D = "DimensionSweepRange";
        public const string LoadFile = "LoadFile";
        public const string TransitionSlope = "TransitionSlope";
        public const string PutLimitsAllCodes = "putLimitsAllCodes";
        public const string DoTrimming = "doTrimming";
        public const string DoFinalVerification = "doFinalVerification";
        public const string DigSrcBitWidth = "DigSrcBitWidth";
        public const string AdjustTrimCode = "AdjustTrimCode";
        public const string ApplyLevelsTimings = "applyLevelsTimings"; // do measure
        public const string ApplyLevelsTiming = "applyLevelsTiming"; //trim func
        public const string Stepsize = "Stepsize";
        public const string WriteFuncResult = "WriteFuncResult";
        public const string FwFlag = "FW_Flag";

        //VBT:MeasUniversalFunc
        public const string TestSequence = "testSequence";
        public const string MeasPins = "measPins";
        public const string ForceConditions = "forceConditions";
        public const string MeasIRange = "measIRange";
        public const string MeasVRange = "measVRange";
        public const string MeasWaitTime = "measWaitTime";
        public const string InterposePrePat = "interposePrePat";
        public const string InterposePreMeas = "interposePreMeas";
        public const string InterposePostMeas = "interposePostMeas";
        public const string InterposePostPat = "interposePostPat";
        public const string MeasName = "measName";
        public const string MeasStoreName = "measStoreName";
        public const string CalcStoreName = "calcStoreName";
        public const string CapStoreName = "capStoreName";
        public const string MeasLimit = "measLimit";
        public const string CalcEquName = "calcEquName";
        public const string IsWriteFuncResult = "isWriteFuncResult";
        public const string PinsKeepForceFlag = "pinsKeepForceFlag";
        public const string DSSCSetup = "DSSCSetup";
        public const string Validating = "Validating_";

    }
}
