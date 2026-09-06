using System.Collections.Generic;

namespace Cautogen.AutoCZ.CharPreProcessor.ReportManager
{
    //public enum MessageLevel
    //{
    //    Error,
    //    Warning,
    //    Message,
    //    Result
    //};

    public enum ErrorLevel
    {
        Error,
        Warning
    };

    public enum ErrorType
    {
        IllegalChar,
        DuplicateTpName,
        DuplicatePattern,
        ErrorShmooRange,
        ErrorMethod,
        ErrorFunctional,
        DupCategory,
        OppositeUsl,
        MissingCategory,
        MissingPinName,
        MissingPatternInPatList,
        MissingPattern,
        WrongGroup,
        MissingHeader,
        WrongRetention,
        MissingShmooValue,
        MissingShmooCondition,
        DummyShmooCondition,
        WrongTpName,
        WrongTpNameEquation,
        WrongUserdef1,
        WrongMeasOfHac,
        WrongUserdef3OfHac,
        WrongMeasCount,
        WrongMeasPin,
        MissingPinSeq,
        MissingPatternInPatInfo,
        WrongUserdef1OfVih,
        WrongShmooNv,
        MissingPinInGlobalSpecs,
        NotMatchRuleEquation,
        NoOtherSupplies,
        WrongVddInPinColumn,
        Over2ShmooTracking,
        PatternOtherThanInPayload1,
        PatternsWithoutPayload,
        MissingHlnCondition,
        TooManyInstanceName,
        MixedSIandDm,
        PinGroupNotMatch,
        WrongUsllslRange,
        MissingForceCondition,
        WrongCapture,
        Userdef3MismatchShmoo,
        WrongShmooSteps,
        HeatAlarm,
        LessThan6ShmooPoints,
        WrongNetName,
        WrongLimitFormat,
        WrongDigSrc,
        MissingRevisonSheet,
        MisMatchDigSrc,
        TestNameOverLength,
        TestNameNotEndswithSingleUnderline,
        EmptySheet,
        EmptyUse,
        WrongPowerRunScenario,
        EmptyHeader,
        PowerGroupApply,
        RtosUserdef6Syntax,
        SelSrmBlockPatNotFound,
        FSelSrmSgmtNotFound,
        TrackingPinMismatchPrimaryStep,
        PrimaryTrackingPinValtInconsist,
        SramPinNotTieLogic,
        InitPatternbehindSelsram,
        SelsramPatternNotFoundInMappingTable,
        SelsramBitMismatch,
        NotInHarvTable,
        MissingManualAcInProgram,
        MissingManualAcInTimeSettingsSheet,
        MissingManualAcInProgAndTsSheet,
        IllegalShmooValue,
        MultiUsedAcForPayload,
        MultiUsedTimesetForPayload,
        MultiUsedDcForPayload,
        MultiUsedLevelForPayload,
        MissingUsedAcForPayload,
        MissingUsedTimesetForPayload,
        MissingUsedDcForPayload,
        MissingUsedLevelForPayload,
        MissingDigSrcSgmtInPattern,
        VoltageHigherThan1P3,
        VddShmooLowToHigh,
        MissingPinMap,
        MissingPatInfo,
        IllegalForNewTChar,
        IllegalCharUslLsl,
    };

    public class ErrorMessage
    {
        public ErrorLevel ErrorLevel { get; set; }
        public ErrorType ErrorType { get; set; }
        public string SheetName { get; set; }
        public int RowNum { get; set; }
        public List<int> ColList = new List<int>();
        public string Message { get; set; }
        public List<string> CommentsList = new List<string>();
    }
}
