namespace CommonLib.ErrorReport.Base
{
    public enum EnumErrorType
    {
        Missing,
        FormatError,
        FormatWarning,
        Warning,
        Business,
        MissingPinName,
        NwireStateDisable,
        NwireConfigMismatch,
        DuplicateProgramSheet
    };


    public enum OtpErrorType
    {
        DuplicateRegisterName,
        WrongCrcIdx,
        NotFoundCrc,
        WrongDefaultOrRealItems,
        NotSupportFormat,
        NotContinuslyIdx,
        DefaultValueOverflow,
        WrongCrcChecksum,
        WrongEccCheckVal,
    };

    public enum PmicErrorType
    {
        MissingPinName,
        MissingPattern,
        MissingRegister,
        DuplicateVbt,
    };

    public enum HardIpErrorType
    {
        //common
        Existential,
        FormatError,
        MissingNeededSheets,
        MissingHeader,
        UnrecognisedHeader,
        MissingHardIpSheet,
        DuplicateInstance,
        MissingHardIpCategory,
        //pin
        MissingHardIpDcPin,
        MissingPinName,
        MissingPatInfoPinInPinMap,
        //PatInfo
        DuplicatePatternInPatInfo,
        MissingSendBitOrSendBitStr,
        MissingCallSubroutine,
        WrongSendInformation,
        WrongMeasSequence,
        WrongDigSrcSignalName,
        MissingMeasureSequence,
        MismatchSequenceAndCallSubrsCnt,
        //Scgh
        DuplicatePayloadInScgh,
        //Pattern
        MissingPatternInPatList,
        MissingPatternInScgh,
        MissingPatternInTestPlan,
        MissingPatternInPatInfo,
        MisPatternForMeasument,
        PatternExistMeasument,
        WrongPatternName,
        //Force condition
        WrongForceCondition,
        RelayRestore,
        //Register
        WrongRegisterAssignment,
        DupBitName,
        DictionaryKey,
        MissingSweepSingleloopTable,
        //ADC calc equation limit error
        ADC_Convertor,
        CalcEqnCheck,
        //Misc Info
        MissingParameter,
        ManualItems,
        RepeatSubBlock,
        ExistedSubBlock,
        MissingSubBlock,
        SubBlockUsage,
        SubBlockNotFound,
        NoBurstSubBlock,
        WrongSweepStep,
        UnknownTrimMode,
        HIPeFuseInput,
        HIPeFuseCatename,
        HIPeFuseDSPWAVE,
        //Meas pin
        WrongTotalMeasCount,
        WrongMeasCountInPatInfo,
        WrongMeasPinInPatInfo,
        WrongMeasType,
        WrongMeasContent,
        WrongMeasC,
        MissingTestplanPin,
        MissingPinSeq,
        PartialCorePower,
        WrongDiffPinLevel,
        //Store name
        DuplicateStoreName,
        //Limit
        WrongLimit,
        WrongLimitValue,
        OppositeLimit,
        MisMatchLimitUnit,
        MissingLimitUnit,
        MissingLimit,
        ClearLimit,
        //Bin
        MissingBinNum,
        //GPIO
        MisMatchGpioPinGroup,
        //Selsram
        SelsramMappingTableError,
        SelsramDigSrcAssignmentNotDefineInTable,
        SelsramDigSrcAssignmentNotDefineInInstance,
        CanNotGetSelsramSetting,
        //Others
        ConflictTName,
        MisVbtModule,
        MissingCalcFunction,
        MisCalculationParaDefine,
        MisStoreNameInPatParaDefine,
        IgxlLimitation,
        DcDefault,
        ClockInstanceNotUsed,
        ClockCheckInfoNotGenerate,
        WrongTimeSet,
        MissingBlockName,
        CountNotMatch,
        MissingInstance,
        MissingBintable,
        FailFlagNotFound,
        UpdateFail,
        FuseWriteJudgeFlag,
        BurstPatJudgeFlag,
        RegAssignError,
        RegAssignWarning,
        MissingFieldName,
        MismatchBankField,
        MismatchFieldJob,
        MismatchFieldWidth,
        MissingRealFieldNameInEfuseHardIP,
        MismatchRealValue,
        MissingInPreWrite,
        FT3Enabled
    };

    public enum EFuseErrorType
    {
        OfBits,
        Limit,
        BaseVoltage,
        Bit,
        crc,
        CMP,
        DefaultValue,
        MaximumBits,
        RowInfo,
        FuseBlowLocation,
        BankUnknown,
        Bank,
        ProgrammingStage,
        Existential,
        Pattern,
        UnusedPattern,
        PatternTestMode,
        FormatError,
        Business,
        EcidNonCpProgrammingStage,
        JudgeDataRevisionByTool,
        LsbMsb,
        NotMatchToBDF,
        MisVbtModule,
        MissingParameter,
        ShadowProgrammingStage,
        MisMatch,
        MisMatchDefaultType
    };

    public enum BasicErrorType
    {
        Existential,
        FormatError,
        FormatWarning,
        Warning,
        Business,
        FSDDIssue,
        MissingPinName,
        NwireStateDisable,
        NwireConfigMismatch,
        NwireOutClkVolWarning,
        MisVbtModule,
        MissingParameter,
        ContiIfold
    };

    public enum PreActionErrorType
    {
        Existential,
        FormatError,
        Business,
        DuplicateVbtModule,
        PinGroupNotMatch,
        DifferentialPair,
        MisVbtModule,
        MissingParameter
    };

    public enum MbistErrorType
    {
        ReferenceFileError,
        Existential,
        FormatError,
        Business,
        MisVbtModule,
        MissingParameter
    };

    public enum HarvestErrorType
    {
        HeaderError,
        OrderError,
        FieldsError,
        FormatError,
        MisVbtModule,
        MissingParameter,
        MissingPatternPinGroup
    }

    public enum ScanErrorType
    {
        Pattern,
        TimeSet,
        Existential,
        FormatError,
        Business,
        NonUsedApplication,
        PatternTimeSet,
        PatternSetName,
        VoltageType,
        MisVbtModule,
        MissingParameter,
        InstanceName,
        UserFunction
    }

    public enum EvsErrorType
    {
        Existential,
        FormatError,
        Business,
        Pattern,
        MisVbtModule,
        MissingParameter
    }

    public enum RtosErrorType
    {
        Existential,
        FormatError,
        Business,
        MisVbtModule,
        MissingParameter
    }

    public enum MainFlowErrorType
    {
        Existential,
        FormatError,
        Business,
        FlowSheet,
        InstanceSheet,
        MissingParameter
    }

    public enum RelayErrorType
    {
        Existential,
        FormatError,
        Business
    };

    public enum BinCutErrorType
    {
        Existential,
        Pattern,
        FormatError,
        ID,
        Equation,
        C,
        M,
        TestingStage,
        RowDuplicate,
        FlowName,
        Missing,
        TTR,
        PatternTimeSet,
        BistInitPattern,
        PatternSetName,
        VoltageType,
        TimeSet,
        Business,
        IdsMax,
        CpMax,
        CpMin,
        Inherit,
        efuseBitDef,
        Formula,
        Resolution,
        Cpgb,
        Cp2Gb,
        FtRoomGb,
        FtHotGb,
        HTOLGb,
        AllowEqual,
        Domain,
        PerformanceMode,
        Interpolation,
        SRAMthresh,
        Mode_EQN,
        Flow,
        JobMisMatch,
        MisVbtModule,
        MissingParameter,
        InstanceName,
        UserFunction
    };

    public enum HtolErrorType
    {
        MisVbtModule
    };

    public enum DuplicateInstance
    {
        Duplicate
    };

    public enum DuplicateTestNumber
    {
        Duplicate
    };

    public enum UnusedDcCategory
    {
        UnusedDcCategory
    };

    public enum ZeroVoltageDcCategory
    {
        ZeroVoltageDcCategory
    };

    public enum PatValtPinCheckerType
    {
        ValueMissingInCategory
    };

    public enum PatInfoType
    {
        //PatInfo
        DuplicatePatternInPatInfo,
        MissingSendBitOrSendBitStr,
        MissingCallSubroutine,
        WrongSendInformation,
        WrongMeasSequence,
        WrongDigSrcSignalName,
        MissingMeasureSequence,
        MismatchSequenceAndCallSubrsCnt,
        MismatchVmandSubr,
        MissingVmvector,
        MismatchGenericName
    }

    public enum SsnErrorType
    {
        MissingSsnInfo,
        WrongSsnInfo,
        CoreNameMismatch,
        DuplicatePinGroup,
        DuplicatePattern,
        IllegalFlagName,
        PinGroupAndFlagMismatch
    }

    public enum FlowMainErrorType
    {
        MissingMappingTarget
    }

    public enum PatternMissing
    {
        MissingPatternFile
    };

    public enum UFInstanceErrorType
    {
        DefaultPowerUpInstanceExist
    }
}
