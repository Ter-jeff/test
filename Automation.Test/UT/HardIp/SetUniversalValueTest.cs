using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;
using Automation.InputManager.Data;
using Automation.Static;

using CommonReaderLib.PatternListCsv;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

using TestPlanLib.VbtLib;

namespace Automation.Test.UT.HardIp;

[TestClass]
public class SetUniversalValueTest : FunctionTestBase
{

    public static IEnumerable<object[]> PatSetAssignCases =>
    [
        // Multiple pattern
        [
            new HardIpPattern
            {
                Pattern = new("test_pattern_name")
                {
                    PatternSetList = [["pp_gg_aa", "pp_gg_bb"]],
                    InstancePayloadName = ["payload_name_1", "payload_name_2"],
                }
            },
            "payload_name_1,payload_name_2"
        ],
        // Single normal
        [
            new HardIpPattern
            {
                Pattern = new("test_pattern_name2")
                {
                    PatternSetList = [["pp_gg_aa"]],
                    InstancePatternName = ["inst_pat_name1", "inst_pat_name2"],
                }
            },
            "inst_pat_name1;inst_pat_name2"
        ],
        // Single no pattern
        [
            new HardIpPattern
            {
                Pattern = new("test_pattern_name2")
                {
                    PatternSetList = [["pp_gg_aa"]],
                    TestPlanPatternName = HardIpConstData.NoPattern,
                }
            },
            string.Empty,
        ],
        // Single instance
        [
            new HardIpPattern
            {
                Pattern = new("test_pattern_name3")
                {
                    PatternSetList = [["pp_gg_aa"]],
                    TestPlanPatternName = "Instance:123",
                }
            },
            string.Empty,
        ],
    ];

    public static IEnumerable<object[]> CpuFlagAssignCases =>
    [
        // False case
        [
            new HardIpPattern
            {
                Pattern = new("test_pattern_name"),
            },
            new HardIpInfo(),
            "false",
        ],
        // True case
        [
            new HardIpPattern
            {
                Pattern = new("NoPattern"),
            },
            new HardIpInfo() { CallSubrs = "A" },
            "true",
        ],
    ];

    public static IEnumerable<object[]> RegAssignAssignCases =>
    [
        [
            "single_pattern",
            new HardIpPattern
            {
                Pattern = new("test_pattern_name1")
                {
                    PatternSetList = [["pp_gg_aa"]],
                },
                RegisterAssignment = "reg1=nestsweep(10)",
                DigSrcEquation = "test",
            },
            new HardIpInfo() { SendBit = "test" },
            "reg1=sweep1",
            string.Empty,
            "test",
            "test",
        ],
        [
            "multiple_pattern",
            new HardIpPattern
            {
                Pattern = new("test_pattern_name2")
                {
                    PatternSetList = [["pp_gg_aa", "pp_gg_bb"]],
                    InstancePayloadName = ["payload_name_1", "payload_name_2"],
                },
                BurstPatterns =
                [
                    new()
                    {
                        RegisterAssignment = "reg1=nestsweep(10)",
                    },
                    new()
                    {
                        RegisterAssignment = "reg2=nestsweep(10)",
                    }
                ],
                DigSrcEquation = "test",
            },
            new HardIpInfo() { SendBit = "test", SendBitName = "AA_0", SendBitStr = "B1_1" },
            "reg1=sweep1;reg2=sweep1",
            "1|1",
            "AA_0|AA_0",
            "test|test",
        ],
    ];

    public static IEnumerable<object[]> SpecialCalcAssignCases =>
    [
        [
            "MeasVdiff",
            new HardIpPattern
            {
                SpecialMeasType = MeasType.MeasVdiff
            },
            "4",
            Times.Once(),
        ],
        [
            "MeasVocm",
            new HardIpPattern
            {
                SpecialMeasType = MeasType.MeasVocm
            },
            "9",
            Times.Once(),
        ],
        [
            "NoSpecial",
            new HardIpPattern
            {
                SpecialMeasType = "Ahh!"
            },
            "9",
            Times.Never(),
        ],
    ];

    public static IEnumerable<object[]> InterposeAssignCases =>
    [
        [
            "allInOne",
            new HardIpPattern
            {
                Pattern = new("test_pattern_name")
                {
                    PatternSetList = [["pp_gg_aa", "pp_gg_bb"]],
                    InstancePayloadName = ["payload_name_1", "payload_name_2"],
                },
                ForceConditionList =
                [
                    new()
                    {
                        ForcePins =
                        [
                            new() { PinName = "pin1", Type = ForceConditionType.Normal, ForceType = "in" },
                            new() { PinName = "pin2", Type = ForceConditionType.Normal, ForceType = "in", ForceJob = "Job1" },
                            new() { PinName = "pin3", Type = ForceConditionType.Others, ForceType = "in", ForceValue = "V1" },
                        ]
                    },
                ],
                MiscInfo = "USL:T1;LSL:T2",
                MeasPins =
                [
                    new()
                    {
                        MeasType = MeasType.MeasV,
                        PinName = "PIN_B",
                        SequenceIndex = 1,
                        ForceConditions =
                        [
                            new()
                            {
                                ForcePins =
                                [
                                    new()
                                    {
                                        PinName = "PIN_A",
                                        ForceType = "V",
                                        ForceValue = "1.2",
                                        ForceLabelVoltages = [],
                                        Type = ForceConditionType.Normal,
                                        IsRestore = true,
                                        ForceJob = "",
                                        ForceInterPostMeas = "test",
                                    }
                                ]
                            }
                        ]
                    },
                ],
                TestPlanSequences = [new(1, 1, 1)],
                InterposePostTest = "test",
            },
            "pin1:in:;pin2:in::Job1;pin3:V1;USL:T1;LSL:T2",
            "PIN_A:V:1.2",
            "test",
            "test",
        ],
    ];

    public static IEnumerable<object[]> MeasCAssignCases =>
    [
        [
            "test1",
            new HardIpPattern()
            {
                Pattern = new("test_pattern_name"),
                MeasPins =
                [
                    new()
                    {
                        SequenceIndex = 1,
                        PinName = "CP=1",
                        CusStr = "A",
                    },
                    new()
                    {
                        SequenceIndex = 2,
                        PinName = "FT=1",
                        CusStr = "B",
                    },
                    new()
                    {
                        SequenceIndex = 3,
                        PinName = "GG=1",
                        CusStr = "C"
                    },
                ],
            },
            new HardIpInfo()
            {
                MeasSeqStr = "1,2,3",
                SeqInfo =
                [
                    new() { SeqName = "SEQ1" },
                    new() { SeqName = "SEQ2" },
                    new() { SeqName = "SEQ3" },
                ]
            },
            "A+B+C",
            "",
            "",
            "SEQ1,SEQ2,SEQ3",
        ],
    ];

    public static IEnumerable<object[]> MiscAssignCases =>
    [
        [
            "test1",
            new HardIpPattern()
            {
                CalcEqn = "reg_assign",
                MiscInfo = $"{HardIpConstData.InstSpecialSetup}:Test;SelSramBlock:AA",
            },
            "reg_assign",
            "KeyNotFoundInSetting",
            "TRUE",
            "",
            "",
            "SelSramBlock:AA",
        ],
    ];

    public static IEnumerable<object[]> DigiCapAssignCases =>
    [
        [
            "test1",
            new HardIpPattern()
            {
                Pattern = new("test_pattern_1"),
            },
            new HardIpInfo
            {
                CapBitStr = "AB1_1",
                CapBit = 1234,
            },
            "JTAG_TDO",
            "1",
            "1234",
            "",
            "",
        ],
    ];
    [ClassInitialize]
    public static new void ClassInitialize(TestContext testContext)
    {
        //string hardIpInfoFile = Path.Combine(InputPath, "HardIp", "HardIP_PatInfo_Default.log");
        //List<HardIpInfo> patInfo = new PatInfoReader().ExtractHardIpInfos(hardIpInfoFile);
        //LocalSpecs.HardIpInfos = new HardIpInfos(patInfo);
    }

    [TestMethod]
    public void Constructor_Always_SetsExpectedReservedMiscInfoKeys()
    {
        HardIpSheet dummySheet = new();
        HardIpInputData inputData = new(new(EnumBlock.HardIp));

        SetUniversalValue setUniversalValue = new(inputData, dummySheet);

        List<string> expected =
        [
            "digCapDataCustomString",
            "SelSramBlock",
            "BypassAutorange",
            "RAK",
            "ConvertDAC",
            "ConvertDACParameter",
            "SignedBinary",
            "RAK",
            "BypassAutorange",
            "DDREmulate",
            "UnSignedGray",
            "SignedGray",
            "TwosComplement",
            "SweepSrc",
            "SweepSrcOneToOne",
            "SweepVoltage",
            "DcvsVoltageClamp",
            "DigCapMsbFirst",
            "DigSrcMsbFirst",
            "RunByModule",
        ];

        CollectionAssert.AreEqual(expected, setUniversalValue.ReservedMiscInfoKeys);
    }

    [TestMethod]
    [DynamicData(nameof(PatSetAssignCases), DynamicDataSourceType.Property)]
    public void TestPatSetAssigner(HardIpPattern hardIpPattern, string expectedResult)
    {
        Mock<Function> mockFunc = new();
        Function function = mockFunc.Object;

        HardIpInputData hardIpInputData = new(new(EnumBlock.HardIp));
        FunctionAssignContext ctx = new(hardIpPattern, "2.5V", hardIpInputData);

        PatSetAssigner assigner = new();
        assigner.Assign(function, ctx);
        mockFunc.Verify(f => f.SetParamValue("patternSet", expectedResult), Times.Once);
    }

    [TestMethod]
    [DynamicData(nameof(CpuFlagAssignCases), DynamicDataSourceType.Property)]
    public void TestCpuFlagAssigner(HardIpPattern hardIpPattern, HardIpInfo hardIpInfo, string expectedResult)
    {
        Mock<Function> mockFunc = new();
        Function function = mockFunc.Object;
        LocalSpecs.HardIpInfos.Clear();
        LocalSpecs.HardIpInfos.Add(hardIpPattern.Pattern.GetPatternName(), [hardIpInfo]);

        HardIpInputData hardIpInputData = new(new(EnumBlock.HardIp));
        FunctionAssignContext ctx = new(hardIpPattern, "2.5V", hardIpInputData);

        CpuFlagAssigner assigner = new();
        assigner.Assign(function, ctx);
        mockFunc.Verify(f => f.SetParamValue("patternCPUAFlag", expectedResult), Times.Once);
    }

    [TestMethod]
    [DynamicData(nameof(RegAssignAssignCases), DynamicDataSourceType.Property)]
    public void TestRegAssignAssigner(string testName, HardIpPattern hardIpPattern, HardIpInfo hardIpInfo, string assignment, string dataWidth, string equation, string sampleSize)
    {
        Mock<Function> mockFunc = new();
        Function function = mockFunc.Object;
        LocalSpecs.HardIpInfos.Clear();

        HardIpInputData hardIpInputData = new(new(EnumBlock.HardIp));
        foreach (string patName in hardIpPattern.Pattern.PatternSetList.SelectMany(p => p))
        {
            LocalSpecs.HardIpInfos.Add(patName, [hardIpInfo]);
        }
        FunctionAssignContext ctx = new(hardIpPattern, "2.5V", hardIpInputData);
        RegAssignAssigner assigner = new();
        assigner.Assign(function, ctx);
        mockFunc.Verify(f => f.SetParamValue("digSrcAssignment", assignment), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("digSrcDataWidth", dataWidth), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("digSrcEquation", equation), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("digSrcSampleSize", sampleSize), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("digSrcPin", "JTAG_TDI"), Times.Once);
    }

    [TestMethod]
    [DynamicData(nameof(SpecialCalcAssignCases), DynamicDataSourceType.Property)]
    public void TestSpecialCalcAssigner(string testName, HardIpPattern hardIpPattern, string expectedResult, Times times)
    {
        Mock<Function> mockFunc = new();
        Function function = mockFunc.Object;
        HardIpInputData hardIpInputData = new(new(EnumBlock.HardIp));
        FunctionAssignContext ctx = new(hardIpPattern, "2.5V", hardIpInputData);

        SpecialCalcAssigner assigner = new();
        assigner.Assign(function, ctx);
        mockFunc.Verify(f => f.SetParamValue("specialCalcSetting", expectedResult), times);
    }

    [TestMethod]
    [DynamicData(nameof(InterposeAssignCases), DynamicDataSourceType.Property)]
    public void TestInterposeAssigner(string testName, HardIpPattern hardIpPattern, string expectedPrePat, string expectedPreMeas, string expectedPostMeas, string expectedPostTest)
    {
        Mock<Function> mockFunc = new();
        Function function = mockFunc.Object;
        HardIpInputData hardIpInputData = new(new(EnumBlock.HardIp));
        FunctionAssignContext ctx = new(hardIpPattern, "2.5V", hardIpInputData);

        InterposeAssigner assigner = new();
        assigner.Assign(function, ctx);
        mockFunc.Verify(f => f.SetParamValue("interposePrePat", expectedPrePat), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("interposePreMeas", expectedPreMeas), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("interposePostMeas", expectedPostMeas), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("interposePostTest", expectedPostTest), Times.Once);
    }

    [TestMethod]
    [DynamicData(nameof(MeasCAssignCases), DynamicDataSourceType.Property)]
    public void TestMeasCAssigner(string testName, HardIpPattern hardIpPattern, HardIpInfo hardIpInfo, string expectedStoreName, string expectedForceV, string expectedForceI, string expectedSequence)
    {
        Mock<Function> mockFunc = new();
        Function function = mockFunc.Object;
        HardIpInputData hardIpInputData = new(new(EnumBlock.HardIp));
        LocalSpecs.HardIpInfos.Clear();
        foreach (string patName in hardIpPattern.Pattern.PatternSetList.SelectMany(p => p))
        {
            LocalSpecs.HardIpInfos.Add(patName, [hardIpInfo]);
        }
        FunctionAssignContext ctx = new(hardIpPattern, "2.5V", hardIpInputData);

        MeasCAssigner assigner = new();
        assigner.Assign(function, ctx);

        mockFunc.Verify(f => f.SetParamValue("measureStoreName", expectedStoreName), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("measureForceV", expectedForceV), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("measureForceI", expectedForceI), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("measureSequence", expectedSequence), Times.Once);
    }

    [TestMethod]
    [DynamicData(nameof(MiscAssignCases), DynamicDataSourceType.Property)]
    public void TestMiscAssigner(string testName, HardIpPattern hardIpPattern, string expectedcalcEquation, string expectedSpecialSetup, string expectedLimitFlag, string expectedIntName, string expectedMsbFlag, string expectedSrcDataCusStr)
    {
        Mock<Function> mockFunc = new();
        Function function = mockFunc.Object;
        HardIpInputData hardIpInputData = new(new(EnumBlock.HardIp));
        LocalSpecs.HardIpInfos.Clear();

        FunctionAssignContext ctx = new(hardIpPattern, "2.5V", hardIpInputData);

        MiscAssigner assigner = new();
        assigner.Assign(function, ctx);

        mockFunc.Verify(f => f.SetParamValue("calcEquation", expectedcalcEquation), Times.Once);
        mockFunc.Verify(f => f.SetParamValue(HardIpConstData.InstSpecialSetup, expectedSpecialSetup), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("singleLimitFlag", expectedLimitFlag), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("digSrcFlowForLoopIntegerName", expectedIntName), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("setupMSBFirstFlag", expectedMsbFlag), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("digSrcDataCustomString", expectedSrcDataCusStr), Times.Once);
    }

    [TestMethod]
    [DynamicData(nameof(DigiCapAssignCases), DynamicDataSourceType.Property)]
    public void TestDigiCapAssigner(string testName, HardIpPattern hardIpPattern, HardIpInfo hardIpInfo, string expectedPin, string expectedDataWidth, string expectedSampleSize, string expectedDataCusStr, string expectedMainCusStr)
    {
        Mock<Function> mockFunc = new();
        Function function = mockFunc.Object;
        HardIpInputData hardIpInputData = new(new(EnumBlock.HardIp));

        LocalSpecs.HardIpInfos.Clear();
        foreach (string patName in hardIpPattern.Pattern.PatternSetList.SelectMany(p => p))
        {
            LocalSpecs.HardIpInfos.Add(patName, [hardIpInfo]);
        }

        FunctionAssignContext ctx = new(hardIpPattern, "2.5V", hardIpInputData);

        DigiCapAssigner assigner = new();
        assigner.Assign(function, ctx);

        mockFunc.Verify(f => f.SetParamValue("digCapPin", expectedPin), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("digCapDataWidth", expectedDataWidth), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("digCapSampleSize", expectedSampleSize), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("digCapDataCustomString", expectedDataCusStr), Times.Once);
        mockFunc.Verify(f => f.SetParamValue("mainProgramCustomString", expectedMainCusStr), Times.Once);
    }
}
