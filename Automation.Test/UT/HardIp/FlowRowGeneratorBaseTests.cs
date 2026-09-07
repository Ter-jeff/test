using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class FlowRowGeneratorBaseTests : FunctionTestBase
    {
        public class FlowRowGeneratorBaseTestStub(HardIpInputData hardIpInputData, string sheetName) : HardIpFlowRowGenerator(hardIpInputData, sheetName)
        {
            public string CallCreateTestOpcode()
            {
                return CreateTestOpcode();
            }
        }

        private static HardIpFlowRowGenerator _generator = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            HardIpParaData paraData = new HardIpParaData(EnumBlock.HardIp);
            HardIpInputData hardIpInputData = new HardIpInputData(paraData);
            _generator = new HardIpFlowRowGenerator(hardIpInputData, "sheetName");
            HardIpPattern pattern = new HardIpPattern
            {
                SweepCodes = new Dictionary<string, List<SweepCode>>
                {
                    {"A",new List<SweepCode> {new() { SendBitName = "SendBitName"} }}
                },
            };
            _generator.Pat = pattern;
        }

        [TestMethod]
        public void NeedRealtimeBinOut_FromMisc()
        {
            _generator.Pat.MiscInfoDict = new Dictionary<string, string>
            {
                { HardIpConstData.RealtimePatBinOut, "1" }
            };

            bool result = _generator.NeedRealtimeBinOut;

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void NoNeedToGen_BlockTypeAndLV_ShouldReturnTrue()
        {
            _generator.Pat.BlockType = "BLOCK";
            _generator.LabelVoltage = "LV";

            bool result = _generator.NoNeedToGen;

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void GenHardIpEnableDisableName_ShouldReturnExpectedFormat()
        {
            // Act
            string result = _generator.GenHardIpEnableDisableName("MyHeader");

            // Assert
            Assert.AreEqual("MyHeader_HardIP_Datalog_Format", result);
        }

        [TestMethod]
        public void GenNwireSetting_NoTimingMatch_ShouldUseKeepDefault()
        {
            SubFlowSheet sheet = new SubFlowSheet("S");

            List<Timing> timings =
            [
                new() { Name = "Other", SuffixAcSpecName = "SpecX" }
            ];

            NwireSingleton.Instance().SettingInfo = new NwireSetting
            {
                NwirePins =
                [
                    new() { OutClk = "PinA" }
                ]
            };

            List<FlowRow> result = _generator.GenNwireSetting(sheet, timings);

            Assert.IsTrue(result[1].Parameter.Contains("keepDefault"));
        }

        [TestMethod]
        public void GenNwireSettingByMisc_ShouldGenerateExpectedRows()
        {
            // Arrange
            var flowSheet = new SubFlowSheet("TestSheet");
            string miscInfo = "FreeRunClk_disable:CLK1;FreeRunClk_enable:CLK2;Other:ABC";
            string enable = "ENABLE_TEST";

            // Act
            List<FlowRow> result = _generator.GenNwireSettingByMisc(flowSheet, miscInfo, enable);

            // Assert
            Assert.AreEqual(2, result.Count, "Expected two FlowRows to be created.");

            FlowRow first = result[0];
            Assert.AreEqual(OpCode.Test, first.Opcode);
            Assert.AreEqual("ENABLE_TEST", first.Enable);
            Assert.AreEqual("FreeRunClk_Disable_CLK1", first.Parameter);

            FlowRow second = result[1];
            Assert.AreEqual(OpCode.Test, second.Opcode);
            Assert.AreEqual("ENABLE_TEST", second.Enable);
            Assert.AreEqual("FreeRunClk_Enable_CLK2", second.Parameter);

            Assert.AreEqual(2, flowSheet.Rows.Count);
        }

        [TestMethod]
        public void GenNwireSettingByMisc_ShouldReturnEmpty_WhenNoFreeRunClkFound()
        {
            // Arrange
            var flowSheet = new SubFlowSheet("EmptySheet");
            string miscInfo = "SomethingElse:123";

            // Act
            List<FlowRow> result = _generator.GenNwireSettingByMisc(flowSheet, miscInfo, "EN");

            // Assert
            Assert.AreEqual(0, result.Count);
            Assert.AreEqual(0, flowSheet.Rows.Count);
        }

        [TestMethod]
        public void GenNwireSettingByMisc_ShouldHandleNullFlowSheet()
        {
            // Arrange
            string miscInfo = "FreeRunClk_enable:CLK3";

            // Act
            List<FlowRow> result = _generator.GenNwireSettingByMisc(null, miscInfo, "ENABLED");

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("FreeRunClk_Enable_CLK3", result[0].Parameter);
        }

        [TestMethod]
        public void GenSweepCodeForRowTest()
        {
            // Arrange
            _generator.Pat.Pattern.PatternSetList =
            [
                ["P1"] // single → else branch
            ];

            _generator.Pat.SweepCodes = new Dictionary<string, List<SweepCode>>
            {
                {
                    "A",
                    new List<SweepCode>
                    {
                        new() { SweepInfo = "0,0,0" }
                    }
                }
            };

            // Act
            List<FlowRow>? result = _generator.GenSweepCodeForRow();

            // Assert
            Assert.AreEqual(1, result!.Count);
            Assert.AreEqual("SrcCodeIndxA = 0; SrcCodeIndxA <= 0; SrcCodeIndxA+=0", result[0].Parameter);
        }

        [TestMethod]
        public void GenSweepCodeForRowTest1()
        {
            // Act
            _generator.Pat.Pattern.PatternSetList = [["P1", "P2"]];
            _generator.Pat.BurstPatterns =
            [
                new()
                {
                    Pattern = new PatternClass("P1") { RealPatternName = "P1" },
                    RegisterAssignment = "Reg_assign_A",
                    SweepCodes = new Dictionary<string, List<SweepCode>>
                    {
                        { "A", new List<SweepCode> { new() { SendBitName = "sweepCode" } } }
                    }
                },
                new()
                {
                    Pattern = new PatternClass("P2") { RealPatternName = "P2" },
                    RegisterAssignment = "Reg_assign_B"
                }
            ];

            List<FlowRow>? result = _generator.GenSweepCodeForRow();

            // Assert
            Assert.AreEqual(1, result!.Count);
            Assert.AreEqual("SrcCodeIndxA = 0; SrcCodeIndxA <= 0; SrcCodeIndxA+=0", result[0].Parameter);
        }

        [TestMethod]
        public void GenSweepCodeForRow_NoSweepCode_ReturnNull()
        {
            // Arrange
            _generator.Pat.SweepCodes = [];
            _generator.Pat.BurstPatterns =
            [
                new()
                {
                    SweepCodes = [] // empty
                }
            ];

            // Act
            List<FlowRow>? result = _generator.GenSweepCodeForRow();

            // Assert
            Assert.AreEqual(null, result);
        }

        [TestMethod]
        public void GenSweepCodeForRow_HasSweepCodeFromBurstPattern()
        {
            // Arrange
            _generator.Pat.Pattern.PatternSetList =
            [
                ["P1", "P2"] // ✅ 修正這裡
            ];

            _generator.Pat.SweepCodes = [];

            _generator.Pat.BurstPatterns =
            [
                new()
                {
                    Pattern = new PatternClass("P1") { RealPatternName = "P1" },
                    SweepCodes = new Dictionary<string, List<SweepCode>>
                    {
                        {
                            "A",
                            new List<SweepCode>
                            {
                                new()
                                {
                                    SendBitName = "test",
                                    SweepInfo = "0,0,0"
                                }
                            }
                        }
                    }
                }
            ];

            // Act
            List<FlowRow>? result = _generator.GenSweepCodeForRow();

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(1, result!.Count);
        }

        [TestMethod]
        public void GenSweepCodeForRow_DuplicateKey_OnlyOneFlowRow()
        {
            // Arrange
            _generator.Pat.Pattern.PatternSetList =
            [
                ["P1", "P2"]
            ];

            _generator.Pat.BurstPatterns =
            [
                new()
                {
                    Pattern = new PatternClass("P1") { RealPatternName = "P1" },
                    SweepCodes = new Dictionary<string, List<SweepCode>>
                    {
                        {
                            "A",
                            new List<SweepCode>
                            {
                                new() { SweepInfo = "0,0,0" }
                            }
                        }
                    }
                },
                new()
                {
                    Pattern = new PatternClass("P2") { RealPatternName = "P2" },
                    SweepCodes = new Dictionary<string, List<SweepCode>>
                    {
                        {
                            "A", // same key
                            new List<SweepCode>
                            {
                                new() { SweepInfo = "1,1,1" }
                            }
                        }
                    }
                }
            ];

            // Act
            List<FlowRow>? result = _generator.GenSweepCodeForRow();

            // Assert
            Assert.AreEqual(1, result!.Count);
        }

        [TestMethod]
        public void GenSweepCodeForRow_DifferentKey_GenerateMultipleFlowRows()
        {
            // Arrange
            _generator.Pat.Pattern.PatternSetList =
            [
                ["P1", "P2"]
            ];

            _generator.Pat.BurstPatterns =
            [
                new()
                {
                    Pattern = new PatternClass("P1") { RealPatternName = "P1" },
                    SweepCodes = new Dictionary<string, List<SweepCode>>
                    {
                        {
                            "A",
                            new List<SweepCode>
                            {
                                new() { SweepInfo = "0,0,0" }
                            }
                        }
                    }
                },
                new()
                {
                    Pattern = new PatternClass("P2") { RealPatternName = "P2" },
                    SweepCodes = new Dictionary<string, List<SweepCode>>
                    {
                        {
                            "B",
                            new List<SweepCode>
                            {
                                new() { SweepInfo = "0,0,0" }
                            }
                        }
                    }
                }
            ];

            // Act
            List<FlowRow>? result = _generator.GenSweepCodeForRow();

            // Assert
            Assert.AreEqual(2, result!.Count);
        }

        [TestMethod]
        public void GenSweepCodeForRow_CustomType_ShouldUseCustomLoop()
        {
            // Arrange
            _generator.Pat.Pattern.PatternSetList =
            [
                ["P1", "P2"] // ✅ ensure IsMultiple = true
            ];

            _generator.Pat.SweepCodes = [];

            _generator.Pat.BurstPatterns =
            [
                new()
                {
                    Pattern = new PatternClass("P1") { RealPatternName = "P1" },
                    SweepCodes = new Dictionary<string, List<SweepCode>>
                    {
                        {
                            "A",
                            new List<SweepCode>
                            {
                                new()
                                {
                                    Type = SweepCode.SweepType.Custom,
                                    SweepInfo = "1,2,3"
                                }
                            }
                        }
                    }
                }
            ];

            // Act
            List<FlowRow>? result = _generator.GenSweepCodeForRow();

            // Assert
            Assert.AreEqual(1, result!.Count);
            Assert.AreEqual("SrcCodeIndxA = 0; SrcCodeIndxA < 3; SrcCodeIndxA++", result[0].Parameter);
        }

        [TestMethod]
        public void SortPatFlowRowsTest()
        {
            // Arrange
            var originFlowRows = new List<FlowRow>
                {
                    new()
                    {
                        Opcode = "Test",
                        Parameter = "Parameter",
                        Comment = "Calc"
                    },
                    new()
                    {
                        Opcode = "Test",
                        Parameter = "Parameter1"
                    }

                };

            LocalSpecs.Options.Device = EnumDevice.AP;

            // Act
            List<FlowRow> result = _generator.SortPatFlowRows(originFlowRows);

            // Assert
            Assert.AreEqual("Parameter1", result[0].Parameter);
            Assert.AreEqual("Parameter", result[1].Parameter);
        }

        [TestMethod]
        public void GenEnable_WithMiscEnableWord()
        {
            _generator.Pat.Pattern = new PatternClass("PP_TEST");
            _generator.Pat.MiscInfo = "Enableword:CZ,DRAM";

            string result = _generator.GenEnable("BASE");

            Assert.IsTrue(result.Contains("CZ"));
            Assert.IsTrue(result.Contains("DRAM"));
        }

        [TestMethod]
        public void GenEnable_WithVoltageFiltering()
        {
            _generator.Pat.MiscInfo = "EnableWord:CZ@NV,HV";
            _generator.LabelVoltage = "NV";

            string result = _generator.GenEnable("BASE");

            Assert.IsTrue(result.Contains("CZ"));
        }

        [TestMethod]
        public void GenEnable_CzPattern_ShouldReturnOnlyBase()
        {
            _generator.Pat.Pattern = new PatternClass("CZ_TEST");

            string result = _generator.GenEnable("BASE");

            Assert.AreEqual("BASE", result);
        }

        [TestMethod]
        public void GenTtrFlagClearRow_ShouldGenerateDistinctFlags()
        {
            List<FlowRow> rows =
            [
                new() { Env = HardIpConstData.EnvTtr, FailAction = "F1" },
                new() { Env = HardIpConstData.EnvTtr, FailAction = "F1" },
                new() { Env = HardIpConstData.EnvTtr, FailAction = "F2" }
            ];

            List<FlowRow> result = _generator.GenTtrFlagClearRow(rows);

            Assert.AreEqual(2, result.Count);
        }

        [TestMethod]
        public void SortFlowRows_DeviceAP()
        {
            // Arrange
            var originFlowRows = new List<FlowRow>
                {
                    new()
                    {
                        Opcode = "Test",
                        Parameter = "Parameter",
                        Comment = "Calc"
                    },
                    new()
                    {
                        Opcode = "Test",
                        Parameter = "Parameter1"
                    }

                };

            LocalSpecs.Options.Device = EnumDevice.AP;

            // Act
            List<FlowRow> result = _generator.SortFlowRows(originFlowRows);

            // Assert
            Assert.AreEqual("Parameter1", result[0].Parameter);
            Assert.AreEqual("Parameter", result[1].Parameter);
        }

        [TestMethod]
        public void SortFlowRows_IsFullMTDPatternVbt()
        {
            // Arrange
            var originFlowRows = new List<FlowRow>
            {
                new()
                {
                    Opcode = "Test",
                    Parameter = "Parameter0",
                    PatName = "A",
                    PatIndex = 0,
                    Comment = "MeasV"
                },
                new()
                {
                    Opcode = "Test",
                    Parameter = "Parameter3",
                    PatName = "B",
                    PatIndex = 1,
                    Comment = "MeasI"
                },
                new()
                {
                    Opcode = "Test",
                    Parameter = "Parameter2",
                    PatName = "B",
                    PatIndex = 1,
                    Comment = "Calc"
                },
                new()
                {
                    Opcode = "Test",
                    Parameter = "Parameter1",
                    PatName = "A",
                    PatIndex = 0,
                    Comment = "Calc"
                }
            };
            _generator.Pat.Pattern.RealPatternName = "A#B";
            _generator.Pat.Pattern.PatternSetList =
            [
                ["A"],
                ["B"]
            ];
            _generator.Pat.Pattern.InstancePatternName = ["A", "B"];
            _generator.Pat.FunctionName = "hardip_mtd_test";

            LocalSpecs.Options.Device = EnumDevice.AP;

            // Act
            List<FlowRow> result = _generator.SortFlowRows(originFlowRows);

            // Assert
            Assert.AreEqual("Parameter0", result[0].Parameter);
            Assert.AreEqual("Parameter3", result[1].Parameter);
            Assert.AreEqual("Parameter1", result[2].Parameter);
            Assert.AreEqual("Parameter2", result[3].Parameter);

            _generator.Pat.Pattern.RealPatternName = "";
            _generator.Pat.Pattern.PatternSetList = [];
        }

        [TestMethod]
        public void SortFlowRows_IsFullMTDPatternCs()
        {
            // Arrange
            var originFlowRows = new List<FlowRow>
            {
                new()
                {
                    Opcode = "Test",
                    Parameter = "Parameter0",
                    PatName = "A",
                    Comment = "MeasV"
                },
                new()
                {
                    Opcode = "Test",
                    Parameter = "Parameter3",
                    PatName = "B",
                    Comment = "MeasI"
                },
                new()
                {
                    Opcode = "Test",
                    Parameter = "Parameter2",
                    PatName = "B",
                    Comment = "Calc"
                },
                new()
                {
                    Opcode = "Test",
                    Parameter = "Parameter1",
                    PatName = "A",
                    Comment = "Calc"
                }
            };
            _generator.Pat.Pattern.RealPatternName = "A#B";
            _generator.Pat.Pattern.PatternSetList =
            [
                ["A"],
                ["B"]
            ];
            _generator.Pat.Pattern.InstancePatternName = ["A", "B"];
            _generator.Pat.FunctionName = "HardIPMultiTimeDomain";

            LocalSpecs.Options.Device = EnumDevice.AP;

            // Act
            List<FlowRow> result = _generator.SortFlowRows(originFlowRows);

            // Assert
            Assert.AreEqual("Parameter0", result[0].Parameter);
            Assert.AreEqual("Parameter1", result[1].Parameter);
            Assert.AreEqual("Parameter3", result[2].Parameter);
            Assert.AreEqual("Parameter2", result[3].Parameter);

            _generator.Pat.Pattern.RealPatternName = "";
            _generator.Pat.Pattern.PatternSetList = [];
        }

        [TestMethod]
        public void SortFlowRows_DeviceRF()
        {
            // Arrange
            var originFlowRows = new List<FlowRow>
            {
                new()
                {
                    Opcode = "Test",
                    Parameter = "Parameter",
                    Comment = "Calc"
                },
                new()
                {
                    Opcode = "Test",
                    Parameter = "Parameter1"
                }
            };

            LocalSpecs.Options.Device = EnumDevice.RF;

            // Act
            List<FlowRow> result = _generator.SortFlowRows(originFlowRows);

            // Assert
            Assert.AreEqual("Parameter1", result[0].Parameter);
        }

        [TestMethod]
        public void SortFlowRows_DeviceNone_ShouldSortByCommentRule()
        {
            // Arrange
            var originFlowRows = new List<FlowRow>
    {
        new()
        {
            Opcode = "Test",
            Parameter = "P0",
            Comment = "Calc"
        },
        new()
        {
            Opcode = "Test",
            Parameter = "P1",
            Comment = "MeasC"
        },
        new()
        {
            Opcode = "Test",
            Parameter = "P2"
        }
    };

            LocalSpecs.Options.Device = EnumDevice.None;

            // Act
            List<FlowRow> result = _generator.SortFlowRows(originFlowRows);

            // Assert
            Assert.AreEqual("P2", result[0].Parameter);
            // MeasC
            Assert.AreEqual("P1", result[1].Parameter);
            // Calc
            Assert.AreEqual("P0", result[2].Parameter);
        }

        [TestMethod]
        public void SortFlowRows_DeviceLCD_ShouldKeepOriginalOrder()
        {
            // Arrange
            var originFlowRows = new List<FlowRow>
    {
        new()
        {
            Opcode = "Test",
            Parameter = "LCD_0",
            Comment = "Calc"
        },
        new()
        {
            Opcode = "Test",
            Parameter = "LCD_1"
        }
    };

            LocalSpecs.Options.Device = EnumDevice.LCD;

            // Act
            List<FlowRow> result = _generator.SortFlowRows(originFlowRows);

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("LCD_0", result[0].Parameter);
            Assert.AreEqual("LCD_1", result[1].Parameter);
        }

        [TestMethod]
        public void SortFlowRows_DeviceRF_WithNoSpecialComment_ShouldKeepOriginalOrder()
        {
            // Arrange
            var originFlowRows = new List<FlowRow>
    {
        new()
        {
            Opcode = "Test",
            Parameter = "RF_0"
        },
        new()
        {
            Opcode = "Test",
            Parameter = "RF_1"
        }
    };

            LocalSpecs.Options.Device = EnumDevice.RF;

            // Act
            List<FlowRow> result = _generator.SortFlowRows(originFlowRows);

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("RF_0", result[0].Parameter);
            Assert.AreEqual("RF_1", result[1].Parameter);
        }

        [TestMethod]
        public void GenNwireSettingTest()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("");
            var timings = new List<Timing>
            {
                new() { Name = "PinA", SuffixAcSpecName = "Spec1" }
            };

            NwireSingleton.Instance().SettingInfo = new NwireSetting
            {
                NwirePins =
                [
                    new() { OutClk = "PinA" }
                ]
            };

            string enable = "Yes";
            string job = "TestJob";

            // Act
            List<FlowRow> result = _generator.GenNwireSetting(subFlowSheet, timings, enable, job);

            // Assert
            Assert.AreEqual(2, result.Count, "Should generate one row for default setting and one for timings.");

            FlowRow defaultRow = result[0];
            Assert.AreEqual(OpCode.Test, defaultRow.Opcode);
            Assert.IsTrue(defaultRow.Parameter.Contains("FreeRunClk_Disable_PinA"));
            Assert.AreEqual(enable, defaultRow.Enable);

            FlowRow timingRow = result[1];
            Assert.IsTrue(timingRow.Parameter.Contains("Spec1"), "The parameter should include the suffix from matched timing.");

        }

        [TestMethod]
        public void CreateTestFailActionByTestPlanDefineTest()
        {
            _generator.Pat.Failflag = "NV@F_GEDR_T1NOBUF_GIPSCAN_UserDefine_NV;\nLV@F_GEDR_T1NOBUF_GIPSCAN_UserDefine_LV;\nHV@F_GEDR_T1NOBUF_GIPSCAN_UserDefine_HV;";

            _generator.LabelVoltage = "LV";
            string resultLV = _generator.CreateTestFailActionByTestPlanDefine();
            Assert.AreEqual("F_GEDR_T1NOBUF_GIPSCAN_UserDefine_LV", resultLV);

            _generator.LabelVoltage = "HV";
            string resultHV = _generator.CreateTestFailActionByTestPlanDefine();
            Assert.AreEqual("F_GEDR_T1NOBUF_GIPSCAN_UserDefine_HV", resultHV);

            _generator.LabelVoltage = "NV";
            string resultNV = _generator.CreateTestFailActionByTestPlanDefine();
            Assert.AreEqual("F_GEDR_T1NOBUF_GIPSCAN_UserDefine_NV", resultNV);
        }

        [TestMethod]
        public void CreateIfConditionByTestPanDefineTest()
        {
            _generator.Pat.SiteFlag = "NV@F_GEDR_T1NOBUF_GIPSCAN_Judge_NV||F_GEDR_T1NOBUF_GIPSCAN_Judge2_NV;\nLV@!(F_GEDR_T1NOBUF_GIPSCAN_Judge_LV||F_GEDR_T1NOBUF_GIPSCAN_Judge2_LV);\nHV@F_GEDR_T1NOBUF_GIPSCAN_Judge_HV;";

            _generator.LabelVoltage = "LV";
            string resultLV = _generator.CreateIfConditionByTestPanDefine();
            Assert.AreEqual("!(F_GEDR_T1NOBUF_GIPSCAN_Judge_LV||F_GEDR_T1NOBUF_GIPSCAN_Judge2_LV)", resultLV);

            _generator.LabelVoltage = "HV";
            string resultHV = _generator.CreateIfConditionByTestPanDefine();
            Assert.AreEqual("F_GEDR_T1NOBUF_GIPSCAN_Judge_HV", resultHV);

            _generator.LabelVoltage = "NV";
            string resultNV = _generator.CreateIfConditionByTestPanDefine();
            Assert.AreEqual("F_GEDR_T1NOBUF_GIPSCAN_Judge_NV||F_GEDR_T1NOBUF_GIPSCAN_Judge2_NV", resultNV);
        }

        [TestMethod]
        public void GetIfFlowRowsTest_Multiflags()
        {
            FlowRow answerIf = new FlowRow() { Opcode = OpCode.If, Parameter = "(F_Multi_1&&F_Multi_2) || F_Debug_all" };
            FlowRow answerTest1 = new FlowRow() { Opcode = OpCode.Test, Parameter = "Test1" };
            FlowRow answerIfTest1 = new FlowRow() { Opcode = OpCode.Test, Parameter = "IfTest1" };
            FlowRow answerIfTest2 = new FlowRow() { Opcode = OpCode.Test, Parameter = "IfTest2" };
            FlowRow answerEndIf = new FlowRow() { Opcode = OpCode.EndIf };
            List<FlowRow> answerIfTests = [answerIf, answerTest1, answerIfTest1, answerIfTest2, answerEndIf];

            string flags = "F_Multi_1&&F_Multi_2";
            FlowRow test = new FlowRow() { Opcode = OpCode.Test, Parameter = "Test1" };
            FlowRow ifTest1 = new FlowRow() { Opcode = OpCode.Test, Parameter = "IfTest1" };
            FlowRow ifTest2 = new FlowRow() { Opcode = OpCode.Test, Parameter = "IfTest2" };
            List<FlowRow> ifTests = [ifTest1, ifTest2];
            List<FlowRow> combinedFlowRows = _generator.GetIfFlowRows(flags, test, ifTests);

            Assert.AreEqual(answerIfTests.Count, combinedFlowRows.Count);
            for (int i = 0; i < answerIfTests.Count; i++)
            {
                Assert.AreEqual(answerIfTests[i].Opcode, combinedFlowRows[i].Opcode);
                Assert.AreEqual(answerIfTests[i].Parameter, combinedFlowRows[i].Parameter);
            }
        }

        [TestMethod]
        public void GetIfFlowRowsTest_Singleflag()
        {
            FlowRow answerTest1 = new FlowRow() { Opcode = OpCode.Test, Parameter = "Test1", DeviceName = "F_Condition", DeviceCondition = "Flag-true" };
            FlowRow answerIfTest1 = new FlowRow() { Opcode = OpCode.Test, Parameter = "IfTest1", DeviceName = "F_Condition", DeviceCondition = "Flag-true" };
            FlowRow answerIfTest2 = new FlowRow() { Opcode = OpCode.Test, Parameter = "IfTest2", DeviceName = "F_Condition", DeviceCondition = "Flag-true" };
            List<FlowRow> answerIfTests = [answerTest1, answerIfTest1, answerIfTest2];

            string flags = "F_Condition";
            FlowRow test = new FlowRow() { Opcode = OpCode.Test, Parameter = "Test1" };
            FlowRow ifTest1 = new FlowRow() { Opcode = OpCode.Test, Parameter = "IfTest1" };
            FlowRow ifTest2 = new FlowRow() { Opcode = OpCode.Test, Parameter = "IfTest2" };
            List<FlowRow> ifTests = [ifTest1, ifTest2];
            List<FlowRow> combinedFlowRows = _generator.GetIfFlowRows(flags, test, ifTests);

            Assert.AreEqual(answerIfTests.Count, combinedFlowRows.Count);
            for (int i = 0; i < answerIfTests.Count; i++)
            {
                Assert.AreEqual(answerIfTests[i].Opcode, combinedFlowRows[i].Opcode);
                Assert.AreEqual(answerIfTests[i].Parameter, combinedFlowRows[i].Parameter);
                Assert.AreEqual(answerIfTests[i].DeviceName, combinedFlowRows[i].DeviceName);
                Assert.AreEqual(answerIfTests[i].DeviceCondition, combinedFlowRows[i].DeviceCondition);
            }
        }

        [TestMethod]
        public void GenSweepVoltageForRow_IsNeedConvert_CountPath()
        {
            _generator.Pat.SweepVoltage = new Dictionary<string, List<SweepVData>>
            {
                {
                    "X",
                    new List<SweepVData>
                    {
                        new("X,0,10,1", isConvert: true)
                    }
                }
            };
            _generator.Pat.FlowControlFlag = 0;
            _generator.Pat.IsFlowInsRepeat = false;

            List<FlowRow>? result = _generator.GenSweepVoltageForRow();

            Assert.AreNotEqual(null, result);
            Assert.AreEqual(1, result!.Count);
            Assert.AreEqual("for", result[0].Opcode);
            Assert.AreEqual("SrcCodeIndexX = 0; SrcCodeIndexX <= 10; SrcCodeIndexX++", result[0].Parameter);
        }

        [TestMethod]
        public void GenSweepVoltageForRow_IsNeedConvert_ExceptionPath()
        {
            _generator.Pat.SweepVoltage = new Dictionary<string, List<SweepVData>>
            {
                {
                    "Y",
                    new List<SweepVData>
                    {
                        new("PINY,A,B,1", isConvert: true)
                    }
                }
            };
            _generator.Pat.FlowControlFlag = 0;
            _generator.Pat.IsFlowInsRepeat = false;

            List<FlowRow>? result = _generator.GenSweepVoltageForRow();

            Assert.AreNotEqual(null, result);
            Assert.AreEqual(1, result!.Count);
            Assert.AreEqual("PINY = A; PINY <= B; PINY+=1", result[0].Parameter);
        }

        [TestMethod]
        public void GenSweepVoltageForRow_NoConvert_Path()
        {
            _generator.Pat.SweepVoltage = new Dictionary<string, List<SweepVData>>
            {
                {
                    "Z",
                    new List<SweepVData>
                    {
                        new("Z,1,5,1", isConvert: false)
                    }
                }
            };
            _generator.Pat.FlowControlFlag = 0;
            _generator.Pat.IsFlowInsRepeat = false;

            List<FlowRow>? result = _generator.GenSweepVoltageForRow();

            Assert.AreNotEqual(null, result);
            Assert.AreEqual(1, result!.Count);
            Assert.AreEqual("for", result[0].Opcode);
            Assert.AreEqual("SrcCodeIndexZ = 1; SrcCodeIndexZ <= 5; SrcCodeIndexZ += 1", result[0].Parameter);
        }

        [TestMethod]
        public void GenSweepVoltageForRow_NestSweep_ShouldReturnNull()
        {
            _generator.Pat.SweepVoltage = new Dictionary<string, List<SweepVData>>
            {
                {
                    "NESTSWEEP_X",
                    new List<SweepVData>
                    {
                        new("X,0,1,1", isConvert: false)
                    }
                }
            };

            List<FlowRow>? result = _generator.GenSweepVoltageForRow();

            Assert.AreEqual(null, result);
        }

        [TestMethod]
        public void GenSweepVoltageForRow_FlowRepeat_ShouldReturnNull()
        {
            _generator.Pat.SweepVoltage = new Dictionary<string, List<SweepVData>>
            {
                {
                    "A",
                    new List<SweepVData>
                    {
                        new("A,0,1,1", isConvert: false)
                    }
                }
            };
            _generator.Pat.FlowControlFlag = 1;
            _generator.Pat.IsFlowInsRepeat = true;

            List<FlowRow>? result = _generator.GenSweepVoltageForRow();

            Assert.AreEqual(null, result);
        }

        [TestMethod]
        public void GenSweepCodeOrVoltageNextRow_ForNumberGtZero_ShouldGenerateNextRows()
        {
            _generator.Pat.SweepCodes = new Dictionary<string, List<SweepCode>>
            {
                {
                    "A",
                    new List<SweepCode>
                    {
                        new()
                        {
                            SendBitName = "SB",
                            SweepInfo = "0,1,1",
                            Width = 1,
                            Type = SweepCode.SweepType.Common
                        }
                    }
                }
            };

            _generator.Pat.FlowControlFlag = 0;
            _generator.Pat.IsFlowInsRepeat = false;

            List<FlowRow>? sweepRows = _generator.GenSweepCodeForRow();
            Assert.AreNotEqual(null, sweepRows);

            List<FlowRow>? result = _generator.GenSweepCodeOrVoltageNextRow();

            Assert.AreNotEqual(null, result);
            Assert.AreEqual("next", result![0].Opcode);

            List<FlowRow>? second = _generator.GenSweepCodeOrVoltageNextRow();
            Assert.AreEqual(null, second);
        }

        [TestMethod]
        public void GenSweepCodeOrVoltageNextRow_ForNumberZero_ShouldReturnNull()
        {
            List<FlowRow>? result = _generator.GenSweepCodeOrVoltageNextRow();
            Assert.AreEqual(null, result);
        }

        [TestMethod]
        public void GenOpcodeRowsAftPat_WithMisc_ShouldGenerate()
        {
            _generator.Pat.MiscInfo = "Opcode:A";

            List<FlowRow> result = _generator.GenOpcodeRowsAftPat();

            Assert.AreNotEqual(null, result);
        }

        [TestMethod]
        public void GenRelayRows_WithRelaySetting_ShouldGenerateRows()
        {
            _generator.Pat.RelaySetting = new Dictionary<string, string>
            {
                { "PIN1", "ON" }
            };

            List<FlowRow> result = _generator.GenRelayRows(false);

            Assert.AreNotEqual(null, result);
            Assert.IsTrue(result.Count > 0);
        }

        [TestMethod]
        public void GenRelayRows_Empty_ShouldReturnEmpty()
        {
            _generator.Pat.RelaySetting = [];

            List<FlowRow> result = _generator.GenRelayRows(false);

            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GenNwireDisOrEnableRows_ShouldGenerate()
        {
            _generator.Pat.MiscInfo = "FreeRunClk_enable:CLK1";

            List<FlowRow> result = _generator.GenNwireDisOrEnableRows();

            Assert.AreEqual(1, result.Count);
        }

        [TestMethod]
        public void GenExtraRowsByMisc_ShouldGenerateCallRows()
        {
            _generator.Pat.MiscInfo = "CallExtraFlow:FLOW_A;Other:XYZ;CallExtraFlow:FLOW_B";

            List<FlowRow> result = _generator.GenExtraRowsByMisc();

            Assert.AreEqual(2, result.Count);
            Assert.AreEqual(OpCode.Call, result[0].Opcode);
            Assert.AreEqual("FLOW_A", result[0].Parameter);
            Assert.AreEqual(OpCode.Call, result[1].Opcode);
            Assert.AreEqual("FLOW_B", result[1].Parameter);
        }

        [TestMethod]
        public void GenExtraRowsByMisc_ShouldIgnoreWhenNoColon()
        {
            _generator.Pat.MiscInfo = "CallExtraFlow;Other:ABC";

            List<FlowRow> result = _generator.GenExtraRowsByMisc();

            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GenExtraRowsByMisc_ShouldReturnEmptyWhenNoMatch()
        {
            _generator.Pat.MiscInfo = "Other:ABC;Test:123";

            List<FlowRow> result = _generator.GenExtraRowsByMisc();

            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void CreateTestOpcode_NoOpcodePrefix_ShouldReturnTest()
        {
            HardIpParaData paraData = new HardIpParaData(EnumBlock.HardIp);
            HardIpInputData input = new HardIpInputData(paraData);
            var gen = new FlowRowGeneratorBaseTestStub(input, "sheet")
            {
                Pat = new HardIpPattern
                {
                    Pattern = new PatternClass("PAT")
                    {
                        RealPatternName = "PAT"
                    }
                }
            };

            string result = gen.CallCreateTestOpcode();

            Assert.AreEqual(OpCode.Test, result);
        }

        [TestMethod]
        public void CreateTestOpcode_ValidOpcode_ShouldReturnOpcode()
        {
            HardIpParaData paraData = new HardIpParaData(EnumBlock.HardIp);
            HardIpInputData input = new HardIpInputData(paraData);
            var gen = new FlowRowGeneratorBaseTestStub(input, "sheet")
            {
                Pat = new HardIpPattern
                {
                    Pattern = new PatternClass("Opcode:Test")
                    {
                        RealPatternName = "Opcode:Test"
                    }
                }
            };

            string result = gen.CallCreateTestOpcode();

            Assert.AreEqual(OpCode.Test, result);
        }

        [TestMethod]
        public void CreateTestOpcode_InvalidOpcode_ShouldFallbackToTest()
        {
            HardIpParaData paraData = new HardIpParaData(EnumBlock.HardIp);
            HardIpInputData input = new HardIpInputData(paraData);
            var gen = new FlowRowGeneratorBaseTestStub(input, "sheet")
            {
                Pat = new HardIpPattern
                {
                    Pattern = new PatternClass("Opcode:FakeOp")
                    {
                        RealPatternName = "Opcode:FakeOp"
                    }
                }
            };

            string result = gen.CallCreateTestOpcode();

            Assert.AreEqual(OpCode.Test, result);
        }
    }
}
