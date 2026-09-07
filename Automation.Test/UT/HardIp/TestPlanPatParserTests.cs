using System.Collections.Generic;
using System.ComponentModel;
using System.IO;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class TestPlanPatParserTests : FunctionTestBase
    {
        private static TestPlanPatParser _parser = null!;
        private static SweepVoltageResolver _sweepVoltageResolver = null!;
        private static CalcEqnResolver _calcEqnResolver = null!;
        private static MeasPinResolver _measPinResolver = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            var planSheet = new TestPlanSheet
            {
                PlanHeaderIdx = new Dictionary<string, int>
                {
                    { "miscInfoIndex", 1 },
                    { "forceIndex", 2 },
                    { "measIndex", 3 },
                },
                SheetName = "SheetName_A1"
            };
            _parser = new TestPlanPatParser(planSheet);
            _sweepVoltageResolver = new SweepVoltageResolver(planSheet);
            _calcEqnResolver = new CalcEqnResolver(planSheet);
            _measPinResolver = new MeasPinResolver(planSheet, new ForceConditionResolver(planSheet, _calcEqnResolver));
        }

        [TestMethod]
        [DisplayName("01_GetSweepVoltage_ShouldHandleForceConditionSweepVoltage")]
        public void GetSweepVoltage_ShouldHandleForceConditionSweepVoltage()
        {
            // Arrange
            var pattern = new PatternRow
            {
                MiscInfo = "forloop:1,2,3,4;B",
                ForceCondition = new ForceClass(),
                SheetName = "Sheet1",
                RowNum = 1,
                Pattern = new PatternClass("PAT1") { RealPatternName = "PAT1" }
            };

            // Act
            Dictionary<string, List<SweepVData>> result = _sweepVoltageResolver.GetSweepVoltage(pattern);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.IsTrue(result.ContainsKey("1"));
            List<SweepVData> sweepList = result["1"];
            Assert.AreNotEqual(null, sweepList);
            Assert.AreEqual(1, sweepList.Count);
        }

        [TestMethod]
        public void GetSweepVoltage_ShouldHandleForceConditionSweepVoltage1()
        {
            // Arrange
            var pattern = new PatternRow
            {
                MiscInfo = "sweep(PinA:V:0.1)",
                ForceCondition = new ForceClass(),
                SheetName = "Sheet1",
                RowNum = 1,
                Pattern = new PatternClass("PAT1") { RealPatternName = "PAT1" }
            };

            // Act
            Dictionary<string, List<SweepVData>> result = _sweepVoltageResolver.GetSweepVoltage(pattern);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetSweepVoltage_ShouldHandleForceConditionSweepVoltage2()
        {
            // Arrange
            var pattern = new PatternRow
            {
                MiscInfo = "sweepX[0.8, 1.0, 1.1]",
                ForceCondition = new ForceClass(),
                SheetName = "Sheet1",
                RowNum = 1,
                Pattern = new PatternClass("PAT1") { RealPatternName = "PAT1" }
            };

            // Act
            Dictionary<string, List<SweepVData>> result = _sweepVoltageResolver.GetSweepVoltage(pattern);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetFilterCalcEqn_TargetMatched_RemoveIfEmptyReplaceKey()
        {
            string input = "[A]+[B]+[C]";
            string result = _calcEqnResolver.GetFilterCalcEqn(input, "B", "");
            Assert.AreEqual("[A]+[]+[C]", result);
        }

        [TestMethod]
        public void GetFilterCalcEqn_TargetMatched_ReplaceWithReplaceKey()
        {
            string input = "[A]+[B]+[C]";
            string result = _calcEqnResolver.GetFilterCalcEqn(input, "B", "X");
            Assert.AreEqual("[A]+[X]+[C]", result);
        }

        [TestMethod]
        public void GetFilterCalcEqn_TargetNotMatched_NoChange()
        {
            string input = "[A]+[B]+[C]";
            string result = _calcEqnResolver.GetFilterCalcEqn(input, "D", "X");
            Assert.AreEqual("[A]+[B]+[C]", result);
        }

        [TestMethod]
        public void GetFilterCalcEqn_NoMatches_ReturnsEmpty()
        {
            string input = "No brackets here";
            string result = _calcEqnResolver.GetFilterCalcEqn(input, "A", "X");
            Assert.AreEqual("No brackets here", result);
        }

        [TestMethod]
        public void GetFilterCalcEqn_CaseInsensitiveTargetMatching()
        {
            string input = "[a]+[B]+[C]";
            string result = _calcEqnResolver.GetFilterCalcEqn(input, "a", "X");
            Assert.AreEqual("[X]+[B]+[C]", result);
        }

        [TestMethod]
        public void ConvertTpPattern_ShouldReturnNull_WhenRowIsNull()
        {
            // Act
            HardIpPattern? result = _parser.ConvertTpPattern(null);

            // Assert
            Assert.AreEqual(null, result, "Method should handle null input gracefully.");
        }

        [TestMethod]
        public void ConvertTpPattern_ShouldParseMultiTimeDomainMiscInfo()
        {
            // Arrange
            var row = new PatternRow
            {
                Pattern = new PatternClass("P1#P2")
                {
                    RealPatternName = "PAT_A#PAT_B",
                    TestPlanPatternName = "PAT_B", // Index 1
                },
                MiscInfo = "ref_subblock:BLOCK_0#BLOCK_1;mode:fast;",
                NoBinOutStr = "BIN1;BIN2"
            };

            // Act
            HardIpPattern? result = _parser.ConvertTpPattern(row);

            // Assert
            Assert.AreEqual("BLOCK_0#BLOCK_1", result!.SubBlock);
            Assert.IsTrue(result.MiscInfo.Contains("ref_subblock:BLOCK_1;"), "MiscInfo should update to the specific ref_subblock.");
            Assert.IsTrue(result.NoBinoutStr.Contains("BIN1"));
        }

        [TestMethod]
        public void ConvertTpPattern()
        {
            string subName = "ConvertTpPattern";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            var rows = new List<PatternRow>()
            {
                 new()
                 {
                    Pattern = new PatternClass("P6")
                    {
                    },
                    MiscInfo = "ateTestCondition;overlayName;Func:CreateOverlay",
                    ForceCondition = new ForceClass(){ ForceCondition = "$VDD_A:1.8;VDD_B:1.8:1.9,1.9:2.0:CP1;VDD_C1,VDD_C2,VDD_C3@VDD_D:1.8:1.9:2.0:2.1;VDD_E;" }
                },
                new()
                {
                    Pattern = new PatternClass("P5")
                    {
                    },
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "MeasVdiff pin = -----",
                                     MiscInfo = "",
                                     ForceCondition = "",
                                     MergeRowNumForMeas = 1,
                                     Limits =
                                     [
                                         new("CP1"){ HiLimit = "TBD" , LoLimit = "TBD"},
                                         new("CP1"){ HiLimit = "1,2" , LoLimit = "1,2"},
                                         new("CP1"){ HiLimit = "1,2,3" , LoLimit = "1,2,3"},
                                     ]
                                },
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P4")
                    {
                    },
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "XXXX pin = JTAG(data):4 \"testName\"",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                },
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P3")
                    {
                    },
                    MiscInfo = "Repeat_Limit:1",
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "WiMeas pin = JTAG(data):4 \"testName\"",
                                     MiscInfo = "",
                                     ForceCondition = "",
                                     RfIntrumentSetup = "A=1;"
                                },
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P2")
                    {
                    },
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "WiMeas pin = JTAG(data):4 \"testName\"",
                                     MiscInfo  = "Repeat_Limit:1",
                                     ForceCondition = ""
                                },
                                new()
                                {
                                     Meas = "WiSrc pin = JTAG(data):4 \"testName\"",
                                     MiscInfo = "",
                                     ForceCondition = "",
                                },
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P1")
                    {
                    },
                    MiscInfo = "calc_eqn:\"AAA\"",
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "MeasC pin = JTAG(data):4 \"testName\"",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                },
                                new()
                                {
                                     Meas = "Calc \"test1:AA\",VDD1_DDR012_S2(10E+2)",
                                     MiscInfo = "forcecalctype:V",
                                     ForceCondition = "",
                                },
                                new()
                                {
                                     Meas = "MeasWait 1",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                }
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P1")
                    {
                    },
                    MiscInfo = "forloop:1,1,1,1;sweep(PinA:V);sweep[PinA:V]",
                    ForceCondition = new ForceClass(){ ForceCondition = "VDD_PIN:v1p:1.8;VSS_PIN:v1n:0.0;sweep(PinA:V);nestsweep[PinA:V:1];nestsweep[PinA:V:1]" }
                },
                new()
                {
                    Pattern = new PatternClass("P5")
                    {
                    },
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "MeasVdiff pin = -----",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                },
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P4")
                    {
                    },
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "XXXX pin = JTAG(data):4 \"testName\"",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                },
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P3")
                    {
                    },
                    MiscInfo = "Repeat_Limit:1",
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "WiMeas pin = JTAG(data):4 \"testName\"",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                },
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P2")
                    {
                    },
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "WiMeas pin = JTAG(data):4 \"testName\"",
                                     MiscInfo  = "Repeat_Limit:1",
                                     ForceCondition = ""
                                },
                                new()
                                {
                                     Meas = "WiSrc pin = JTAG(data):4 \"testName\"",
                                     MiscInfo = "",
                                     ForceCondition = "",
                                },
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P1")
                    {
                    },
                    MiscInfo = "calc_eqn:\"AAA\"",
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "MeasC pin = JTAG(data):4 \"testName\"",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                },
                                new()
                                {
                                     Meas = "Calc \"test1:AA\",VDD_B(10E+2)+N/A",
                                     MiscInfo = "forcecalctype:V",
                                     ForceCondition = "",
                                },
                                new()
                                {
                                     Meas = "MeasWait 1",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                }
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P1")
                    {
                    },
                    MiscInfo = "forloop:1,1,1,1;sweep(PinA:V);sweep[PinA:V]",
                    ForceCondition = new ForceClass(){ ForceCondition = "sweep(PinA:V);nestsweep[PinA:V:1];nestsweep[PinA:V:1]" }
                },
                new()
                {
                    Pattern = new PatternClass("P5")
                    {
                    },
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "MeasVdiff pin = -----",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                },
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P4")
                    {
                    },
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "XXXX pin = JTAG(data):4 \"testName\"",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                },
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P3")
                    {
                    },
                    MiscInfo = "Repeat_Limit:1",
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "WiMeas pin = JTAG(data):4 \"testName\"",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                },
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P2")
                    {
                    },
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "WiMeas pin = JTAG(data):4 \"testName\"",
                                     MiscInfo  = "Repeat_Limit:1",
                                     ForceCondition = ""
                                },
                                new()
                                {
                                     Meas = "WiSrc pin = JTAG(data):4 \"testName\"",
                                     MiscInfo = "",
                                     ForceCondition = "",
                                },
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P1")
                    {
                    },
                    MiscInfo = "calc_eqn:\"AAA\"",
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = "MeasC pin = JTAG(data):4 \"testName\"",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                },
                                new()
                                {
                                     Meas = "Calc \"test1:AA\",VDD_B(10E+2)+N/A",
                                     MiscInfo = "forcecalctype:V",
                                     ForceCondition = "",
                                },
                                new()
                                {
                                     Meas = "MeasWait 1",
                                     MiscInfo = "",
                                     ForceCondition = ""
                                }
                            ]
                        }
                    ]
                },
                new()
                {
                    Pattern = new PatternClass("P1")
                    {
                    },
                    MiscInfo = "forloop:1,1,1,1;sweep(PinA:V);sweep[PinA:V]",
                    ForceCondition = new ForceClass(){ ForceCondition = "sweep(PinA:V);nestsweep[PinA:V:1];nestsweep[PinA:V:1]" }
                },
                new()
                {
                    Pattern = new PatternClass("P1#P2")
                    {
                        RealPatternName = "PAT_A#PAT_B",
                        TestPlanPatternName = "PAT_B",
                    },
                    MiscInfo = "ref_subblock:BLOCK_0#BLOCK_1;mode:fast",
                    NoBinOutStr = "BIN1;BIN2",
                    PatChildRows=
                    [
                        new PatSubChildRow()
                        {
                            TpRows=
                            [
                                new()
                                {
                                     Meas = MeasType.MeasWait,
                                     MiscInfo= "MeasRange:100;MeasWait:1;"
                                }
                            ]
                        }
                    ]
                }
            };

            // Act
            List<HardIpPattern> results = [];
            foreach (PatternRow row in rows)
            {
                results.Add(_parser.ConvertTpPattern(row)!);
            }

            // Assert
            string json = JsonConvert.SerializeObject(results, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [DataTestMethod]
        [DataRow("NV@Something", "NV")]
        [DataRow("HV@Something", "HV")]
        [DataRow("LV@Something", "LV")]
        [DataRow("Other@Something", "")]
        public void GetInstanceVoltageInCondition_MapsPrefixToVoltage(string forceCondition, string expected)
        {
            // Act
            string result = _sweepVoltageResolver.GetInstanceVoltageInCondition(forceCondition);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetMeasWaitTime_MiscInfoHasMeasWait_ReturnsValue()
        {
            // Arrange
            var tpRow = new TestPlanRow { MiscInfo = "MeasWait:5", Meas = "" };

            // Act
            string result = _measPinResolver.GetMeasWaitTime(tpRow);

            // Assert
            Assert.AreEqual("5", result);
        }

        [TestMethod]
        public void GetMeasWaitTime_MeasContainsMeasWaitKeyword_ReturnsRemainder()
        {
            // Arrange
            var tpRow = new TestPlanRow { MiscInfo = "", Meas = "MeasWait 100" };

            // Act
            string result = _measPinResolver.GetMeasWaitTime(tpRow);

            // Assert
            Assert.AreEqual("100", result);
        }

        [TestMethod]
        public void GetMeasWaitTime_NoMatch_ReturnsEmpty()
        {
            // Arrange
            var tpRow = new TestPlanRow { MiscInfo = "", Meas = "" };

            // Act
            string result = _measPinResolver.GetMeasWaitTime(tpRow);

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetMeasIRange_MiscInfoHasMeasRange_ReturnsValue()
        {
            // Arrange
            var tpRow = new TestPlanRow { MiscInfo = "MeasRange:10", Meas = "" };

            // Act
            string result = _measPinResolver.GetMeasIRange(tpRow);

            // Assert
            Assert.AreEqual("10", result);
        }

        [TestMethod]
        public void GetMeasIRange_NoMatch_ReturnsEmpty()
        {
            // Arrange
            var tpRow = new TestPlanRow { MiscInfo = "", Meas = "" };

            // Act
            string result = _measPinResolver.GetMeasIRange(tpRow);

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetChiplet_SheetNameMatchesPattern_ReturnsChipletFromSheetName()
        {
            // Act
            string result = _parser.GetChiplet("Sheet_A1", ["B2"]);

            // Assert
            Assert.AreEqual("A1", result);
        }

        [TestMethod]
        public void GetChiplet_NoMatch_FallsBackToChipletListFirst()
        {
            // Act
            string result = _parser.GetChiplet("NoMatchingSheetName", ["B2"]);

            // Assert
            Assert.AreEqual("B2", result);
        }

        [TestMethod]
        public void GetChiplet_NoMatchAndNullList_ReturnsEmpty()
        {
            // Act
            string result = _parser.GetChiplet("NoMatchingSheetName", null);

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

    }
}
