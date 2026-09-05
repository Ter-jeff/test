using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.DividerManager.FlowDividerManager;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;

using CommonLib.Enums;

using CommonReaderLib.PatternListCsv;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class FlowLimitDividerTests : FunctionTestBase
    {
        private static FlowLimitDivider _divider = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            _divider = new FlowLimitDivider();
            LocalSpecs.Options.Device = EnumDevice.LCD;
        }

        [TestMethod]
        public void DivideUseLimit_Should_Return_Same_PatternList_When_ExtendLimits_IsFalse()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                SubBlock = "FT",
                MeasPins =
                [
                    new()
                    {
                        PinName = "VDD",
                        MeasType = "MEAS_DC",
                        MeasLimitsH = [new("JOB1") { HiLimit = "1.2", LoLimit = "0.8" }],
                        ForceConditions = []
                    }
                ]
            };

            HardIpParaData paraData = new HardIpParaData(EnumBlock.HardIp);
            var hardIpInputData = new HardIpInputData(paraData);

            // Act
            List<HardIpPattern> result = _divider.DivideUseLimit([pattern], hardIpInputData);

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(pattern, result[0]);
        }

        [TestMethod]
        public void GenerateLimitByJobTest()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                SubBlock = "FT",
                Pattern = new PatternClass("dd_cmna0_a_fulp_an_rx01_bst_jtg_elb_allfv_si_pwrdnf3_tx3l4")
                {
                    PatternSetList =
                    [
                        ["dd_cmna0_a_fulp_an_rx01_bst_jtg_elb_allfv_si_pwrdnf3_tx3l4"]
                    ]
                },
                MeasPins =
                [
                    new()
                    {
                        PinName = "VDD",
                        MeasType = "MEAS_DC",
                        MeasLimitsH = [new("JOB1") { HiLimit = "1.2", LoLimit = "0.8" }],
                        ForceConditions = []
                    }
                ],
                TestPlanSequences = [new(1, 1, 1)]
            };

            // Act
            List<MeasPin> result = _divider.GenerateLimitByJob(pattern, "NV");
            bool flag = _divider.CheckMergePower(result[0]);

            // Assert
            Assert.AreEqual("VDD", result[0].PinName);
            Assert.IsFalse(flag);
        }

        [TestMethod]
        [DataRow("", "CP1", "0.1", "0.9")]
        [DataRow("H", "CP2", "1.0", "2.0")]
        [DataRow("N", "CP2", "1.0", "2.0")]
        [DataRow("L", "CP2", "1.0", "2.0")]
        public void GroupLimitsTest(string voltage, string jobName, string lo, string hi)
        {
            // Arrange
            var pin = new MeasPin();
            switch (voltage)
            {
                case "H":
                    pin.MeasLimitsH.Add(new MeasLimit(jobName) { LoLimit = lo, HiLimit = hi });
                    break;
                case "L":
                    pin.MeasLimitsL.Add(new MeasLimit(jobName) { LoLimit = lo, HiLimit = hi });
                    break;
                case "N":
                case "":
                    pin.MeasLimitsN.Add(new MeasLimit(jobName) { LoLimit = lo, HiLimit = hi });
                    break;
            }

            // Act
            Dictionary<string, List<string>> result = _divider.GroupLimits(pin, voltage);

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual($"{lo}${hi}", result.Keys.ToList()[0]);
        }

        [DataTestMethod]
        [DataRow("CP1=Pin_A", "Pin_A")]
        [DataRow("Pin_B", "Pin_B")]
        [DataRow("JOB1=CLK_1", "CLK_1")]
        [DataRow("=", "")]
        public void RemoveJobName_ShouldReturnExpectedPinName(string input, string expected)
        {
            // Act
            string result = _divider.RemoveJobName(input);

            // Assert
            Assert.AreEqual(expected, result);
        }
    }
}
