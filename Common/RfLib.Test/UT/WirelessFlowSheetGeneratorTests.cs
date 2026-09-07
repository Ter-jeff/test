using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;
using Automation.Test.UT;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using RfLib.Dvdc.GenFlow;

namespace RfLibLib.Test.UT
{
    [TestClass]
    public class WirelessFlowSheetGeneratorTests : FunctionTestBase
    {
        [ClassInitialize]
        public static void ClassInit(TestContext context)
        {
            LocalSpecs.Clear();
            LocalSpecs.IsUnitTest = true;
            LocalSpecs.Options.Device = EnumDevice.RF;
        }

        [TestMethod]
        public void UpdateLimits_Should_Add_MeasPins_With_Correct_HighLowLimits_For_Psat()
        {
            var limits = new List<MeasPin>();

            var pin = new MeasPin
            {
                PinName = "VDD",
                TestName = "A_B_C_D_E_F_G",
                MeasType = "measi",
                InterPoseFunc = "psat",
                MeasLimitsH =
        [
            new MeasLimit("JOB1") { HiLimit = "1.2", LoLimit = "0.8" }
        ],
                MeasLimitsL = [],
                MeasLimitsN = []
            };

            WirelessFlowSheetGenerator.UpdateLimits(limits, pin, "H");

            Assert.AreEqual(3, limits.Count);
            Assert.AreEqual("1.2", limits[0].HighLimit);
            Assert.AreEqual("0.8", limits[0].LowLimit);
            Assert.AreEqual("1.2", limits[1].HighLimit);
            Assert.AreEqual("0.8", limits[2].LowLimit);
        }
        [TestMethod]
        public void UpdateLimits_WiSrc_ShouldNotAddAnyLimit()
        {
            var limits = new List<MeasPin>();

            var pin = new MeasPin
            {
                PinName = "VDD",
                TestName = "A_B_C_D_E",
                MeasType = "WiSrc",
                InterPoseFunc = "psat",
                MeasLimitsH =
            [
                new MeasLimit("JOB1") { HiLimit = "1.0", LoLimit = "0.5" }
            ]
            };

            WirelessFlowSheetGenerator.UpdateLimits(limits, pin, "H");

            Assert.AreEqual(0, limits.Count);
        }
    }
}
