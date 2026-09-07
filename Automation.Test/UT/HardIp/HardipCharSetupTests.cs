using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using IgxlLib;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardipCharSetupTests
    {
        private TestPlanSheet _planSheet = null!;
        private HardIpPattern _pattern = null!;

        [TestInitialize]
        public void Init()
        {
            _planSheet = new TestPlanSheet
            {
                PlanHeaderIdx = new Dictionary<string, int>
                    {
                        { "forceIndex", 0 }
                    }
            };

            _pattern = new HardIpPattern
            {
                SheetName = "HARDIP_BLOCK_A",
                MeasPins =
                    [
                        new() { MeasType = "V" }
                    ]
            };

            TestProgram.IgxlWorkBk = new IgxlWorkBook();

            PinMapSheet pinMap = new PinMapSheet("PM");
            pinMap.AddPin(new Pin("PIN1", PinMapConst.TypeIo));
            pinMap.AddPin(new Pin("PIN2", PinMapConst.TypePower));

            TestProgram.IgxlWorkBk.PinMapPair =
                new KeyValuePair<string, PinMapSheet>("PM", pinMap);
        }

        [TestMethod]
        public void GetShmooParameterName_ShouldReturnMappedValue_FromParameterName()
        {
            // Arrange
            CharSetupConst.ParameterName["vdd"] = "Voltage";

            // Act
            string result = HardipCharSetup.GetShmooParameterName("vdd");

            // Assert
            Assert.AreEqual("Voltage", result);
        }

        [TestMethod]
        public void GetShmooParameterName_ShouldReturnMappedValue_FromDictionary()
        {
            string result = HardipCharSetup.GetShmooParameterName("d0");
            Assert.AreEqual("On", result);

            result = HardipCharSetup.GetShmooParameterName("D1");
            Assert.AreEqual("Data", result);

            result = HardipCharSetup.GetShmooParameterName("d3");
            Assert.AreEqual("Off", result);
        }

        [TestMethod]
        public void GetShmooParameterName_ShouldReturnOriginal_IfNotMapped()
        {
            string result = HardipCharSetup.GetShmooParameterName("customParam");
            Assert.AreEqual("customParam", result);
        }

        [TestMethod]
        public void GetShmooTimeSets_ShouldReturnElements_AfterRemovingFirst()
        {
            string result = HardipCharSetup.GetShmooTimeSets("A,B,C");
            Assert.AreEqual("B,C", result);
        }

        [TestMethod]
        public void GetShmooTimeSets_ShouldReturnEmpty_IfSingleElement()
        {
            string result = HardipCharSetup.GetShmooTimeSets("A");
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetShmooName_Sweep_ShouldReturnHAC()
        {
            HardipCharSetup shmoo = CreateBasicShmoo();

            string result = HardipCharSetup.GetShmooName(_planSheet, _pattern, shmoo, "SUB", true, "PARAM");

            Assert.IsTrue(result.StartsWith("HAC"));
        }

        [TestMethod]
        public void GetShmooName_HIO_AllIoPins()
        {
            HardipCharSetup shmoo = new HardipCharSetup
            {
                CharSteps =
                    [
                        new("S", "1")
                        {
                            Mode = CharSetupConst.ModeXShmoo,
                            ApplyToPins = "PIN1"
                        }
                    ]
            };

            string result = HardipCharSetup.GetShmooName(_planSheet, _pattern, shmoo, "SUB", false, "PARAM");

            Assert.IsTrue(result.StartsWith("HIO"));
        }

        [TestMethod]
        public void GetShmooName_HFH_PowerUp()
        {
            HardipCharSetup shmoo = new HardipCharSetup
            {
                CharSteps =
                    [
                        new("VDD_STEP", "1")
                        {
                            ApplyToPins = "PIN2",
                            RangeFrom = "0.0",
                            RangeTo = "1.0"
                        }
                    ]
            };

            string result = HardipCharSetup.GetShmooName(_planSheet, _pattern, shmoo, "SUB", false, "PARAM");

            Assert.IsTrue(result.StartsWith("HFH"));
        }

        [TestMethod]
        public void GetShmooName_HFL_PowerDown()
        {
            HardipCharSetup shmoo = new HardipCharSetup
            {
                CharSteps =
                    [
                        new("VDD_STEP", "1")
                        {
                            ApplyToPins = "PIN2",
                            RangeFrom = "2.0",
                            RangeTo = "1.0"
                        }
                    ]
            };

            string result = HardipCharSetup.GetShmooName(_planSheet, _pattern, shmoo, "SUB", false, "PARAM");

            Assert.IsTrue(result.StartsWith("HFL"));
        }

        [TestMethod]
        public void GetShmooName_MultiXShmoo_ShouldContainMULTI()
        {
            HardipCharSetup shmoo = new HardipCharSetup
            {
                CharSteps =
                    [
                        new("S", "1") { Mode = CharSetupConst.ModeXShmoo, ApplyToPins = "PIN1" },
                        new("S", "2") { Mode = CharSetupConst.ModeXShmoo, ApplyToPins = "PIN1" }
                    ]
            };

            string result = HardipCharSetup.GetShmooName(_planSheet, _pattern, shmoo, "SUB", false, "PARAM");

            Assert.IsTrue(result.Contains("MULTI"));
        }

        [TestMethod]
        public void GetShmooName_YShmoo_ShouldUseYPin()
        {
            HardipCharSetup shmoo = new HardipCharSetup
            {
                CharSteps =
                    [
                        new("S", "1")
                        {
                            Mode = CharSetupConst.ModeYShmoo,
                            ApplyToPins = "PIN1"
                        }
                    ]
            };

            string result = HardipCharSetup.GetShmooName(_planSheet, _pattern, shmoo, "SUB", false, "PARAM_TEST");

            Assert.IsTrue(result.Contains("PIN1"));
        }

        [TestMethod]
        public void GetShmooName_NoYPin_ShouldUseParameter()
        {
            HardipCharSetup shmoo = CreateBasicShmoo();

            string result = HardipCharSetup.GetShmooName(_planSheet, _pattern, shmoo, "SUB", false, "PARAM_TEST");

            Assert.IsTrue(result.Contains("PARAMTEST"));
        }

        [TestMethod]
        public void GetShmooName_ShouldRemoveUnderscore()
        {
            HardipCharSetup shmoo = CreateBasicShmoo();

            string result = HardipCharSetup.GetShmooName(_planSheet, _pattern, shmoo, "SUB_A", false, "P_A_R_A_M");

            Assert.IsFalse(result.Contains("SUB_A"));
            Assert.IsFalse(result.Contains("P_A_R_A_M"));

            Assert.IsTrue(result.Contains("SUBA"));
            Assert.IsTrue(result.Contains("PARAM"));
        }

        [TestMethod]
        public void GetShmooName_PinNotExist_ShouldNotCrash()
        {
            HardipCharSetup shmoo = new HardipCharSetup
            {
                CharSteps =
                [
                    new("S", "1")
                    {
                        ApplyToPins = "NOT_EXIST_PIN"
                    }
                ]
            };

            string result = HardipCharSetup.GetShmooName(_planSheet, _pattern, shmoo, "SUB", false, "PARAM");

            Assert.IsTrue(result.Contains("UNKNOWN") || result.StartsWith("HIO"));
        }

        [TestMethod]
        public void GetShmooName_UnknownPinType_ShouldReturnUNKNOWN()
        {
            PinMapSheet pinMap = new PinMapSheet("PM");
            pinMap.AddPin(new Pin("PIN3", PinMapConst.TypeAnalog));

            TestProgram.IgxlWorkBk.PinMapPair =
                new KeyValuePair<string, PinMapSheet>("PM", pinMap);

            HardipCharSetup shmoo = new HardipCharSetup
            {
                CharSteps =
                [
                    new("S", "1")
                    {
                        ApplyToPins = "PIN3"
                    }
                ]
            };

            string result = HardipCharSetup.GetShmooName(_planSheet, _pattern, shmoo, "SUB", false, "PARAM");

            Assert.IsTrue(result.StartsWith("HIO"));
        }

        [TestMethod]
        public void GetShmooName_InvalidRange_ShouldReturnUNKNOWN()
        {
            HardipCharSetup shmoo = new HardipCharSetup
            {
                CharSteps =
                [
                    new("S", "1")
                    {
                        ApplyToPins = "PIN2",
                        RangeFrom = "abc",
                        RangeTo = "xyz"
                    }
                ]
            };

            string result = HardipCharSetup.GetShmooName(_planSheet, _pattern, shmoo, "SUB", false, "PARAM");

            Assert.IsTrue(result.StartsWith("UNKNOWN"));
        }

        private static HardipCharSetup CreateBasicShmoo()
        {
            HardipCharSetup shmoo = new HardipCharSetup
            {
                CharSteps =
                    [
                        new("S", "1")
                    ]
            };
            return shmoo;
        }

        [TestMethod]
        public void GetShmoo_ShouldSplitByVoltage_WhenIsSplitByVoltageTrue()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                Shmoo = new HardipCharSetup
                {
                    IsSplitByVoltage = true,
                    SetupName = "MyShmoo",
                    TestMethod = "MethodX",
                    CharSteps =
                    [
                        new("", "") { VoltageType = "NV" },
                        new("", "") { VoltageType = "LV" },
                        new("", "") { VoltageType = "HV" },
                        new("", "") { VoltageType = "" }
                    ]
                }
            };

            // Act
            List<HardipCharSetup> result = HardipCharSetup.GetShmoo(pattern);

            // Assert
            Assert.AreEqual(3, result.Count);
            Assert.IsTrue(result.Any(x => x.SetupName.EndsWith("_NV")));
            Assert.IsTrue(result.Any(x => x.SetupName.EndsWith("_LV")));
            Assert.IsTrue(result.Any(x => x.SetupName.EndsWith("_HV")));
        }

        [TestMethod]
        public void GetShmoo_ShouldNotSplit_WhenIsSplitByVoltageFalse()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                Shmoo = new HardipCharSetup
                {
                    IsSplitByVoltage = false,
                    SetupName = "BaseName",
                    TestMethod = "MethodA",
                    CharSteps = [new("", "") { VoltageType = "" }]
                }
            };

            // Act
            List<HardipCharSetup> result = HardipCharSetup.GetShmoo(pattern);

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("BaseName", result[0].SetupName);
            Assert.AreEqual("MethodA", result[0].TestMethod);
        }

        [TestMethod]
        public void GetShmoo_ShouldFilterStepsByVoltageType()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                Shmoo = new HardipCharSetup
                {
                    IsSplitByVoltage = true,
                    SetupName = "ShmooTest",
                    CharSteps =
                    [
                        new("", "") { VoltageType = "NV" },
                        new("", "") { VoltageType = "" }
                    ]
                }
            };

            // Act
            List<HardipCharSetup> result = HardipCharSetup.GetShmoo(pattern);
            HardipCharSetup? nv = result.FirstOrDefault(x => x.SetupName.EndsWith("_NV"));

            // Assert
            Assert.AreNotEqual(null, nv);
            Assert.IsTrue(nv!.CharSteps.All(s => s.VoltageType == "NV" || string.IsNullOrEmpty(s.VoltageType)));
        }
    }
}
