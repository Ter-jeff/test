using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.Static;

using IgxlLib;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class ForceClassTests
    {
        private static void InitTestProgram()
        {
            TestProgram.Clear();

            IgxlWorkBook workbook = new IgxlWorkBook();

            List<Selector> acSelectors = [];
            AcSpec acSpec = new AcSpec("TCK", acSelectors);
            AcSpecSheet acSheet = new AcSpecSheet("AC_SPEC");
            acSheet.Rows.Add(acSpec);

            List<Selector> dcSelectors = [];
            DcSpec dcSpec = new DcSpec("VDD", dcSelectors);
            DcSpecSheet dcSheet = new DcSpecSheet("DC_SPEC");
            dcSheet.Rows.Add(dcSpec);

            workbook.AcSpecSheets.Add("AC", acSheet);
            workbook.DcSpecSheets.Add("DC", dcSheet);

            TestProgram.IgxlWorkBk = workbook;

            TestPlanStatic.PowerMergeSheet = new TestPlanLib.PowerMerge.PowerMergeSheet();
        }

        [TestMethod]
        public void GetLevelSetting_ShouldReturnLevelName_WhenLevelPatternExists()
        {
            // Arrange
            var force = new ForceClass
            {
                ForceCondition = "Level:LV1;Other:XYZ"
            };

            // Act
            string result = force.GetLevelSetting();

            // Assert
            Assert.AreEqual("LV1", result);
        }

        [TestMethod]
        public void GetAcSetting_ShouldReturnAcOnlyForValidPins()
        {
            // Arrange
            var force = new ForceClass
            {
                ForceCondition = "AC:TCK:Spec1;AC:ShiftIn:Spec2;AC:OtherPin:Spec3"
            };

            // Act
            string result = force.GetAcSetting();

            // Assert
            Assert.IsTrue(result.Contains("TCK"));
            Assert.IsTrue(result.Contains("ShiftIn"));
            Assert.IsFalse(result.Contains("OtherPin"));
        }

        [TestMethod]
        public void GetDcCategory_ShouldReturnOnlyDcEntries()
        {
            // Arrange
            var force = new ForceClass
            {
                ForceCondition = "DC:CAT1;AC:CAT2;DC:CAT3"
            };

            // Act
            string result = force.GetDcCategory();

            // Assert
            Assert.AreEqual("DC:CAT1;DC:CAT3", result);
        }

        [TestMethod]
        public void GetPrePatForceCondition_ShouldRemoveKnownPatterns()
        {
            // Arrange
            var force = new ForceClass
            {
                ForceCondition = "Level:LV1;AC:TCK:Spec1;DC:CAT1;TimeSets:TS1;CustomSweep:1:2:0.1"
            };

            // Act
            string result = force.GetPrePatForceCondition();

            // Assert
            Assert.AreEqual("CustomSweep:1:2:0.1", result);
        }

        [TestMethod]
        public void GetShmooInstanceVoltage_ShouldDetectVoltagePrefixes()
        {
            // Arrange
            var force = new ForceClass();
            var dummySetup = new HardipCharSetup();

            // Assert
            Assert.AreEqual("NV", force.GetShmooInstanceVoltage(dummySetup, "NV@forceX"));
            Assert.AreEqual("HV", force.GetShmooInstanceVoltage(dummySetup, "HV@forceX"));
            Assert.AreEqual("LV", force.GetShmooInstanceVoltage(dummySetup, "LV@forceX"));
            Assert.AreEqual("", force.GetShmooInstanceVoltage(dummySetup, "OtherForce"));
        }

        [TestMethod]
        public void CorrectTypo_ShouldFixRelaySyntax()
        {
            // Arrange
            var force = new ForceClass
            {
                ForceCondition = "relay_on:PIN_A;relay_off:PIN_B;Normal:OK"
            };

            // Act
            string result = force.ForceCondition;

            // Assert
            Assert.IsTrue(result.Contains("PIN_A:relay_on"));
            Assert.IsTrue(result.Contains("PIN_B:relay_off"));
            Assert.IsTrue(result.Contains("Normal:OK"));
        }

        [TestMethod]
        public void GetShmooInstanceVoltage_ShouldReturnHV_WhenContainsHV()
        {
            // Arrange
            var force = new ForceClass();
            var setup = new HardipCharSetup();

            // Act
            string result = force.GetShmooInstanceVoltage(setup, "HV@forceCondition");

            // Assert
            Assert.AreEqual("HV", result);
        }

        [TestMethod]
        public void GetShmooInstanceVoltage_ShouldReturnEmpty_WhenNoVoltagePrefix()
        {
            // Arrange
            var force = new ForceClass();
            var setup = new HardipCharSetup();

            // Act
            string result = force.GetShmooInstanceVoltage(setup, "NormalForce");

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void Copy_ShouldCloneAllFlags()
        {
            // Arrange
            var original = new ForceClass
            {
                IsVtShmoo = true,
                IsShmooInForce = true,
                IsShmooInProdInst = false,
                IsShmooInProdFlow = true,
                IsShmooInCharInst = false,
                IsShmooInCharFlow = true,
                IsCz2InstName = true,
                ForceCondition = "TestForce"
            };

            // Act
            ForceClass copied = original.Copy();

            // Assert
            Assert.AreNotEqual(null, copied);
            Assert.AreNotSame(original, copied);

            Assert.AreEqual(original.IsVtShmoo, copied.IsVtShmoo);
            Assert.AreEqual(original.IsShmooInForce, copied.IsShmooInForce);
            Assert.AreEqual(original.IsShmooInProdInst, copied.IsShmooInProdInst);
            Assert.AreEqual(original.IsShmooInProdFlow, copied.IsShmooInProdFlow);
            Assert.AreEqual(original.IsShmooInCharInst, copied.IsShmooInCharInst);
            Assert.AreEqual(original.IsShmooInCharFlow, copied.IsShmooInCharFlow);
            Assert.AreEqual(original.IsCz2InstName, copied.IsCz2InstName);
            Assert.AreEqual(original.ForceCondition, copied.ForceCondition);
        }

        [TestMethod]
        public void CopyConstructor_NullSource_ShouldNotThrow()
        {
            // Act
            ForceClass copied = new ForceClass(null);

            // Assert
            Assert.AreNotEqual(null, copied);
            Assert.IsFalse(copied.IsVtShmoo);
            Assert.IsFalse(copied.IsShmooInForce);
            Assert.IsFalse(copied.IsShmooInProdInst);
            Assert.IsFalse(copied.IsShmooInProdFlow);
            Assert.IsFalse(copied.IsShmooInCharInst);
            Assert.IsFalse(copied.IsShmooInCharFlow);
            Assert.IsFalse(copied.IsCz2InstName);
        }

        [TestMethod]
        public void GetAcSelector_ShouldReturnOnlyAcSelectorEntries()
        {
            // Arrange
            var force = new ForceClass
            {
                ForceCondition = "ACSelector:NV:CLK1;DCSelector:HV:PIN1;Other:XXX"
            };

            // Act
            string result = force.GetAcSelector();

            // Assert
            Assert.AreEqual("ACSelector:NV:CLK1", result);
        }

        [TestMethod]
        public void GetDcSelector_ShouldReturnOnlyDcSelectorEntries()
        {
            // Arrange
            var force = new ForceClass
            {
                ForceCondition = "DCSelector:NV:PIN1;ACSelector:HV:CLK;Other:XXX"
            };

            // Act
            string result = force.GetDcSelector();

            // Assert
            Assert.AreEqual("DCSelector:NV:PIN1", result);
        }

        [TestMethod]
        public void GetRtosTimeSet_ShouldReturnTimeSetName()
        {
            // Arrange
            var force = new ForceClass
            {
                ForceCondition = "Level:LV1;TimeSets:TS_FAST;Other:XXX"
            };

            // Act
            string result = force.GetRtosIdsTimeSet();

            // Assert
            Assert.AreEqual("TS_FAST", result);
        }

        [TestMethod]
        public void GetLevelSetting_ShouldReturnEmpty_WhenNoLevel()
        {
            ForceClass force = new ForceClass
            {
                ForceCondition = "AC:TCK:Spec1"
            };

            string result = force.GetLevelSetting();

            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetAcSetting_ShouldHandleInvalidFormats()
        {
            ForceClass force = new ForceClass
            {
                ForceCondition = "AC:INVALID;Random:XXX"
            };

            string result = force.GetAcSetting();

            Assert.AreEqual("", result);
        }

        [TestMethod]

        public void GetRtosTimeSet_ShouldReturnEmpty_WhenNoTimeSet()
        {
            ForceClass force = new ForceClass
            {
                ForceCondition = "Level:LV1"
            };

            string result = force.GetRtosIdsTimeSet();

            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetAcCategory_ShouldReturnOnlyAcPins()
        {
            ForceClass force = new ForceClass
            {
                ForceCondition = "AC:TCK,OtherPin:Spec1"
            };

            string result = force.GetAcCategory();

            Assert.IsTrue(result.Contains("TCK"));
            Assert.IsFalse(result.Contains("OtherPin"));
        }

        [TestMethod]
        public void GetPrePatForceCondition_ShouldRemoveAllKnownPatterns()
        {
            ForceClass force = new ForceClass
            {
                ForceCondition = "Level:LV1;AC:TCK:Spec1;DC:CAT1;ACSelector:NV:X;DCSelector:HV:Y;TimeSets:TS1;Custom:ABC"
            };

            string result = force.GetPrePatForceCondition();

            Assert.AreEqual("Custom:ABC", result);
        }

        [TestMethod]
        public void CorrectTypo_ShouldIgnoreInvalidFormat()
        {
            ForceClass force = new ForceClass
            {
                ForceCondition = "relay_on;NoColon"
            };

            string result = force.ForceCondition;

            Assert.IsTrue(result.Contains("relay_on"));
        }

        [TestMethod]
        public void GetAcCategory_ShouldFilterInvalidPins()
        {
            ForceClass force = new ForceClass
            {
                ForceCondition = "AC:TCK,InvalidPin:Spec1"
            };

            string result = force.GetAcCategory();

            Assert.IsTrue(result.Contains("TCK"));
        }

        [TestMethod]
        public void GetMcgSetting_ShouldReturnNonAcPins()
        {
            ForceClass force = new ForceClass
            {
                ForceCondition = "AC:TCK,MCGPIN:Spec1"
            };

            string result = force.GetMcgSetting();

            Assert.IsTrue(result.Contains("MCGPIN"));
        }

        [TestMethod]
        public void GetAcSetting_ShouldReturnEmpty_WhenInvalid()
        {
            ForceClass force = new ForceClass
            {
                ForceCondition = "AC:InvalidFormat"
            };

            string result = force.GetAcSetting();

            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetPrePatForceCondition_ShouldReturnEmpty_WhenAllFiltered()
        {
            ForceClass force = new ForceClass
            {
                ForceCondition = "Level:LV1;AC:TCK:Spec1"
            };

            string result = force.GetPrePatForceCondition();

            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void CorrectTypo_ShouldHandleMalformedInput()
        {
            ForceClass force = new ForceClass
            {
                ForceCondition = "relay_on;noColon"
            };

            string result = force.ForceCondition;

            Assert.IsTrue(result.Contains("relay_on"));
        }

        [TestMethod]
        public void GetShmoo_XShmoo_ShouldPopulateStepFields()
        {
            InitTestProgram();

            ForceClass force = new ForceClass
            {
                ForceCondition = "xshmoo(VDD:0:0,1,0.1)"
            };

            TestPlanSheet sheet = new TestPlanSheet
            {
                SheetName = "Block1_Sub1"
            };

            HardIpPattern pattern = new HardIpPattern
            {
                Pattern = new PatternClass("PAT"),
                PatternType = "Type1",
                SheetName = "Block1_Sub1"
            };

            HardipCharSetup setup = force.GetShmoo(sheet, pattern, force.ForceCondition, "SB1");

            Assert.AreNotEqual(null, setup);
            Assert.IsTrue(setup.CharSteps.Count == 1);

            CharStep step = setup.CharSteps[0];

            Assert.AreNotEqual(null, step);

            Assert.IsTrue(step.ParameterName.Length > 0);
            Assert.IsTrue(step.ParameterType.Length > 0);

            Assert.IsTrue(step.RangeFrom.Length > 0);
            Assert.IsTrue(step.RangeTo.Length > 0);
            Assert.IsTrue(step.RangeStepSize.Length > 0);

            Assert.IsTrue(step.StepName.Length > 0);
            Assert.IsTrue(step.Mode.Length > 0);

            Assert.IsTrue(step.AlgorithmName.Length > 0);
        }

        [TestMethod]
        public void GetShmoo_YShmoo_ShouldCreateYModeStep()
        {
            InitTestProgram();

            ForceClass force = new ForceClass
            {
                ForceCondition = "yshmoo(VDD:0:0,1,0.1)"
            };

            TestPlanSheet sheet = new TestPlanSheet
            {
                SheetName = "Block1_Sub1"
            };

            HardIpPattern pattern = new HardIpPattern
            {
                Pattern = new PatternClass("PAT"),
                PatternType = "Type1",
                SheetName = "Block1_Sub1"
            };

            HardipCharSetup setup = force.GetShmoo(sheet, pattern, force.ForceCondition, "SB1");

            Assert.AreNotEqual(null, setup);
            Assert.IsTrue(setup.CharSteps.Count == 1);

            CharStep step = setup.CharSteps[0];

            Assert.AreNotEqual(null, step);

            Assert.IsTrue(step.Mode.Length > 0);

            Assert.IsTrue(step.ParameterName.Length > 0);
            Assert.IsTrue(step.RangeFrom.Length > 0);
            Assert.IsTrue(step.RangeTo.Length > 0);

            Assert.IsTrue(step.AlgorithmName.Length > 0);
            Assert.AreEqual(CharSetupConst.ModeYShmoo, step.Mode);
        }

        [TestMethod]
        public void GetShmoo_Multiple_ShouldCreateTwoSteps()
        {
            InitTestProgram();

            ForceClass force = new ForceClass
            {
                ForceCondition = "xshmoo(VDD:0:0,1,0.1);yshmoo(VDD:0:0,1,0.1)"
            };

            TestPlanSheet sheet = new TestPlanSheet
            {
                SheetName = "Block1_Sub1"
            };

            HardIpPattern pattern = new HardIpPattern
            {
                Pattern = new PatternClass("PAT"),
                PatternType = "Type1",
                SheetName = "Block1_Sub1"
            };

            HardipCharSetup setup = force.GetShmoo(sheet, pattern, force.ForceCondition, "SB1");

            Assert.AreNotEqual(null, setup);
            Assert.IsTrue(setup.CharSteps.Count == 2);
        }

        [TestMethod]
        public void GetShmoo_NoShmoo_ShouldReturnNull()
        {
            InitTestProgram();

            ForceClass force = new ForceClass
            {
                ForceCondition = "Level:LV1"
            };

            TestPlanSheet sheet = new TestPlanSheet();
            HardIpPattern pattern = new HardIpPattern
            {
                Pattern = new PatternClass("PAT"),
                PatternType = "Type1",
                SheetName = "TestSheet"
            };

            HardipCharSetup setup = force.GetShmoo(sheet, pattern, force.ForceCondition, "SB1");

            Assert.AreNotEqual(null, setup);
            Assert.IsTrue(setup.CharSteps.Count == 0);
        }

        [TestMethod]
        public void GetShmoo_InvalidFormat_ShouldReturnEmptySteps()
        {
            InitTestProgram();

            ForceClass force = new ForceClass
            {
                ForceCondition = "xshmoo(INVALID)"
            };

            TestPlanSheet sheet = new TestPlanSheet
            {
                SheetName = "Block1_Sub1"
            };

            HardIpPattern pattern = new HardIpPattern
            {
                Pattern = new PatternClass("PAT"),
                PatternType = "Type1",
                SheetName = "Block1_Sub1"
            };

            HardipCharSetup setup = force.GetShmoo(sheet, pattern, force.ForceCondition, "SB1");

            Assert.AreNotEqual(null, setup);
            Assert.IsTrue(setup.CharSteps.Count == 0);
        }

        [TestMethod]
        public void GetShmoo_WithAlgorithm_ShouldSetAlgorithmArguments()
        {
            InitTestProgram();

            ForceClass force = new ForceClass
            {
                ForceCondition = "xshmoo(VDD:0:0,1,0.1:retest:Linear,5)"
            };

            TestPlanSheet sheet = new TestPlanSheet
            {
                SheetName = "Block1_Sub1"
            };

            HardIpPattern pattern = new HardIpPattern
            {
                Pattern = new PatternClass("PAT"),
                PatternType = "Type1",
                SheetName = "Block1_Sub1"
            };

            HardipCharSetup setup = force.GetShmoo(sheet, pattern, force.ForceCondition, "SB1");

            Assert.AreNotEqual(null, setup);
            Assert.IsTrue(setup.CharSteps.Count == 1);

            CharStep step = setup.CharSteps[0];

            Assert.AreEqual(CharSetupConst.AlgorithmNameLinear, step.AlgorithmName);
        }
    }
}
