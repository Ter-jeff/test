using System.Collections.Generic;

using Automation.Const;
using Automation.GenerateIgxl.BistBira.Base;
using Automation.Reader.ConfigFile.NamingRule.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.BistBira
{
    [TestClass]
    public class BistNamingTests
    {
        private BistNaming _bistNaming = null!;

        [TestInitialize]
        public void Setup()
        {
            var dummyRepairFlags = new Dictionary<string, List<KeyAndPosition>>();
            var dummySpecialRules = new Dictionary<string, MbistProductionSpecialNamingRule>();
            var dummyConfig = new MbistConfig
            {
                RepairFlagSetting = dummyRepairFlags,
                SpecialNamingRules = dummySpecialRules
            };
            _bistNaming = new BistNaming(dummyConfig);
        }

        [DataTestMethod]
        [DataRow("PLLP_BIRA_001", "AnyLabel", true, DisplayName = "01_IsBira_PatternEndsWith_BIRA")]
        [DataRow("somepattern_BIRA_FULL", "label", true, DisplayName = "02_IsBira_Contains_BIRA_FULL")]
        [DataRow("pattern_ERT", "something_BIRA", true, DisplayName = "03_IsBira_LabelContains_BIRA")]
        [DataRow("pattern_RETENTION_BIRA", "label", true, DisplayName = "04_IsBira_RETENTION_BIRA")]
        [DataRow("pattern", "label", false, DisplayName = "05_IsBira_NoBIRA_False")]
        public void IsBira_ShouldIdentifyCorrectly(string pattern, string label, bool expected)
        {
            bool result = _bistNaming.IsBira(pattern, label);
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("Vmax", MbistConst.ConNv, DisplayName = "01_GetVoltageType_Vmax_NV")]
        [DataRow("Vmin", MbistConst.ConNv, DisplayName = "02_GetVoltageType_Vmin_NV")]
        [DataRow("Vnom", MbistConst.ConNv, DisplayName = "03_GetVoltageType_Vnom_NV")]
        [DataRow("VMARGIN1", MbistConst.ConNv, DisplayName = "04_GetVoltageType_VMARGIN1_NV")]
        [DataRow("VMARGIN3", MbistConst.ConNv, DisplayName = "05_GetVoltageType_VMARGIN3_NV")]
        [DataRow("RandomVoltage", MbistConst.ConNv, DisplayName = "06_GetVoltageType_Default_NV")]
        public void GetVoltageType_ShouldReturnExpectedType(string input, string expected)
        {
            string result = BistNaming.GetVoltageType(input);
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("VMARGIN1", true, DisplayName = "01_IsNormalVoltage_VMARGIN1_True")]
        [DataRow("MHV2", true, DisplayName = "02_IsNormalVoltage_MHV2_True")]
        [DataRow("LV", true, DisplayName = "03_IsNormalVoltage_LV_True")]
        [DataRow("INVALID", false, DisplayName = "04_IsNormalVoltage_Invalid_False")]
        public void IsNormalVoltage_ShouldReturnExpected(string voltage, bool expected)
        {
            bool result = BistNaming.IsNormalVoltage(voltage);
            Assert.AreEqual(expected, result);
        }

        [DataTestMethod]
        [DataRow("Vmax", BistBinTableType.BinNv, DisplayName = "01_GetBinType_Vmax_BinNv")]
        [DataRow("Vmin", BistBinTableType.BinNv, DisplayName = "02_GetBinType_Vmin_BinNv")]
        [DataRow("Vnom", BistBinTableType.BinNv, DisplayName = "03_GetBinType_Vnom_BinNv")]
        [DataRow("VMARGIN1", BistBinTableType.BinNv, DisplayName = "04_GetBinType_VMARGIN1_BinNv")]
        [DataRow("VMARGIN3", BistBinTableType.BinNv, DisplayName = "05_GetBinType_VMARGIN3_BinNv")]
        public void GetBinType_ShouldMapCorrectly(string voltage, BistBinTableType bistBinTableType)
        {
            BistBinTableType result = _bistNaming.GetBinType(voltage);
            Assert.AreEqual(bistBinTableType, result);
        }

        [TestMethod]
        public void GetSubName_ByRule_ShouldReturnExpectedParts()
        {
            string result = _bistNaming.GetSubName("A_B_C_D", "1,3");
            Assert.AreEqual("B_D", result);
        }

        [TestMethod]
        public void GetSubName_ByIndex_ShouldReturnNthSegment()
        {
            string result = _bistNaming.GetSubName("AA_BB_CC_DD", 2);
            Assert.AreEqual("CC", result);
        }

        [TestMethod]
        public void CreateRetentionTestNameNew_ShouldBuildExpectedName()
        {
            string result = _bistNaming.CreateRetentionTestNameNew("SOC", "CAT1", "SheetX", "5", "2", "WaitOnly");
            Assert.AreEqual("SOCMbist_SheetX_CAT1_Wait5mS_STEP2_WaitOnly", result);
        }

        [TestMethod]
        public void JudgePattern_ShouldCompare8thSegment()
        {
            string pattern = "A_B_C_D_E_F_G_H_I_J";
            bool result = _bistNaming.JudgePattern(pattern, "I");
            Assert.IsTrue(result);
        }
    }
}
