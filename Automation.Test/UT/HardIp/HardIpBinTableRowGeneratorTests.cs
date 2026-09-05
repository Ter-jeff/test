using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTableRow;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.Static;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardIpBinTableRowGeneratorTests : FunctionTestBase
    {
        private static HardIpBinTableRowGenerator _generator = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            _generator = new HardIpBinTableRowGenerator("Sheet1_X", []);
        }

        [TestInitialize]
        public void Setup()
        {
            _generator.SheetName = "Sheet1_X";
            _generator.BlockName = "BLK";
            _generator.SubBlockName = "SUB";
            _generator.TimingAc = "";
            _generator.InstNameSubStr = "";
            _generator.NoPattern = false;
            _generator.IsHipEfuseRead = false;
            _generator.Pattern = new HardIpPattern { Pattern = new PatternClass("PatX") };
            LocalSpecs.Options.Device = EnumDevice.AP;
        }

        #region CreateHardIpItemList

        [TestMethod]
        public void CreateHardIpItemList_IsHipEfuseRead_ReturnsEfuseReadFailAction()
        {
            _generator.IsHipEfuseRead = true;

            string result = _generator.CreateHardIpItemList("");

            Assert.AreEqual("F_HARDIP_Fuse_Read_Non_Zero_Check_Flag", result);
        }

        [TestMethod]
        public void CreateHardIpItemList_EmptyVoltage_ReturnsAllThreeFlagsJoined()
        {
            string result = _generator.CreateHardIpItemList("");

            Assert.AreEqual(
                "F_Sheet1_BLK_SUB_patx_N_Flag,F_Sheet1_BLK_SUB_patx_H_Flag,F_Sheet1_BLK_SUB_patx_L_Flag",
                result);
        }

        [TestMethod]
        public void CreateHardIpItemList_NvVoltage_ReturnsOnlyNFlag()
        {
            string result = _generator.CreateHardIpItemList("NV");

            Assert.AreEqual("F_Sheet1_BLK_SUB_patx_N_Flag", result);
        }

        [TestMethod]
        public void CreateHardIpItemList_HvVoltage_ReturnsOnlyHFlag()
        {
            string result = _generator.CreateHardIpItemList("HV");

            Assert.AreEqual("F_Sheet1_BLK_SUB_patx_H_Flag", result);
        }

        [TestMethod]
        public void CreateHardIpItemList_LvVoltage_ReturnsOnlyLFlag()
        {
            string result = _generator.CreateHardIpItemList("LV");

            Assert.AreEqual("F_Sheet1_BLK_SUB_patx_L_Flag", result);
        }

        [TestMethod]
        public void CreateHardIpItemList_UserDefinedFailFlags_OverrideDefaultsForMatchingVoltagesOnly()
        {
            _generator.Pattern = new HardIpPattern
            {
                Pattern = new PatternClass("PatX"),
                Failflag = "NV@CustomFlagN;HV@CustomFlagH"
            };

            string result = _generator.CreateHardIpItemList("");

            Assert.AreEqual("CustomFlagN,CustomFlagH,F_Sheet1_BLK_SUB_patx_L_Flag", result);
        }

        [TestMethod]
        public void CreateHardIpItemList_RfDevice_UsesWirelessFailFlagFormat()
        {
            LocalSpecs.Options.Device = EnumDevice.RF;

            string result = _generator.CreateHardIpItemList("NV");

            LocalSpecs.Options.Device = EnumDevice.AP;

            Assert.AreEqual("F_BLK_SUB_N_Flag", result);
        }

        [TestMethod]
        public void CreateHardIpItemList_SubBlockNameContainsMergeSuffix_StripsMergeSuffix()
        {
            _generator.SubBlockName = "SUB-MERGE";

            string result = _generator.CreateHardIpItemList("");

            Assert.AreEqual(
                "F_Sheet1_BLK_SUB_patx_N_Flag,F_Sheet1_BLK_SUB_patx_H_Flag,F_Sheet1_BLK_SUB_patx_L_Flag",
                result);
        }

        [TestMethod]
        public void CreateHardIpItemList_RfDeviceWithMergeSuffix_UsesWirelessFormatAndStripsMergeSuffix()
        {
            _generator.SubBlockName = "SUB-MERGE";
            LocalSpecs.Options.Device = EnumDevice.RF;

            string result = _generator.CreateHardIpItemList("");

            LocalSpecs.Options.Device = EnumDevice.AP;

            Assert.AreEqual("F_BLK_SUB_N_Flag,F_BLK_SUB_H_Flag,F_BLK_SUB_L_Flag", result);
        }

        #endregion

        #region CreateHardIpName

        [TestMethod]
        public void CreateHardIpName_IsHipEfuseRead_ReturnsEfuseReadBinParameter()
        {
            _generator.IsHipEfuseRead = true;

            string result = _generator.CreateHardIpName();

            Assert.AreEqual("Bin_Sheet1_Fuse_Read_Non_Zero_Check", result);
        }

        [TestMethod]
        public void CreateHardIpName_NotHipEfuseRead_ReturnsFlowBinParameter()
        {
            string result = _generator.CreateHardIpName();

            Assert.AreEqual("Bin_BLK_SUB", result);
        }

        [TestMethod]
        public void CreateHardIpName_SubBlockNameContainsMergeSuffix_StripsMergeSuffix()
        {
            _generator.SubBlockName = "SUB-MERGE";

            string result = _generator.CreateHardIpName();

            Assert.AreEqual("Bin_BLK_SUB", result);
        }

        #endregion

        #region CreateHardIpItems

        [TestMethod]
        public void CreateHardIpItems_Hv_ReturnsTInMiddle()
        {
            List<string> result = _generator.CreateHardIpItems("HV");

            CollectionAssert.AreEqual(new List<string> { "", "T", "" }, result);
        }

        [TestMethod]
        public void CreateHardIpItems_Lv_ReturnsTInLast()
        {
            List<string> result = _generator.CreateHardIpItems("LV");

            CollectionAssert.AreEqual(new List<string> { "", "", "T" }, result);
        }

        [TestMethod]
        public void CreateHardIpItems_Nv_ReturnsTInFirst()
        {
            List<string> result = _generator.CreateHardIpItems("NV");

            CollectionAssert.AreEqual(new List<string> { "T", "", "" }, result);
        }

        [TestMethod]
        public void CreateHardIpItems_Default_ReturnsAllT()
        {
            List<string> result = _generator.CreateHardIpItems("");

            CollectionAssert.AreEqual(new List<string> { "T", "T", "T" }, result);
        }

        #endregion
    }
}
