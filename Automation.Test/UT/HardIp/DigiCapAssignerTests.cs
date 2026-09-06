using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class DigiCapAssignerTests
    {
        private static List<string> InvokeGetSweepVoltageByVoltage(List<string> strList, string voltage)
        {
            return DigiCapAssigner.GetSweepVoltageByVoltage(strList, voltage);
        }

        [TestMethod]
        public void GetSweepVoltageByVoltage_NvVoltage_ExcludesLvSweepvoltageEntry()
        {
            // Act
            List<string> result = InvokeGetSweepVoltageByVoltage(["LV@sweepvoltage_X;NormalItem"], "NV");

            // Assert
            Assert.AreEqual("NormalItem", result[0]);
        }

        [TestMethod]
        public void GetSweepVoltageByVoltage_NvVoltage_ExcludesHvSweepvoltageEntry()
        {
            // Act
            List<string> result = InvokeGetSweepVoltageByVoltage(["HV@sweepvoltage_X;NormalItem"], "NV");

            // Assert
            Assert.AreEqual("NormalItem", result[0]);
        }

        [TestMethod]
        public void GetSweepVoltageByVoltage_LvVoltage_ExcludesNvAndHvSweepvoltageEntries()
        {
            // Act
            List<string> result = InvokeGetSweepVoltageByVoltage(["NV@sweepvoltage_A;HV@sweepvoltage_B;NormalItem"], "LV");

            // Assert
            Assert.AreEqual("NormalItem", result[0]);
        }

        [TestMethod]
        public void GetSweepVoltageByVoltage_HvVoltage_ExcludesNvAndLvSweepvoltageEntries()
        {
            // Act
            List<string> result = InvokeGetSweepVoltageByVoltage(["NV@sweepvoltage_A;LV@sweepvoltage_B;NormalItem"], "HV");

            // Assert
            Assert.AreEqual("NormalItem", result[0]);
        }

        [TestMethod]
        public void GetSweepVoltageByVoltage_MatchingOwnVoltagePrefix_ReplacesWithSweepvoltage()
        {
            // Act
            List<string> result = InvokeGetSweepVoltageByVoltage(["NV@sweepvoltage_X"], "NV");

            // Assert
            Assert.AreEqual("sweepvoltage_X", result[0]);
        }

        [TestMethod]
        public void GetSweepVoltageByVoltage_EmptyVoltage_NvPrefixReplacedWithSweepvoltage()
        {
            // Act
            List<string> result = InvokeGetSweepVoltageByVoltage(["NV@sweepvoltage_X"], "");

            // Assert
            Assert.AreEqual("sweepvoltage_X", result[0]);
        }

        [TestMethod]
        public void GetSweepVoltageByVoltage_EmptyVoltage_LvPrefixReplacedWithSweepvoltage()
        {
            // Act
            List<string> result = InvokeGetSweepVoltageByVoltage(["LV@sweepvoltage_X"], "");

            // Assert
            Assert.AreEqual("sweepvoltage_X", result[0]);
        }

        [TestMethod]
        public void GetSweepVoltageByVoltage_EmptyVoltage_HvPrefixReplacedWithSweepvoltage()
        {
            // Act
            List<string> result = InvokeGetSweepVoltageByVoltage(["HV@sweepvoltage_X"], "");

            // Assert
            Assert.AreEqual("sweepvoltage_X", result[0]);
        }

        [TestMethod]
        public void GetSweepVoltageByVoltage_NoMatchingPrefix_PassesThroughUnchanged()
        {
            // Act
            List<string> result = InvokeGetSweepVoltageByVoltage(["RandomText"], "NV");

            // Assert
            Assert.AreEqual("RandomText", result[0]);
        }

        [TestMethod]
        public void GetSweepVoltageByVoltage_MultipleSemicolonSeparatedSubStrings_JoinsInsideEntries()
        {
            // Act
            List<string> result = InvokeGetSweepVoltageByVoltage(["A;B"], "NV");

            // Assert
            Assert.AreEqual("A;B", result[0]);
        }

        [TestMethod]
        public void GetSweepVoltageByVoltage_MultipleListEntries_ReturnsOneResultPerEntry()
        {
            // Act
            List<string> result = InvokeGetSweepVoltageByVoltage(["A", "B"], "NV");

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("A", result[0]);
            Assert.AreEqual("B", result[1]);
        }
    }
}
