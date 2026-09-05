using Automation.GenerateIgxl.HardIp.InputObject;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class SweepCodeTests
    {
        [TestMethod]
        [DataRow("0,10,1", 0, 10, 1)]         // Standard common
        [DataRow("5,20", 5, 20, 0)]           // Missing step
        [DataRow("-5,5,2", -5, 5, 2)]         // Negative start
        [DataRow("0,10,NON_INT", 0, 10, 1)]   // Step is string (should return 1)
        public void CommonType_Parsing_ReturnsCorrectValues(string info, int start, int end, int step)
        {
            // Arrange
            var sweep = new SweepCode { Type = SweepCode.SweepType.Common, SweepInfo = info };

            // Assert
            Assert.AreEqual(start, sweep.Start, "Start mismatch");
            Assert.AreEqual(end, sweep.End, "End mismatch");
            Assert.AreEqual(step, sweep.Step, "Step mismatch");
        }

        [TestMethod]
        public void CustomType_Parsing_ReturnsSpecificValues()
        {
            // Arrange
            var sweep = new SweepCode
            {
                Type = SweepCode.SweepType.Custom,
                SweepInfo = "Val1,Val2,Val3"
            };

            // Assert
            Assert.AreEqual(0, sweep.Start, "Custom Start should always be 0");
            Assert.AreEqual(3, sweep.Step, "Custom Step should be the count of elements");
            Assert.AreEqual("", sweep.StepMisc, "Custom StepMisc should be empty");
        }

        [TestMethod]
        [DataRow("0,10,1,graycode", "Bit1:8:0:1:BinToGray")]
        [DataRow("0,10,1,graycode_sign", "Bit1:8:0:1:BinToGray_Sign")]
        [DataRow("0,10,1,Normal", "Bit1:8:0:1:Normal")]
        public void GetFlowForLoopInfo_CommonType_ReturnsCorrectFormat(string info, string expected)
        {
            // Arrange
            var sweep = new SweepCode
            {
                Type = SweepCode.SweepType.Common,
                SendBitName = "Bit1",
                Width = 8,
                SweepInfo = info
            };

            // Act
            string result = sweep.GetFlowForLoopInfo(isCsUsing: false);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetFlowForLoopInfo_CustomType_ReturnsCorrectFormat()
        {
            var sweep = new SweepCode
            {
                Type = SweepCode.SweepType.Custom,
                SendBitName = "Bit1",
                Width = 4,
                SweepInfo = "1,2,4,8"
            };

            string result = sweep.GetFlowForLoopInfo();

            Assert.AreEqual("Bit1:[1,2,4,8]:4;", result);
        }

        [TestMethod]
        public void GetFlowForLoopInfo_CommonType_IsCsUsingTrue_ShouldNotIncludeRange()
        {
            var sweep = new SweepCode
            {
                Type = SweepCode.SweepType.Common,
                SendBitName = "Bit1",
                Width = 8,
                SweepInfo = "0,10,1,Normal"
            };

            string result = sweep.GetFlowForLoopInfo(true);

            Assert.AreEqual("Bit1:8:Normal", result);
        }

        [TestMethod]
        public void GetFlowForLoopInfo_CommonType_NoMisc_ShouldNotAppendMisc()
        {
            var sweep = new SweepCode
            {
                Type = SweepCode.SweepType.Common,
                SendBitName = "Bit1",
                Width = 8,
                SweepInfo = "0,10,1"
            };

            string result = sweep.GetFlowForLoopInfo(false);

            Assert.AreEqual("Bit1:8:0:1", result);
        }

        [TestMethod]
        public void GetFlowForLoopInfo_CommonType_MiscCaseInsensitive_ShouldWork()
        {
            var sweep = new SweepCode
            {
                Type = SweepCode.SweepType.Common,
                SendBitName = "Bit1",
                Width = 8,
                SweepInfo = "0,10,1,GrAyCoDe"
            };

            string result = sweep.GetFlowForLoopInfo(false);

            Assert.AreEqual("Bit1:8:0:1:BinToGray", result);
        }

        [TestMethod]
        public void GetFlowForLoopInfo_CustomType_ShouldIgnoreIsCsUsing()
        {
            var sweep = new SweepCode
            {
                Type = SweepCode.SweepType.Custom,
                SendBitName = "Bit1",
                Width = 4,
                SweepInfo = "1,2"
            };

            string result = sweep.GetFlowForLoopInfo(true);

            Assert.AreEqual("Bit1:[1,2]:4;", result);
        }

        [TestMethod]
        public void GetFlowForLoopInfo_CustomType_EmptySweepInfo()
        {
            var sweep = new SweepCode
            {
                Type = SweepCode.SweepType.Custom,
                SendBitName = "Bit1",
                Width = 4,
                SweepInfo = ""
            };

            string result = sweep.GetFlowForLoopInfo();

            Assert.AreEqual("Bit1:[]:4;", result);
        }

        [TestMethod]
        public void Copy_ReturnsDeepCopy()
        {
            // Arrange
            var original = new SweepCode { SendBitName = "Original", Width = 10 };

            // Act
            SweepCode copy = original.Copy();
            copy.SendBitName = "Changed";

            // Assert
            Assert.AreNotEqual(original.SendBitName, copy.SendBitName);
            Assert.AreEqual(original.Width, copy.Width);
        }
    }

}
