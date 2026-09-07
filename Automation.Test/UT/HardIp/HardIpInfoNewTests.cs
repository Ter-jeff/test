using Automation.GenerateIgxl.HardIp.InputObject;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardIpInfoNewTests
    {
        [TestMethod]
        public void ExtractMeasSeqInfo_Should_HandleEmptyMeasSeq()
        {
            var info = new HardIpInfoNew
            {
                MeasSeq = "",
                MeasPin = "",
                ForcePin = "",
                ForceType = "",
                ForceValue = "",
                ExpectValue = "",
                ExpectFreq = "",
                HLimit = "",
                LLimit = "",
                MeasWait = "",
                RfSetup = "",
                MeasName = "",
                MeasStoreName = "",
                CalcEquation = "",
                CalcStoreName = ""
            };

            info.ExtractMeasSeqInfo();

            Assert.AreNotEqual(null, info.SeqInfo);
            Assert.AreEqual(0, info.SeqInfo.Count);
        }

        [TestMethod]
        public void ExtractMeasSeqInfo_Should_ParseSimpleSeq()
        {
            var info = new HardIpInfoNew
            {
                MeasSeq = "WISRC|WIMEAS",
                MeasPin = "PIN_A|PIN_B",
                ForcePin = "FORCE_A|FORCE_B",
                ForceType = "I|V",
                ForceValue = "0.1|0.2",
                ExpectValue = "1|2",
                ExpectFreq = "2.2GHz|1GHz",
                HLimit = "10|20",
                LLimit = "1|2",
                MeasWait = "5|10",
                RfSetup = "SETUP1|SETUP2",
                MeasName = "MEAS1|MEAS2",
                MeasStoreName = "STORE1|STORE2",
                CalcEquation = "",
                CalcStoreName = ""
            };

            info.ExtractMeasSeqInfo();

            Assert.AreNotEqual(null, info.SeqInfo);
            Assert.AreEqual(2, info.SeqInfo.Count);

            HardIpSeqInfoNew first = info.SeqInfo[0];
            Assert.AreEqual("WISRC", first.MeasSeq);
            CollectionAssert.Contains(first.MeasPin, "PIN_A");
            CollectionAssert.Contains(first.ForcePin, "FORCE_A");
            CollectionAssert.Contains(first.ForceType, "I");
            CollectionAssert.Contains(first.ForceValue, "0.1");
            CollectionAssert.Contains(first.ExpectValue, "1");
            CollectionAssert.Contains(first.ExpectFreq, "2.2GHz");

            HardIpSeqInfoNew second = info.SeqInfo[1];
            Assert.AreEqual("WIMEAS", second.MeasSeq);
            CollectionAssert.Contains(second.MeasPin, "PIN_B");
            CollectionAssert.Contains(second.ForcePin, "FORCE_B");
            CollectionAssert.Contains(second.ForceType, "V");
            CollectionAssert.Contains(second.ForceValue, "0.2");
            CollectionAssert.Contains(second.ExpectValue, "2");
            CollectionAssert.Contains(second.ExpectFreq, "1GHz");
        }

        [TestMethod]
        public void ExtractMeasSeqInfo_Should_HandleMeasWait()
        {
            var info = new HardIpInfoNew
            {
                MeasSeq = "MeasWait",
                MeasWait = "10",
                MeasPin = ""
            };

            info.ExtractMeasSeqInfo();

            Assert.AreEqual(1, info.SeqInfo.Count);
            HardIpSeqInfoNew seq = info.SeqInfo[0];
            Assert.AreEqual("N", seq.MeasSeq);
            Assert.AreEqual("10", seq.MeasPins[0].MeasWaitTime);
        }

        [TestMethod]
        public void ExtractMeasSeqInfo_Should_HandleCalcEquation()
        {
            var info = new HardIpInfoNew
            {
                MeasSeq = "CALC",
                MeasPin = "",
                CalcEquation = "EQ1>1>10|EQ2>2>20",
                CalcStoreName = "STORE1|STORE2",
                MeasName = "M1>M2>M3",
                HLimit = "10>20>30",
                LLimit = "1>2>3"
            };

            info.ExtractMeasSeqInfo();

            Assert.AreEqual(1, info.SeqInfo.Count);
            HardIpSeqInfoNew seq = info.SeqInfo[0];
            Assert.AreEqual(3, seq.Calc.Count);

            MeasPin calc1 = seq.Calc[0];
            Assert.AreEqual("EQ1", calc1.CalcEqn);
            Assert.AreEqual("1", calc1.LowLimit);
            Assert.AreEqual("10", calc1.HighLimit);
            Assert.AreEqual("STORE1", calc1.CusStr);

            MeasPin calc2 = seq.Calc[1];
            Assert.AreEqual("1", calc2.CalcEqn);
            Assert.AreEqual("2", calc2.LowLimit);
            Assert.AreEqual("20", calc2.HighLimit);
            Assert.AreEqual("", calc2.CusStr);
        }
    }
}
