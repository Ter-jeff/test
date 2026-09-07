using CommonLib.Utility.FrcCalc;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Utility.FrcCalc
{
    [TestClass]
    public class PaEngineItemTests
    {
        [TestMethod]
        public void PaEngineItem_DefaultConstructor_AllZero()
        {
            var item = new PaEngineItem();
            Assert.AreEqual(0.0, item.TargetFreq);
            Assert.AreEqual(0, item.M2);
            Assert.AreEqual(0.0, item.PatGenFreq);
            Assert.AreEqual(0, item.X2);
            Assert.AreEqual(0, item.D2);
            Assert.AreEqual(0.0, item.ClkD8Freq);
            Assert.AreEqual(0, item.M);
            Assert.AreEqual(0.0, item.PdfFreq);
            Assert.AreEqual(0, item.D1);
            Assert.AreEqual(0.0, item.PllInputFreq);
        }

        [TestMethod]
        public void PaEngineItem_SetTargetFreq_ReturnsCorrectValue()
        {
            var item = new PaEngineItem { TargetFreq = 100000000.0 };
            Assert.AreEqual(100000000.0, item.TargetFreq);
        }

        [TestMethod]
        public void PaEngineItem_SetM2_ReturnsCorrectValue()
        {
            var item = new PaEngineItem { M2 = 2 };
            Assert.AreEqual(2, item.M2);
        }

        [TestMethod]
        public void PaEngineItem_SetPatGenFreq_ReturnsCorrectValue()
        {
            var item = new PaEngineItem { PatGenFreq = 200000000.0 };
            Assert.AreEqual(200000000.0, item.PatGenFreq);
        }

        [TestMethod]
        public void PaEngineItem_SetX2_ReturnsCorrectValue()
        {
            var item = new PaEngineItem { X2 = 1 };
            Assert.AreEqual(1, item.X2);
        }

        [TestMethod]
        public void PaEngineItem_SetD2_ReturnsCorrectValue()
        {
            var item = new PaEngineItem { D2 = 5 };
            Assert.AreEqual(5, item.D2);
        }

        [TestMethod]
        public void PaEngineItem_SetPllInputFreq_ReturnsCorrectValue()
        {
            var item = new PaEngineItem { PllInputFreq = 120000000.0 };
            Assert.AreEqual(120000000.0, item.PllInputFreq);
        }

        [TestMethod]
        public void PaEngineItem_SetAllProperties_AllReturnCorrectValues()
        {
            var item = new PaEngineItem
            {
                TargetFreq = 400000000.0,
                M2 = 1,
                PatGenFreq = 400000000.0,
                X2 = 1,
                D2 = 3,
                ClkD8Freq = 150000000.0,
                M = 25,
                PdfFreq = 6000000.0,
                D1 = 20,
                PllInputFreq = 120000000.0
            };

            Assert.AreEqual(400000000.0, item.TargetFreq);
            Assert.AreEqual(1, item.M2);
            Assert.AreEqual(400000000.0, item.PatGenFreq);
            Assert.AreEqual(1, item.X2);
            Assert.AreEqual(3, item.D2);
            Assert.AreEqual(150000000.0, item.ClkD8Freq);
            Assert.AreEqual(25, item.M);
            Assert.AreEqual(6000000.0, item.PdfFreq);
            Assert.AreEqual(20, item.D1);
            Assert.AreEqual(120000000.0, item.PllInputFreq);
        }
    }
}
