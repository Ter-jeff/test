using CommonLib.Datalog;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Datalog
{
    [TestClass]
    public class LineBaseTests
    {
        private LineBase _lineBase;

        [TestInitialize]
        public void Initialize()
        {
            _lineBase = new LineBase();
        }

        [TestMethod]
        public void GetChannelIndex_ValidChannelFormat_ReturnsCorrectIndex()
        {
            string[] parts = ["Test", "Header", "23.x211h", "Value"];
            int index = LineBase.GetChannelIndex(parts);

            Assert.AreEqual(2, index);
        }

        [TestMethod]
        public void GetChannelIndex_NoChannelFormat_ReturnsZero()
        {
            string[] parts = ["Test", "Header", "NoMatch"];
            int index = LineBase.GetChannelIndex(parts);

            Assert.AreEqual(0, index);
        }

        [TestMethod]
        public void GetChannelIndex_EmptyArray_ReturnsZero()
        {
            string[] parts = [];
            int index = LineBase.GetChannelIndex(parts);

            Assert.AreEqual(0, index);
        }

        [TestMethod]
        public void GetPowerBinningIndex_ValidPowerBinFormat_ReturnsCorrectIndex()
        {
            string[] parts = ["Test", "Header", "10.20abc"];
            int index = LineBase.GetPowerBinningIndex(parts);

            Assert.AreEqual(2, index);
        }

        [TestMethod]
        public void GetPowerBinningIndex_NoPowerBinFormat_ReturnsZero()
        {
            string[] parts = ["Test", "Header", "NoMatch"];
            int index = LineBase.GetPowerBinningIndex(parts);

            Assert.AreEqual(0, index);
        }

        [TestMethod]
        public void GetMeasureIndex_ValidMeasureFormat_ReturnsCorrectIndex()
        {
            string[] parts = ["Test", "Header", "10.20", "100.5", "200.3", "Result"];
            int index = LineBase.GetMeasureIndex(2, parts);

            // Should find the second numeric value after index 2
            Assert.AreEqual(4, index);
        }

        [TestMethod]
        public void GetMeasureIndex_WithNAValue_ReturnsCorrectIndex()
        {
            string[] parts = ["Test", "Header", "10.20", "N/A", "200.3", "Result"];
            int index = LineBase.GetMeasureIndex(2, parts);

            // Should count N/A as numeric for measure index
            Assert.IsTrue(index > 0);
        }

        [TestMethod]
        public void GetMeasureIndex_NotEnoughValues_ReturnsZero()
        {
            string[] parts = ["Test", "100.5"];
            int index = LineBase.GetMeasureIndex(1, parts);

            Assert.AreEqual(0, index);
        }

        [TestMethod]
        public void GetSite_ValidSiteFormat_ReturnsCorrectSite()
        {
            _lineBase.Line = "[Site 5] Data";
            int site = _lineBase.GetSite();

            Assert.AreEqual(5, site);
        }

        [TestMethod]
        public void GetSite_CasInsensitive_ReturnsCorrectSite()
        {
            _lineBase.Line = "[site 10] Data";
            int site = _lineBase.GetSite();

            Assert.AreEqual(10, site);
        }

        [TestMethod]
        public void GetSite_NoSiteFormat_ReturnsMinus1()
        {
            _lineBase.Line = "No Site Here";
            int site = _lineBase.GetSite();

            Assert.AreEqual(-1, site);
        }

        [TestMethod]
        public void GetSite_InvalidSiteNumber_ReturnsMinus1()
        {
            _lineBase.Line = "[Site ABC] Data";
            int site = _lineBase.GetSite();

            Assert.AreEqual(-1, site);
        }

        [TestMethod]
        public void LineProperties_CanSetValues()
        {
            _lineBase.LineNo = 42;
            _lineBase.Line = "Test Line";

            Assert.AreEqual(42, _lineBase.LineNo);
            Assert.AreEqual("Test Line", _lineBase.Line);
        }
    }
}
