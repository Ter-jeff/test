using CommonLib.Datalog;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.Datalog
{
    [TestClass]
    public class LimitLineTests
    {
        private LimitLine _limitLine;

        [TestInitialize]
        public void Initialize()
        {
            _limitLine = new LimitLine();
        }

        [TestMethod]
        public void GetTestName_ValidLine_ReturnsTestName()
        {
            _limitLine.Line = "50138538 0 IDS VDD_SOC";
            string result = _limitLine.GetTestName();
            Assert.AreEqual("IDS", result);
        }

        [TestMethod]
        public void GetTestName_NoTestName_ReturnsEmpty()
        {
            _limitLine.Line = "OnlyTwoElements";
            string result = _limitLine.GetTestName();
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetTestName_ConvertsCaseToUpper()
        {
            _limitLine.Line = "50138538 0 ids VDD_SOC";
            string result = _limitLine.GetTestName();
            Assert.AreEqual("IDS", result);
        }

        [TestMethod]
        public void GetTestName_WithSpecialFormatA_ReturnsTestName()
        {
            _limitLine.Line = "50138538 0 (A)CURRENT VDD_SOC";
            string result = _limitLine.GetTestName();
            Assert.AreEqual("CURRENT", result);
        }

        [TestMethod]
        public void RegEfuseLine_ValidEfuseLine_Matches()
        {
            string line = "Site(0) EFUSE Write Values";
            Assert.IsTrue(LimitLine.RegEfuseLine.IsMatch(line));
        }

        [TestMethod]
        public void RegEfuseLine_ReadEfuseLine_Matches()
        {
            string line = "Site(5) EFUSE Read Values";
            Assert.IsTrue(LimitLine.RegEfuseLine.IsMatch(line));
        }

        [TestMethod]
        public void RegEfuseLine_CaseInsensitive_Matches()
        {
            string line = "Site(0) efuse write values";
            Assert.IsTrue(LimitLine.RegEfuseLine.IsMatch(line));
        }

        [TestMethod]
        public void RegEfuseLine_InvalidLine_DoesNotMatch()
        {
            string line = "Not an efuse line";
            Assert.IsFalse(LimitLine.RegEfuseLine.IsMatch(line));
        }

        [TestMethod]
        public void GetSiteData_ValidLineWithSite_ParsesSiteAndValue()
        {
            // channelIndex = 4 ("11.x404"), measureIndex = 6 ("26.0000"), unit = "mA"
            _limitLine.Line = "50138538 0 IDS VDD_SOC 11.x404 0.0000 26.0000 mA";
            _limitLine.GetSiteData(out int site, out double log);

            Assert.AreEqual(0, site);
            Assert.AreEqual(26.0000, log);
        }

        [TestMethod]
        public void GetSiteData_InvalidLine_ReturnsMinus1()
        {
            _limitLine.Line = "InvalidLine";
            _limitLine.GetSiteData(out int site, out double log);
            Assert.AreEqual(-1, site);
            Assert.AreEqual(-1, log);
        }

        [TestMethod]
        public void GetSiteData_WithVoltUnit_MultipliesValue()
        {
            _limitLine.Line = "50138538 2 IDS VDD_SOC 11.x315 0.0000 1.5 V";
            _limitLine.GetSiteData(out int site, out double log);

            Assert.AreEqual(2, site);
            // 1.5 * 1000.0
            Assert.AreEqual(1500.0, log);
        }

        [TestMethod]
        public void GetSiteData_WithMicroAmpUnit_DividesValue()
        {
            _limitLine.Line = "50138538 1 IDS VDD_SOC 11.x315 0.0000 450.0 uA";
            _limitLine.GetSiteData(out int site, out double log);

            Assert.AreEqual(1, site);
            // 450.0 / 1000.0
            Assert.AreEqual(0.45, log);
        }

        [TestMethod]
        public void GetSiteOnly_ValidLine_ParsesCorrectly()
        {
            _limitLine.Line = "50138538 3 IDS VDD_SOC";
            _limitLine.GetSiteOnly(out int site);
            Assert.AreEqual(3, site);
        }

        [TestMethod]
        public void GetSiteOnly_MalformedLine_ReturnsMinus1()
        {
            _limitLine.Line = "50138538 ABC IDS VDD_SOC";
            _limitLine.GetSiteOnly(out int site);
            Assert.AreEqual(-1, site);
        }

        [TestMethod]
        public void IsFail_LineContainsFailMarker_ReturnsTrue()
        {
            _limitLine.Line = "50138538 0 IDS VDD_SOC 11.x404 0.0000 (F) 26.0000 mA";
            Assert.IsTrue(_limitLine.IsFail());
        }

        [TestMethod]
        public void IsFail_LineDoesNotContainMarker_ReturnsFalse()
        {
            _limitLine.Line = "50138538 0 IDS VDD_SOC 11.x404 0.0000 26.0000 mA";
            Assert.IsFalse(_limitLine.IsFail());
        }

        [TestMethod]
        public void ToRow_InfoOrEfuseOrEmpty_ReturnsNull()
        {
            _limitLine.Line = "[INFO] Starting collection sequence";
            Assert.IsNull(_limitLine.ToRow());

            _limitLine.Line = "Site(0) EFUSE Write Values";
            Assert.IsNull(_limitLine.ToRow());

            _limitLine.Line = "";
            Assert.IsNull(_limitLine.ToRow());
        }

        [TestMethod]
        public void ToRow_StandardValidLine_PopulatesLimitRowData()
        {
            // Indexing matching step targets:
            // spt: [0]:"50138538", [1]:"0", [2]:"IDS", [3]:"VDD_SOC", [4]:"11.x404", [5]:"0.0000", [6]:"26.0000", [7]:"103.2000"
            // regexIdx = 4 ("11.x404"). step is found when moreTwoStep == 2 -> step = 6 ("26.0000")
            // value1 (Measured) = spt[6] -> 26.0000
            // value2 (LowLimit) = spt[5] -> 0.0000
            // value3 (HighLimit) = spt[7] -> 103.2000
            _limitLine.Line = "50138538 0 IDS VDD_SOC 11.x404 0.0000 26.0000 103.2000";
            LimitRow row = _limitLine.ToRow();

            Assert.IsNotNull(row);
            Assert.AreEqual("50138538", row.Number);
            Assert.AreEqual(0, row.Site);
            Assert.AreEqual("IDS", row.TestName);
            Assert.AreEqual("VDD_SOC", row.Pin);
            Assert.AreEqual(26.0000, row.Measured);
            Assert.AreEqual(0.0000, row.LowLimit);
            Assert.AreEqual(103.2000, row.HighLimit);
        }

        [TestMethod]
        public void GetExecuteUnitVal_WithMegaUnitModifier_ScalesLimits()
        {
            string line = "Some logs containing M data";
            string[] spt = ["Dummy", "5.0", "12.5", "25.0"];
            // step-1 = "5.0", step+1 = "25.0"
            int step = 2;

            LimitLine.GetExecuteUnitVal(line, spt, step, out double low, out double high);

            // 5.0 * 1,000,000
            Assert.AreEqual(5000000.0, low);
            // 25.0 * 1,000,000
            Assert.AreEqual(25000000.0, high);
        }

        [TestMethod]
        public void GetExecuteUnitVal_NoUnitModifier_KeepsBaseValues()
        {
            string line = "Standard text format line";
            string[] spt = ["Dummy", "10.0", "20.0", "30.0"];
            int step = 2;

            LimitLine.GetExecuteUnitVal(line, spt, step, out double low, out double high);

            Assert.AreEqual(10.0, low);
            Assert.AreEqual(30.0, high);
        }

        [TestMethod]
        public void RegexSite_ValidSite_Matches()
        {
            Assert.IsTrue(LineBase.RegexSite.IsMatch("[Site 5]"));
            Assert.IsTrue(LineBase.RegexSite.IsMatch("[site 10]"));
        }

        [TestMethod]
        public void RegexSite_InvalidSite_DoesNotMatch()
        {
            Assert.IsFalse(LineBase.RegexSite.IsMatch("[NoSite]"));
        }

        [TestMethod]
        public void RegexChannel_ValidChannel_Matches()
        {
            Assert.IsTrue(LineBase.RegexChannel.IsMatch("23.x211h"));
            Assert.IsTrue(LineBase.RegexChannel.IsMatch("5.a1b"));
        }

        [TestMethod]
        public void RegexChannel_InvalidChannel_DoesNotMatch()
        {
            Assert.IsFalse(LineBase.RegexChannel.IsMatch("abc.xyz"));
        }

        [TestMethod]
        public void RegexPowerBin_ValidPowerBin_Matches()
        {
            Assert.IsTrue(LineBase.RegexPowerBin.IsMatch("10.20abc"));
            Assert.IsTrue(LineBase.RegexPowerBin.IsMatch("5.5"));
        }

        [TestMethod]
        public void RegexPowerBin_InvalidPowerBin_DoesNotMatch()
        {
            Assert.IsFalse(LineBase.RegexPowerBin.IsMatch("abc.xyz"));
        }
    }
}
