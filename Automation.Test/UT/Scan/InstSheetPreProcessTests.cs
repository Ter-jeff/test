using Automation.GenerateIgxl.Scan.NonBinCut;
using Automation.Reader.ConfigFile.NamingRule.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.Test.UT.Scan
{
    [TestClass]
    public class InstSheetPreProcessTests
    {
        private InstSheetPreProcess _preProcess = null!;

        [TestInitialize]
        public void Setup()
        {
            _preProcess = new InstSheetPreProcess(new ScanConfig());
        }

        private static BinCutInstanceRow NewRow(string flowName = "")
        {
            return new BinCutInstanceRow("Sheet1") { FlowName = flowName };
        }

        #region GetPatSetName

        [TestMethod]
        public void GetPatSetName_PatSetNameOrangePresent_IgnoresInstanceAppendage()
        {
            // Arrange
            BinCutInstanceRow row = NewRow();
            row.PatSetNameOrange = "ORANGE";
            row.Instance = "INST";
            row.SplitKey = "";

            // Act
            string result = _preProcess.GetPatSetName(row, "D", "B", "M", "INIT", "PAY", "INST");

            // Assert
            Assert.AreEqual("DB_M_ORANGE", result);
        }

        [TestMethod]
        public void GetPatSetName_PatSetNameOrangeEmptyWithInstance_AppendsInstance()
        {
            // Arrange
            BinCutInstanceRow row = NewRow();
            row.PatSetNameOrange = "";
            row.SplitKey = "";

            // Act
            string result = _preProcess.GetPatSetName(row, "D", "B", "M", "INIT", "PAY", "INST");

            // Assert
            Assert.AreEqual("DB_M_INIT_PAY_INST", result);
        }

        [TestMethod]
        public void GetPatSetName_SplitKeyPresent_IncludedInPatSetName()
        {
            // Arrange
            BinCutInstanceRow row = NewRow();
            row.PatSetNameOrange = "ORANGE";
            row.SplitKey = "SPLIT";

            // Act
            string result = _preProcess.GetPatSetName(row, "D", "B", "M", "INIT", "PAY", "");

            // Assert
            Assert.AreEqual("DB_M_SPLIT_ORANGE", result);
        }

        #endregion

        #region GetPayloadName

        [TestMethod]
        public void GetPayloadName_PayloadListNonEmpty_UsesPayloadList()
        {
            // Arrange
            BinCutInstanceRow row = NewRow();
            row.PayloadList.Add("PAY1");

            // Act
            string result = _preProcess.GetPayloadName(row);

            // Assert
            Assert.AreEqual("PAY1", result);
        }

        [TestMethod]
        public void GetPayloadName_PayloadListEmpty_UsesPatternList()
        {
            // Arrange
            BinCutInstanceRow row = NewRow();
            row.PatternList.Add("PAT1");

            // Act
            string result = _preProcess.GetPayloadName(row);

            // Assert
            Assert.AreEqual("PAT1", result);
        }

        #endregion

        #region GetMode

        [TestMethod]
        public void GetMode_TokenMatchesPerformanceModePattern_ReturnsMatchingToken()
        {
            // Arrange
            BinCutInstanceRow row = NewRow("ABC MEX001 XYZ");

            // Act
            string result = _preProcess.GetMode(row);

            // Assert
            Assert.AreEqual("MEX001", result);
        }

        [TestMethod]
        public void GetMode_NoTokenMatches_ReturnsEmptyString()
        {
            // Arrange - the base class (unlike the BinCut override) falls back to "" rather
            // than the first token
            BinCutInstanceRow row = NewRow("ABC XYZ");

            // Act
            string result = _preProcess.GetMode(row);

            // Assert
            Assert.AreEqual("", result);
        }

        #endregion
    }
}
