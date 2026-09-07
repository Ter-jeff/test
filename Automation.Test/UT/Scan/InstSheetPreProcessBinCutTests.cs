using System.Collections.Generic;

using Automation.GenerateIgxl.Scan.NonBinCut;
using Automation.Reader.ConfigFile.NamingRule.Base;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.Test.UT.Scan
{
    [TestClass]
    public class InstSheetPreProcessBinCutTests
    {
        private InstSheetPreProcessBinCut _preProcess = null!;

        [TestInitialize]
        public void Setup()
        {
            _preProcess = new InstSheetPreProcessBinCut(new ScanConfig());
        }

        private static BinCutInstanceRow NewRow(string flowName = "", List<string> payloadList = null, List<string> initList = null)
        {
            var row = new BinCutInstanceRow("Sheet1") { FlowName = flowName };
            if (payloadList != null)
            {
                row.PayloadList.AddRange(payloadList);
            }
            if (initList != null)
            {
                row.InitList.AddRange(initList);
            }
            return row;
        }

        #region GetBlockByPattern

        [TestMethod]
        public void GetBlockByPattern_ContainsScToken_ReturnsTd()
        {
            // Act
            string result = _preProcess.GetBlockByPattern("PP_BRNA0_SC_CFXX_SAA");

            // Assert
            Assert.AreEqual("Td", result);
        }

        [TestMethod]
        public void GetBlockByPattern_ContainsBiToken_ReturnsMbist()
        {
            // Act
            string result = _preProcess.GetBlockByPattern("FA_CFUA0_BI_P001_SNE");

            // Assert
            Assert.AreEqual("Mbist", result);
        }

        [TestMethod]
        public void GetBlockByPattern_NoMatchingToken_ReturnsEmpty()
        {
            // Act
            string result = _preProcess.GetBlockByPattern("SOME_OTHER_PATTERN");

            // Assert
            Assert.AreEqual("", result);
        }

        #endregion

        #region GetBlock

        [TestMethod]
        public void GetBlock_PayloadListWithScPattern_ReturnsTd()
        {
            // Arrange
            BinCutInstanceRow row = NewRow(payloadList: ["PP_BRNA0_SC_CFXX_SAA"]);

            // Act
            string result = _preProcess.GetBlock(row);

            // Assert
            Assert.AreEqual("Td", result);
        }

        [TestMethod]
        public void GetBlock_EmptyPayloadListInitListWithBiPattern_ReturnsMbist()
        {
            // Arrange
            BinCutInstanceRow row = NewRow(initList: ["FA_CFUA0_BI_P001_SNE"]);

            // Act
            string result = _preProcess.GetBlock(row);

            // Assert
            Assert.AreEqual("Mbist", result);
        }

        [TestMethod]
        public void GetBlock_NoPayloadNoInit_FallsBackToFlowNameClassification()
        {
            // Arrange - no PayloadList/InitList, but FlowName has multiple underscore-separated
            // tokens and contains "TD", so the flow-name fallback classifies it as "Td"
            BinCutInstanceRow row = NewRow(flowName: "TD_Flow_Extra");

            // Act
            string result = _preProcess.GetBlock(row);

            // Assert
            Assert.AreEqual("Td", result);
        }

        [TestMethod]
        public void GetBlock_NoPayloadNoInitSingleTokenFlowName_ReturnsEmpty()
        {
            // Arrange - FlowName has no spaces or underscores, so the fallback condition
            // (Split(' ').Length>1 || Split('_').Length>1) is never satisfied
            BinCutInstanceRow row = NewRow(flowName: "SingleToken");

            // Act
            string result = _preProcess.GetBlock(row);

            // Assert
            Assert.AreEqual("", result);
        }

        #endregion

        #region GetPayloadNameNew

        [TestMethod]
        public void GetPayloadNameNew_MbistBlockSinglePattern_AppendsSingleSuffix()
        {
            // Arrange - a single non-Unknown, non-Init pattern under "Mbist" block
            BinCutInstanceRow row = NewRow(payloadList: ["FA_CFUA0_S_PL01_BI_P001_SNE_JTG_UNS_ALLFRV_SI_SRVA"]);

            // Act
            string result = _preProcess.GetPayloadNameNew(row, "Mbist");

            // Assert
            StringAssert.EndsWith(result, "_SINGLE");
        }

        [TestMethod]
        public void GetPayloadNameNew_MbistBlockMultiplePatterns_DoesNotAppendSingleSuffix()
        {
            // Arrange - two non-Unknown, non-Init patterns
            BinCutInstanceRow row = NewRow(payloadList:
            [
                "FA_CFUA0_S_PL01_BI_P001_SNE_JTG_UNS_ALLFRV_SI_SRVA",
                "FA_CFUA0_S_PL02_BI_P002_SNE_JTG_UNS_ALLFRV_SI_SRVB"
            ]);

            // Act
            string result = _preProcess.GetPayloadNameNew(row, "Mbist");

            // Assert
            Assert.IsFalse(result.EndsWith("_SINGLE"));
        }

        #endregion

        #region GetMode

        [TestMethod]
        public void GetMode_TokenMatchesPerformanceModePattern_ReturnsMatchingToken()
        {
            // Arrange
            BinCutInstanceRow row = NewRow(flowName: "ABC MEX001 XYZ");

            // Act
            string result = _preProcess.GetMode(row);

            // Assert
            Assert.AreEqual("MEX001", result);
        }

        [TestMethod]
        public void GetMode_NoTokenMatches_FallsBackToFirstToken()
        {
            // Arrange - unlike the base class (which falls back to ""), the BinCut override
            // falls back to the first whitespace-separated token when no regex match is found
            BinCutInstanceRow row = NewRow(flowName: "ABC XYZ");

            // Act
            string result = _preProcess.GetMode(row);

            // Assert
            Assert.AreEqual("ABC", result);
        }

        #endregion
    }
}
