using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.Static;

using CommonLib.ErrorReport;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class PatInfoCheckerTests
    {
        private readonly PatInfoChecker _checker = new();

        [TestInitialize]
        public void Setup()
        {
            ErrorReportManager.ClearErrors();
            LocalSpecs.HardIpInfos = [];
        }

        [TestCleanup]
        public void Cleanup()
        {
            ErrorReportManager.ClearErrors();
            LocalSpecs.HardIpInfos = [];
        }

        private static HardIpPattern NewPattern(string realPatternName, string sheetName = "S1", int rowNum = 0, string miscInfo = "")
        {
            return new HardIpPattern
            {
                SheetName = sheetName,
                RowNum = rowNum,
                MiscInfo = miscInfo,
                Pattern = new PatternClass(realPatternName)
            };
        }

        private static Dictionary<string, HardIpSheet> NewPlanDic(params HardIpPattern[] patterns)
        {
            var sheet = new HardIpSheet { SheetName = "S1", Rows = [.. patterns] };
            return new Dictionary<string, HardIpSheet> { { "S1", sheet } };
        }

        private static void SeedHardIpInfo(HardIpInfo info)
        {
            LocalSpecs.HardIpInfos = new HardIpInfos(info);
        }

        #region CheckPatInfo

        [TestMethod]
        public void CheckPatInfo_PatternNotInPatInfo_AddsMissingPatternError()
        {
            // Arrange - LocalSpecs.HardIpInfos is empty, so "PAT1" is never found
            Dictionary<string, HardIpSheet> planDic = NewPlanDic(NewPattern("PAT1"));

            // Act
            _checker.CheckPatInfo(planDic);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfo_PatternExistsInPatInfo_NoMissingPatternError()
        {
            // Arrange
            SeedHardIpInfo(new HardIpInfo { Payload = "PAT1", PatInfoExist = true, UseThisVersion = true });
            Dictionary<string, HardIpSheet> planDic = NewPlanDic(NewPattern("PAT1"));

            // Act
            _checker.CheckPatInfo(planDic);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfo_MismatchedDigSrcSignalNameAndVmVector_AddsError()
        {
            // Arrange
            SeedHardIpInfo(new HardIpInfo
            {
                Payload = "PAT1",
                PatInfoExist = true,
                UseThisVersion = true,
                SendBitStr = "A_1+B_2",
                SendBitName = "X+Y",
                DigSrcSignalName = "SIG1",
                VmVector = "SIG2"
            });
            Dictionary<string, HardIpSheet> planDic = NewPlanDic(NewPattern("PAT1"));

            // Act
            _checker.CheckPatInfo(planDic);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfo_MatchedDigSrcSignalNameAndVmVector_NoError()
        {
            // Arrange
            SeedHardIpInfo(new HardIpInfo
            {
                Payload = "PAT1",
                PatInfoExist = true,
                UseThisVersion = true,
                SendBitStr = "A_1+B_2",
                SendBitName = "X+Y",
                DigSrcSignalName = "SIG1",
                VmVector = "SIG1"
            });
            Dictionary<string, HardIpSheet> planDic = NewPlanDic(NewPattern("PAT1"));

            // Act
            _checker.CheckPatInfo(planDic);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfo_DuplicateSendBitNamePair_AddsWrongSendInformationError()
        {
            // Arrange - both send-bit segments produce the same "name&prefix" pair
            SeedHardIpInfo(new HardIpInfo
            {
                Payload = "PAT1",
                PatInfoExist = true,
                UseThisVersion = true,
                SendBitStr = "A_1+A_2",
                SendBitName = "X+X",
                DigSrcSignalName = "SIG1",
                VmVector = "SIG1"
            });
            Dictionary<string, HardIpSheet> planDic = NewPlanDic(NewPattern("PAT1"));

            // Act
            _checker.CheckPatInfo(planDic);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfo_MeasSeqPresentCallSubrsEmpty_AddsMissingCallSubroutineError()
        {
            // Arrange
            SeedHardIpInfo(new HardIpInfo
            {
                Payload = "PAT1",
                PatInfoExist = true,
                UseThisVersion = true,
                MeasSeqStr = "Seq1",
                CallSubrs = ""
            });
            Dictionary<string, HardIpSheet> planDic = NewPlanDic(NewPattern("PAT1"));

            // Act
            _checker.CheckPatInfo(planDic);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfo_MeasSeqEmptyCallSubrsPresent_AddsMissingMeasureSequenceError()
        {
            // Arrange
            SeedHardIpInfo(new HardIpInfo
            {
                Payload = "PAT1",
                PatInfoExist = true,
                UseThisVersion = true,
                MeasSeqStr = "",
                CallSubrs = "Subr1"
            });
            Dictionary<string, HardIpSheet> planDic = NewPlanDic(NewPattern("PAT1"));

            // Act
            _checker.CheckPatInfo(planDic);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfo_MeasSeqAndCallSubrsCountMismatch_AddsMismatchError()
        {
            // Arrange
            SeedHardIpInfo(new HardIpInfo
            {
                Payload = "PAT1",
                PatInfoExist = true,
                UseThisVersion = true,
                MeasSeqStr = "Seq1,Seq2",
                CallSubrs = "Subr1",
                CallSubrsCnt = 1
            });
            Dictionary<string, HardIpSheet> planDic = NewPlanDic(NewPattern("PAT1"));

            // Act
            _checker.CheckPatInfo(planDic);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfo_MeasSeqAndCallSubrsCountMatch_NoError()
        {
            // Arrange
            SeedHardIpInfo(new HardIpInfo
            {
                Payload = "PAT1",
                PatInfoExist = true,
                UseThisVersion = true,
                MeasSeqStr = "Seq1,Seq2",
                CallSubrs = "Subr1,Subr2",
                CallSubrsCnt = 2
            });
            Dictionary<string, HardIpSheet> planDic = NewPlanDic(NewPattern("PAT1"));

            // Act
            _checker.CheckPatInfo(planDic);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfo_IgnorePatInfoMiscInfo_SkipsPatternEntirely()
        {
            // Arrange - miscInfo matches the ignore-pattern-comment marker, so the whole
            // pattern is skipped even though it would otherwise be missing from PatInfo
            Dictionary<string, HardIpSheet> planDic = NewPlanDic(NewPattern("PAT1", miscInfo: "Ignore_Patt_Comment"));

            // Act
            _checker.CheckPatInfo(planDic);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfo_NoPatternToken_SkippedEntirely()
        {
            // Arrange
            Dictionary<string, HardIpSheet> planDic = NewPlanDic(NewPattern("No_patt"));

            // Act
            _checker.CheckPatInfo(planDic);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        #endregion

        #region CheckPatInfoAll

        [TestMethod]
        public void CheckPatInfoAll_DigSrcSignalNameWithoutSendBit_AddsError()
        {
            // Arrange
            var patInfos = new List<HardIpInfo> { new() { Payload = "P1", DigSrcSignalName = "SIG1", SendBit = "" } };

            // Act
            _checker.CheckPatInfoAll(patInfos);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfoAll_DigSrcSignalNameWithSendBit_NoError()
        {
            // Arrange
            var patInfos = new List<HardIpInfo> { new() { Payload = "P1", DigSrcSignalName = "SIG1", SendBit = "1" } };

            // Act
            _checker.CheckPatInfoAll(patInfos);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfoAll_VmVectorSubroutineMismatch_AddsError()
        {
            // Arrange
            var patInfos = new List<HardIpInfo> { new() { Payload = "P1", VmVector = "V1", Subroutine = "wrong" } };

            // Act
            _checker.CheckPatInfoAll(patInfos);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfoAll_VmVectorSubroutineMatch_NoError()
        {
            // Arrange
            var patInfos = new List<HardIpInfo> { new() { Payload = "P1", VmVector = "V1", Subroutine = "V1_srm_meas" } };

            // Act
            _checker.CheckPatInfoAll(patInfos);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfoAll_MlpsubPattern_SkipsVmSubroutineCheck()
        {
            // Arrange - payload matches the MLPSUB pattern regex, bypassing the vm/subr check
            // even though VmVector and Subroutine would otherwise mismatch
            var patInfos = new List<HardIpInfo> { new() { Payload = "X_BI_Y_MLPSUB", VmVector = "V1", Subroutine = "wrong" } };

            // Act
            _checker.CheckPatInfoAll(patInfos);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfoAll_PayloadVmVectorVersionMismatch_AddsError()
        {
            // Arrange
            var patInfos = new List<HardIpInfo> { new() { Payload = "P1", Version = "V2", VmVector = "WRONG" } };

            // Act
            _checker.CheckPatInfoAll(patInfos);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfoAll_PayloadVmVectorVersionMatch_NoError()
        {
            // Arrange
            var patInfos = new List<HardIpInfo> { new() { Payload = "P1", Version = "V2", VmVector = "P1_V2" } };

            // Act
            _checker.CheckPatInfoAll(patInfos);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfoAll_SubroutineWithoutVmVector_AddsError()
        {
            // Arrange
            var patInfos = new List<HardIpInfo> { new() { Payload = "P1", Subroutine = "Subr1", VmVector = "" } };

            // Act
            _checker.CheckPatInfoAll(patInfos);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatInfoAll_SubroutineWithVmVector_NoError()
        {
            // Arrange - Subroutine follows the "{VmVector}_srm_meas" convention so the
            // separate vm/subroutine mismatch check (a different branch) also stays silent
            var patInfos = new List<HardIpInfo> { new() { Payload = "P1", Subroutine = "V1_srm_meas", VmVector = "V1" } };

            // Act
            _checker.CheckPatInfoAll(patInfos);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        #endregion
    }
}
