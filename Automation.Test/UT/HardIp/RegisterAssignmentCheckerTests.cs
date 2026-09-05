using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class RegisterAssignmentCheckerTests
    {
        private RegisterAssignmentChecker _checker = null!;
        private HardIpPattern _pattern = null!;

        [TestInitialize]
        public void Setup()
        {
            _pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                RowNum = 1,
                Pattern = new PatternClass(""),
                RegisterAssignment = "=AAA"
            };
            var hardIpSheet = new HardIpSheet
            {
                PlanHeaderIdx = new Dictionary<string, int> { { "registerIndex", 1 } }
            };

            _checker = new RegisterAssignmentChecker(hardIpSheet, _pattern, []);
            ErrorReportManager.ClearErrors();
        }

        [TestMethod]
        public void CheckTest()
        {
            // Act
            _checker.Check();

            // Assert
            List<Error> errors = ErrorReportManager.GetErrorList();
            Assert.AreEqual(2, errors.Count);
        }

        [TestMethod]
        public void CheckSweepCode_ShouldNotReportError_WhenStepsAreConsistent()
        {
            // Arrange
            var sweepCodes = new List<SweepCode>
            {
                new() { Type = SweepCode.SweepType.Common, SweepInfo = "0,3,1" },
                new() { Type = SweepCode.SweepType.Common, SweepInfo = "4,7,1" }
            };

            // Act
            _checker.CheckSweepCode(_pattern, sweepCodes, 0);

            // Assert
            List<Error> errors = ErrorReportManager.GetErrorList();
            Assert.AreEqual(0, errors.Count);
        }

        [TestMethod]
        public void CheckSweepCode_ShouldReportError_WhenStepsAreInconsistent()
        {
            // Arrange
            var sweepCodes = new List<SweepCode>
            {
                new() { Type = SweepCode.SweepType.Common, SweepInfo = "0,3,1" },
                new() { Type = SweepCode.SweepType.Common, SweepInfo = "4,8,1" }
            };

            // Act
            _checker.CheckSweepCode(_pattern, sweepCodes, 0);

            // Assert
            List<Error> errors = ErrorReportManager.GetErrorList();
            Assert.IsTrue(errors.Exists(e => e.Message.Contains("Sweep format issue")));
        }

        private static RegisterAssignmentChecker NewChecker(Dictionary<string, string> planDic = null)
        {
            var pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                RowNum = 1,
                Pattern = new PatternClass(""),
                RegisterAssignment = "=AAA"
            };
            var hardIpSheet = new HardIpSheet
            {
                PlanHeaderIdx = new Dictionary<string, int> { { "registerIndex", 1 } }
            };
            return new RegisterAssignmentChecker(hardIpSheet, pattern, planDic ?? []);
        }

        private static (bool Result, bool CheckSendBitCnt, int SendBitCount) InvokeProcessEachAssign(
            RegisterAssignmentChecker checker, string eachAssign, string assign, string nameStr,
            List<string> wrongMessage, List<string> missingKey, List<string> currentBit,
            bool checkSendBitCnt = true, int sendBitCount = 0)
        {
            bool result = checker.ProcessEachAssign(eachAssign, assign, nameStr, null, 1, wrongMessage, missingKey, currentBit, ref checkSendBitCnt, ref sendBitCount);
            return (result, checkSendBitCnt, sendBitCount);
        }

        private static string InvokeHandleRegAssignLabels(RegisterAssignmentChecker checker, string assignStr, Dictionary<string, string> dicLabel)
        {
            checker.HandleRegAssignLabels(ref assignStr, dicLabel);
            return assignStr;
        }

        [TestMethod]
        public void ProcessEachAssign_RegAssignPrefix_ReturnsFalse()
        {
            // Act
            (bool result, bool _, int _) = InvokeProcessEachAssign(_checker, "Reg_assign:foo", "assign", "name", [], [], []);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void ProcessEachAssign_BinaryFormat_TracksBitsAndReturnsTrue()
        {
            // Arrange
            var currentBit = new List<string>();

            // Act
            (bool result, bool checkSendBitCnt, int sendBitCount) = InvokeProcessEachAssign(_checker, "0011", "assign", "name", [], [], currentBit);

            // Assert
            Assert.IsTrue(result);
            Assert.IsTrue(checkSendBitCnt);
            Assert.AreEqual(4, sendBitCount);
            Assert.AreEqual(1, currentBit.Count);
        }

        [TestMethod]
        public void ProcessEachAssign_KeyOnlyFormat_NotFoundInCapLists_AddsMissingKeyAndReturnsFalse()
        {
            // Arrange
            var missingKey = new List<string>();

            // Act
            (bool result, bool checkSendBitCnt, int _) = InvokeProcessEachAssign(_checker, "myKey", "myKey", "name", [], missingKey, []);

            // Assert
            Assert.IsFalse(result);
            Assert.IsFalse(checkSendBitCnt);
            Assert.AreEqual(1, missingKey.Count);
            Assert.AreEqual("myKey", missingKey[0]);
        }

        [TestMethod]
        public void ProcessEachAssign_KeyOnlyFormat_FoundInCapListDir_NotAddedToMissingKey()
        {
            // Arrange
            RegisterAssignmentChecker checker = NewChecker(new Dictionary<string, string> { { "myKey", "x" } });
            var missingKey = new List<string>();

            // Act
            (bool result, bool _, int _) = InvokeProcessEachAssign(checker, "myKey", "myKey", "name", [], missingKey, []);

            // Assert
            Assert.IsFalse(result);
            Assert.AreEqual(0, missingKey.Count);
        }

        [TestMethod]
        public void ProcessEachAssign_KeyIntervalFormat_StartGreaterThanEnd_AddsWarningAndTracksBitWidth()
        {
            // Arrange
            var missingKey = new List<string>();

            // Act
            (bool result, bool _, int sendBitCount) = InvokeProcessEachAssign(_checker, "myKey:3:0", "myKey:3:0", "name", [], missingKey, []);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual(4, sendBitCount);
            Assert.AreEqual(1, missingKey.Count);
            Assert.IsTrue(ErrorReportManager.Contains(HardIpErrorType.W_WrongRegisterAssignment_01.FormatMessage("3", "0")));
        }

        [TestMethod]
        public void ProcessEachAssign_KeyIntervalFormat_StartLessThanEnd_NoWarning()
        {
            // Act
            (bool result, bool _, int sendBitCount) = InvokeProcessEachAssign(_checker, "myKey:0:3", "myKey:0:3", "name", [], [], []);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual(4, sendBitCount);
            Assert.IsFalse(ErrorReportManager.Contains(HardIpErrorType.W_WrongRegisterAssignment_01.FormatMessage("0", "3")));
        }

        [TestMethod]
        public void ProcessEachAssign_KeyIntervalFormat_KeyFoundInCapListDir_NotAddedToMissingKey()
        {
            // Arrange
            RegisterAssignmentChecker checker = NewChecker(new Dictionary<string, string> { { "myKey", "x" } });
            var missingKey = new List<string>();

            // Act
            InvokeProcessEachAssign(checker, "myKey:0:3", "myKey:0:3", "name", [], missingKey, []);

            // Assert
            Assert.AreEqual(0, missingKey.Count);
        }

        [TestMethod]
        public void ProcessEachAssign_UnrecognizedFormat_AddsWrongMessageAndReturnsTrue()
        {
            // Arrange
            var wrongMessage = new List<string>();

            // Act
            (bool result, bool _, int _) = InvokeProcessEachAssign(_checker, "!!!invalid???", "assign", "name", wrongMessage, [], []);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual(1, wrongMessage.Count);
        }

        [TestMethod]
        public void HandleRegAssignLabels_NonSweepLabel_AddsToDictionaryAndReplacesInString()
        {
            // Arrange
            var dicLabel = new Dictionary<string, string>();

            // Act
            string result = InvokeHandleRegAssignLabels(_checker, "myLabel(42);restOfString", dicLabel);

            // Assert
            Assert.AreEqual("42;restOfString", result);
            Assert.AreEqual("42", dicLabel["myLabel"]);
        }

        [TestMethod]
        public void HandleRegAssignLabels_SweepLabel_NotAddedNorReplaced()
        {
            // Arrange
            var dicLabel = new Dictionary<string, string>();

            // Act
            string result = InvokeHandleRegAssignLabels(_checker, "sweep(1,2,3);rest", dicLabel);

            // Assert
            Assert.AreEqual("sweep(1,2,3);rest", result);
            Assert.AreEqual(0, dicLabel.Count);
        }

        [TestMethod]
        public void HandleRegAssignLabels_SelsramLabel_NotAddedNorReplaced()
        {
            // Arrange
            var dicLabel = new Dictionary<string, string>();

            // Act
            string result = InvokeHandleRegAssignLabels(_checker, "selsram(x);rest", dicLabel);

            // Assert
            Assert.AreEqual("selsram(x);rest", result);
            Assert.AreEqual(0, dicLabel.Count);
        }

        [TestMethod]
        public void HandleRegAssignLabels_DuplicateLabel_NotOverwrittenButStillReplaced()
        {
            // Arrange
            var dicLabel = new Dictionary<string, string> { { "myLabel", "old" } };

            // Act
            string result = InvokeHandleRegAssignLabels(_checker, "myLabel(42);rest", dicLabel);

            // Assert
            Assert.AreEqual("42;rest", result);
            Assert.AreEqual("old", dicLabel["myLabel"]);
        }
    }
}
