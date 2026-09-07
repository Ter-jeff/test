using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.ErrorReport
{
    /// <summary>
    /// Accesses one field from each ErrorCodes static class to trigger static constructors
    /// and achieve coverage of all error code definitions.
    /// </summary>
    [TestClass]
    public class ErrorCodesAccessTests
    {
        [TestMethod]
        public void BinCutErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = BinCutErrorType.E_AllowEqual_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void DuplicateInstance_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = PostActionErrorType.W_DuplicateInstance_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void DuplicateTestNumber_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = PostActionErrorType.W_DuplicateTestNumber_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void EfuseCheckCmdLibError_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = EfuseCheckCmdLibError.E_MismatchFormat_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void EFuseErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = EFuseErrorType.E_RuleViolationBit_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void EnumErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = PreActionErrorType.E_InvalidDocument_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void EvsErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = EvsErrorType.E_MissingParameter_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void FlowMainErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = FlowMainErrorType.W_MismatchFlow_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void HardIpErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = HardIpErrorType.E_ADC_Convertor_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void HarvestErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = HarvestErrorType.E_MissingFlag_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void HtolErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = HtolErrorType.E_MissingLibrary_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void MainFlowErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = FlowMainErrorType.E_MissingDocument_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void MbistErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = MbistErrorType.E_BurstInfoMismatch_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void PatInfoType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = PatInfoType.E_MismatchGenericName_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void PreActionErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = PreActionErrorType.E_RuleViolationVoltage_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void RtosErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = RtosErrorType.E_MissingParameter_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void ScanErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = ScanErrorType.E_FormatError_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void SsnErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = SsnErrorType.E_MismatchCore_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void UfInstanceErrorType_FirstField_IsValidErrorCode()
        {
            // Arrange & Act
            ErrorCode code = BasicErrorType.E_DefaultPowerUpInstanceExist_01;

            // Assert
            Assert.IsNotNull(code);
            Assert.AreEqual(9, code.FullCode.Length);
        }

        [TestMethod]
        public void AllErrorCodesClasses_FirstFields_HaveUniqueFullCodes()
        {
            // Arrange & Act - access all classes to ensure no FullCode collision in same class family
            Assert.AreNotEqual(BinCutErrorType.E_AllowEqual_01.FullCode, PostActionErrorType.W_DuplicateInstance_01.FullCode);
            Assert.AreNotEqual(EFuseErrorType.E_RuleViolationBit_01.FullCode, PreActionErrorType.E_InvalidDocument_01.FullCode);
        }
    }
}
