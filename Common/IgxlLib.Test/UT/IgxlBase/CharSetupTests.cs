using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class CharSetupTests
    {
        [TestMethod]
        public void CharSetup_DefaultConstructor_InitializesProperties()
        {
            // Arrange & Act
            var charSetup = new CharSetup();

            // Assert
            Assert.AreEqual("", charSetup.SetupName);
            Assert.AreEqual("", charSetup.TestMethod);
            Assert.AreEqual(0, charSetup.CharSteps.Count);
        }

        [TestMethod]
        public void CharSetup_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var charSetup = new CharSetup
            {
                // Act
                SetupName = "CharSetup1",
                TestMethod = "VoltageCharacterization"
            };

            // Assert
            Assert.AreEqual("CharSetup1", charSetup.SetupName);
            Assert.AreEqual("VoltageCharacterization", charSetup.TestMethod);
        }

        [TestMethod]
        public void CharSetup_AddStep_AddsCharStepToList()
        {
            // Arrange
            var charSetup = new CharSetup();
            var charStep = new CharStep("", "") { StepName = "Step1" };

            // Act
            charSetup.AddStep(charStep);

            // Assert
            Assert.AreEqual(1, charSetup.CharSteps.Count);
            Assert.AreEqual(charStep, charSetup.CharSteps[0]);
        }

        [TestMethod]
        public void CharSetup_AddMultipleSteps_AddsAllSteps()
        {
            // Arrange
            var charSetup = new CharSetup();
            var step1 = new CharStep("", "") { StepName = "Step1" };
            var step2 = new CharStep("", "") { StepName = "Step2" };
            var step3 = new CharStep("", "") { StepName = "Step3" };

            // Act
            charSetup.AddStep(step1);
            charSetup.AddStep(step2);
            charSetup.AddStep(step3);

            // Assert
            Assert.AreEqual(3, charSetup.CharSteps.Count);
            Assert.AreEqual("Step1", charSetup.CharSteps[0].StepName);
            Assert.AreEqual("Step2", charSetup.CharSteps[1].StepName);
            Assert.AreEqual("Step3", charSetup.CharSteps[2].StepName);
        }

        [TestMethod]
        public void CharSetup_CopyConstructor_DeepCopiesCharSteps()
        {
            // Arrange
            var originalSetup = new CharSetup
            {
                SetupName = "OriginalSetup",
                TestMethod = "OriginalMethod"
            };
            originalSetup.AddStep(new CharStep("", "") { StepName = "Step1", VoltageType = "DC" });
            originalSetup.AddStep(new CharStep("", "") { StepName = "Step2", VoltageType = "AC" });

            // Act
            var copiedSetup = new CharSetup(originalSetup);

            // Assert
            Assert.AreEqual("OriginalSetup", copiedSetup.SetupName);
            Assert.AreEqual("OriginalMethod", copiedSetup.TestMethod);
            Assert.AreEqual(2, copiedSetup.CharSteps.Count);

            // Verify deep copy - modifying copied setup doesn't affect original
            copiedSetup.CharSteps[0].StepName = "ModifiedStep";
            Assert.AreEqual("Step1", originalSetup.CharSteps[0].StepName);
            Assert.AreEqual("ModifiedStep", copiedSetup.CharSteps[0].StepName);
        }

        [TestMethod]
        public void CharSetup_CopyConstructor_WithNull_CreatesEmptyInstance()
        {
            // Arrange & Act
            var charSetup = new CharSetup(null);

            // Assert
            Assert.AreEqual("", charSetup.SetupName);
            Assert.AreEqual("", charSetup.TestMethod);
        }

        [TestMethod]
        public void CharSetup_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var charSetup = new CharSetup();

            // Assert
            Assert.IsInstanceOfType(charSetup, typeof(IgxlRow));
        }
    }
}
