using System.Collections.Generic;

using Automation.Reader;

using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class NwireReaderTests
    {
        [TestMethod]
        public void ResolveVoltage_ValidDouble_ShouldSetOutClkVoltage()
        {
            // Arrange
            var pin = new ProtocolAwarePin();
            string targetValue = "3.3";

            // Act
            NwireReader.ResolveVoltage(targetValue, v => pin.OutClkVoltage = v);

            // Assert
            Assert.AreEqual(3.3, pin.OutClkVoltage);
        }

        [TestMethod]
        public void ResolveVoltage_ValidDouble_ShouldSetRefClkVoltage()
        {
            // Arrange
            var pin = new ProtocolAwarePin();
            string targetValue = "1.8";

            // Act
            NwireReader.ResolveVoltage(targetValue, v => pin.RefClkVoltage = v);

            // Assert
            Assert.AreEqual(1.8, pin.RefClkVoltage);
        }

        [TestMethod]
        public void ResolveVoltage_ValidDouble_ShouldSetOutAndRefVoltage()
        {
            // Arrange
            var pin = new ProtocolAwarePin();
            string targetValue = "2.5";

            // Act
            NwireReader.ResolveVoltage(targetValue, v =>
            {
                pin.OutClkVoltage = v;
                pin.RefClkVoltage = v;
            });

            // Assert
            Assert.AreEqual(2.5, pin.OutClkVoltage);
            Assert.AreEqual(2.5, pin.RefClkVoltage);
        }

        [TestMethod]
        public void ResolveVoltage_InvalidDouble_ShouldNotInvokeSetter()
        {
            // Arrange
            bool setterInvoked = false;
            string targetValue = "ABC";

            // Act
            NwireReader.ResolveVoltage(targetValue, v => setterInvoked = true);

            // Assert
            Assert.IsFalse(setterInvoked);
        }

        [TestMethod]
        public void ResolveVoltage_InvalidDouble_ShouldNotChangeExistingVoltage()
        {
            // Arrange
            var pin = new ProtocolAwarePin
            {
                OutClkVoltage = 1.2,
                RefClkVoltage = 0.9
            };
            string targetValue = "XYZ";

            // Act
            NwireReader.ResolveVoltage(targetValue, v =>
            {
                pin.OutClkVoltage = v;
                pin.RefClkVoltage = v;
            });

            // Assert
            Assert.AreEqual(1.2, pin.OutClkVoltage);
            Assert.AreEqual(0.9, pin.RefClkVoltage);
        }

        [TestMethod]
        public void TryAddPattern_LengthLessOrEqual20_ShouldReturnFalse()
        {
            var dic = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            int cnt = 1;

            bool result = NwireReader.TryAddPattern(dic, "SHORT_NAME", ref cnt);

            Assert.IsFalse(result);
            Assert.AreEqual(0, dic.Count);
            Assert.AreEqual(1, cnt);
        }

        [TestMethod]
        public void TryAddPattern_ValidNewLongPattern_ShouldAddAndIncreaseCounter()
        {
            var dic = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            int cnt = 1;

            bool result = NwireReader.TryAddPattern(dic, "THIS_IS_A_VERY_LONG_PATTERN_NAME", ref cnt);

            Assert.IsTrue(result);
            Assert.AreEqual(1, dic.Count);
            Assert.AreEqual("Pattern_1", dic["THIS_IS_A_VERY_LONG_PATTERN_NAME"]);
            Assert.AreEqual(2, cnt);
        }

        [TestMethod]
        public void TryAddPattern_DuplicatePattern_ShouldNotAddAgain()
        {
            var dic = new Dictionary<string, string>(StringExtensions.IgnoreCase)
        {
            { "THIS_IS_A_VERY_LONG_PATTERN_NAME", "Pattern_1" }
        };
            int cnt = 2;

            bool result = NwireReader.TryAddPattern(dic, "THIS_IS_A_VERY_LONG_PATTERN_NAME", ref cnt);

            Assert.IsFalse(result);
            Assert.AreEqual(1, dic.Count);

        }

        [TestMethod]
        public void HasMissingRequiredSetting_AnyRowIsMinusOne_ShouldReturnTrue()
        {
            Assert.IsTrue(NwireReader.HasMissingRequiredSetting(-1, 0));

            Assert.IsTrue(NwireReader.HasMissingRequiredSetting(0, -1));

            Assert.IsFalse(NwireReader.HasMissingRequiredSetting(0, 0));

        }

        [TestMethod]
        public void HasMissingRequiredSetting_AllRowsValid_ShouldReturnFalse()
        {
            Assert.IsFalse(NwireReader.HasMissingRequiredSetting(1, 2));
        }
    }
}
