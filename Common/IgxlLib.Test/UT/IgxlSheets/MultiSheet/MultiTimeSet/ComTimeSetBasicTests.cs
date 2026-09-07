using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets.MultiSheet.MultiTimeSet
{
    [TestClass]
    public class ComTimeSetBasicTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void ComTimeSetBasic_Constructor_DefaultInitialization()
        {
            // Arrange & Act
            var comTimeSetBasic = new ComTimeSetBasic();

            // Assert
            Assert.IsNotNull(comTimeSetBasic);
            Assert.AreEqual(0, comTimeSetBasic.SubCommentVariable.Count);
        }

        [TestMethod]
        public void ComTimeSetBasic_SubCommentVariable_CanAddVariable()
        {
            // Arrange
            var comTimeSetBasic = new ComTimeSetBasic();

            // Act
            comTimeSetBasic.SubCommentVariable.Add("VAR1", 10.5);
            comTimeSetBasic.SubCommentVariable.Add("VAR2", 20.3);

            // Assert
            Assert.AreEqual(2, comTimeSetBasic.SubCommentVariable.Count);
            Assert.AreEqual(10.5, comTimeSetBasic.SubCommentVariable["VAR1"]);
            Assert.AreEqual(20.3, comTimeSetBasic.SubCommentVariable["VAR2"]);
        }

        [TestMethod]
        public void ComTimeSetBasic_SubContextVariable_CanAddVariable()
        {
            // Arrange
            var comTimeSetBasic = new ComTimeSetBasic();

            // Act
            comTimeSetBasic.SubContextVariable.Add("CONTEXT1");
            comTimeSetBasic.SubContextVariable.Add("CONTEXT2");

            // Assert
            Assert.AreEqual(2, comTimeSetBasic.SubContextVariable.Count);
            Assert.AreEqual("CONTEXT1", comTimeSetBasic.SubContextVariable[0]);
            Assert.AreEqual("CONTEXT2", comTimeSetBasic.SubContextVariable[1]);
        }

        [TestMethod]
        public void ComTimeSetBasic_ShiftInReserve_CanAddValue()
        {
            // Arrange
            var comTimeSetBasic = new ComTimeSetBasic();

            // Act
            comTimeSetBasic.ShiftInReserve.Add("SHIFT1", 5.5);
            comTimeSetBasic.ShiftInReserve.Add("SHIFT2", 10.2);

            // Assert
            Assert.AreEqual(2, comTimeSetBasic.ShiftInReserve.Count);
            Assert.AreEqual(5.5, comTimeSetBasic.ShiftInReserve["SHIFT1"]);
            Assert.AreEqual(10.2, comTimeSetBasic.ShiftInReserve["SHIFT2"]);
        }

        [TestMethod]
        public void ComTimeSetBasic_InheritsTSet()
        {
            // Arrange & Act
            var comTimeSetBasic = new ComTimeSetBasic();

            // Assert
            Assert.IsInstanceOfType(comTimeSetBasic, typeof(TSet));
        }

        [TestMethod]
        public void ComTimeSetBasic_CyclePeriod_CanBeSet()
        {
            // Arrange
            var comTimeSetBasic = new ComTimeSetBasic
            {
                // Act
                CyclePeriod = "10ns"
            };

            // Assert
            Assert.AreEqual("10ns", comTimeSetBasic.CyclePeriod);
        }

        [TestMethod]
        public void ComTimeSetBasic_MultipleInstances_Independent()
        {
            // Arrange
            var com1 = new ComTimeSetBasic();
            var com2 = new ComTimeSetBasic();

            // Act
            com1.SubCommentVariable.Add("VAR1", 10.5);
            com2.SubCommentVariable.Add("VAR2", 20.3);

            // Assert
            Assert.AreEqual(1, com1.SubCommentVariable.Count);
            Assert.AreEqual(1, com2.SubCommentVariable.Count);
            Assert.IsFalse(com1.SubCommentVariable.ContainsKey("VAR2"));
        }
    }
}
