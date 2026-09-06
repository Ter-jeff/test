using System.Collections.Generic;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class FlowRowTests
    {
        [TestMethod]
        public void FlowRow_DefaultConstructor_InitializesWithEmptyStrings()
        {
            // Arrange & Act
            var flowRow = new FlowRow();

            // Assert
            Assert.AreEqual(string.Empty, flowRow.Label);
            Assert.AreEqual(string.Empty, flowRow.Enable);
            Assert.AreEqual(string.Empty, flowRow.Parameter);
            Assert.AreEqual(string.Empty, flowRow.Opcode);
            Assert.IsFalse(flowRow.IsSsn);
        }

        [TestMethod]
        public void FlowRow_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var flowRow = new FlowRow
            {
                // Act
                Label = "Test1",
                Enable = "1",
                Job = "JobA",
                Parameter = "Param1",
                Opcode = "OpCode1",
                TName = "TestName",
                LoLim = "0",
                HiLim = "10",
                BinPass = "100",
                BinFail = "200",
                Comment = "Flow test"
            };

            // Assert
            Assert.AreEqual("Test1", flowRow.Label);
            Assert.AreEqual("1", flowRow.Enable);
            Assert.AreEqual("JobA", flowRow.Job);
            Assert.AreEqual("Param1", flowRow.Parameter);
            Assert.AreEqual("OpCode1", flowRow.Opcode);
            Assert.AreEqual("TestName", flowRow.TName);
            Assert.AreEqual("0", flowRow.LoLim);
            Assert.AreEqual("10", flowRow.HiLim);
            Assert.AreEqual("100", flowRow.BinPass);
            Assert.AreEqual("200", flowRow.BinFail);
            Assert.AreEqual("Flow test", flowRow.Comment);
        }

        [TestMethod]
        public void FlowRow_SetDebugAndProfileProperties_UpdatesCorrectly()
        {
            // Arrange
            var flowRow = new FlowRow
            {
                // Act
                DebugAssume = "Assume1",
                DebugSites = "Site1",
                CtProfileDataElapsedTime = "100ms",
                CtProfileDataBackgroundType = "Background",
                CtProfileDataSerialize = "Yes",
                IsSsn = true
            };

            // Assert
            Assert.AreEqual("Assume1", flowRow.DebugAssume);
            Assert.AreEqual("Site1", flowRow.DebugSites);
            Assert.AreEqual("100ms", flowRow.CtProfileDataElapsedTime);
            Assert.AreEqual("Background", flowRow.CtProfileDataBackgroundType);
            Assert.AreEqual("Yes", flowRow.CtProfileDataSerialize);
            Assert.IsTrue(flowRow.IsSsn);
        }

        [TestMethod]
        public void FlowRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var flowRow = new FlowRow();

            // Assert
            Assert.IsInstanceOfType(flowRow, typeof(IgxlRow));
        }

        [TestMethod]
        public void FlowRow_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var row1 = new FlowRow { Parameter = "Param1", Opcode = "Op1" };
            var row2 = new FlowRow { Parameter = "Param2", Opcode = "Op2" };

            // Assert
            Assert.AreEqual("Param1", row1.Parameter);
            Assert.AreEqual("Param2", row2.Parameter);
            Assert.AreNotEqual(row1.Opcode, row2.Opcode);
        }

        [TestMethod]
        public void FlowRow_Copy_ShouldCloneAllFields()
        {
            var original = new FlowRow
            {
                Label = "L1",
                Job = "J1",
                PatIndex = 5,
                IsSsn = true
            };

            FlowRow copy = original.Copy();

            Assert.AreEqual(original.Label, copy.Label);
            Assert.AreEqual(original.Job, copy.Job);
            Assert.AreEqual(original.PatIndex, copy.PatIndex);
            Assert.AreEqual(original.IsSsn, copy.IsSsn);

            Assert.AreNotSame(original, copy);
        }

        [TestMethod]
        public void FlowRow_CopyConstructor_NullInput_ShouldNotThrow()
        {
            FlowRow row = new FlowRow(null);

            Assert.IsNotNull(row);
            Assert.AreEqual(string.Empty, row.Label);
        }

        [TestMethod]
        public void GetJobs_ShouldSplitCorrectly()
        {
            FlowRow row = new FlowRow { Job = "A,B,C" };

            List<string> result = row.GetJobs();

            CollectionAssert.AreEqual(new List<string> { "A", "B", "C" }, result);
        }

        [TestMethod]
        public void GetJobs_EmptyString()
        {
            FlowRow row = new FlowRow { Job = "" };

            List<string> result = row.GetJobs();

            CollectionAssert.AreEqual(new List<string>(), result);
        }

        [TestMethod]
        public void GenEndIf_ShouldGenerateCorrectOpcode()
        {
            FlowRow result = FlowRow.GenEndIf("J1");

            Assert.AreEqual("EndIf", result.Opcode);
        }

        [TestMethod]
        public void GenIfCondition_ShouldFormatParameter()
        {
            FlowRow result = FlowRow.GenIfCondition("A==1", "J1");

            Assert.AreEqual("If", result.Opcode);
            Assert.AreEqual("(A==1) || F_Debug_all", result.Parameter);
        }

        [TestMethod]
        public void AddColumnA_FirstInsert()
        {
            FlowRow row = new FlowRow();

            row.AddColumnA("A");

            Assert.AreEqual("A", row.ColumnA);
        }

        [TestMethod]
        public void AddColumnA_Append()
        {
            FlowRow row = new FlowRow();
            row.AddColumnA("A");

            row.AddColumnA("B");

            Assert.AreEqual("A,B", row.ColumnA);
        }

        [TestMethod]
        public void DefaultValues_ShouldBeCorrect()
        {
            FlowRow row = new FlowRow();

            Assert.AreEqual(-1, row.PatIndex);
            Assert.AreEqual(0, row.BinTableFlagCount);
        }

        [TestMethod]
        public void Copy_ShouldBeIndependent()
        {
            FlowRow original = new FlowRow { Label = "A" };
            FlowRow copy = original.Copy();

            copy.Label = "B";

            Assert.AreEqual("A", original.Label);
        }
    }
}
