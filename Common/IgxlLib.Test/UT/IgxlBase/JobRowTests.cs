using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class JobRowTests
    {
        [TestMethod]
        public void JobRow_DefaultConstructor_InitializesWithEmptyStrings()
        {
            // Arrange & Act
            var jobRow = new JobRow();

            // Assert
            Assert.AreEqual("", jobRow.JobName);
            Assert.AreEqual("", jobRow.PinMap);
            Assert.AreEqual("", jobRow.TestInstances);
            Assert.AreEqual("", jobRow.FlowTable);
            Assert.AreEqual("", jobRow.Comment);
        }

        [TestMethod]
        public void JobRow_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var jobRow = new JobRow
            {
                // Act
                JobName = "Job1",
                PinMap = "PinMap1",
                TestInstances = "TestInstances1",
                FlowTable = "FlowTable1",
                AcSpecs = "ACSpecs1",
                DcSpecs = "DCSpecs1",
                PatternSets = "PatternSets1",
                BinTable = "BinTable1",
                Comment = "Job comment"
            };

            // Assert
            Assert.AreEqual("Job1", jobRow.JobName);
            Assert.AreEqual("PinMap1", jobRow.PinMap);
            Assert.AreEqual("TestInstances1", jobRow.TestInstances);
            Assert.AreEqual("FlowTable1", jobRow.FlowTable);
            Assert.AreEqual("ACSpecs1", jobRow.AcSpecs);
            Assert.AreEqual("DCSpecs1", jobRow.DcSpecs);
            Assert.AreEqual("PatternSets1", jobRow.PatternSets);
            Assert.AreEqual("BinTable1", jobRow.BinTable);
        }

        [TestMethod]
        public void JobRow_SetAllProperties_UpdatesAll()
        {
            // Arrange
            var jobRow = new JobRow
            {
                // Act
                JobName = "CompleteJob",
                PinMap = "PinMap",
                TestInstances = "Instances",
                FlowTable = "Flow",
                AcSpecs = "AC",
                DcSpecs = "DC",
                PatternSets = "Patterns",
                PatternGroups = "Groups",
                BinTable = "Bins",
                Characterization = "Char",
                TestProcedures = "Procs",
                MixedSignalTiming = "MST",
                WaveDefinitions = "Waves",
                PSets = "PSets",
                Signals = "Signals",
                PortMap = "PortMap",
                FractionalBus = "FBus",
                ConcurrentSequence = "CSeq",
                Comment = "Complete job"
            };

            // Assert
            Assert.AreEqual("CompleteJob", jobRow.JobName);
            Assert.AreEqual("AC", jobRow.AcSpecs);
            Assert.AreEqual("DC", jobRow.DcSpecs);
            Assert.AreEqual("Patterns", jobRow.PatternSets);
            Assert.AreEqual("Complete job", jobRow.Comment);
        }

        [TestMethod]
        public void JobRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var jobRow = new JobRow();

            // Assert
            Assert.IsInstanceOfType(jobRow, typeof(IgxlRow));
        }

        [TestMethod]
        public void JobRow_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var job1 = new JobRow { JobName = "Job1", PinMap = "PinMap1" };
            var job2 = new JobRow { JobName = "Job2", PinMap = "PinMap2" };

            // Assert
            Assert.AreEqual("Job1", job1.JobName);
            Assert.AreEqual("Job2", job2.JobName);
            Assert.AreNotEqual(job1.JobName, job2.JobName);
        }
    }
}
