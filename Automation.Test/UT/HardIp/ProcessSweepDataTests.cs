using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using CommonLib.ErrorReport;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class ProcessSweepDataTests : FunctionTestBase
    {
        [TestMethod]
        public void PatternIsMultiple_ShouldReturnJoinedAssignments()
        {
            // Arrange
            var hardIpPattern = new HardIpPattern
            {
                Pattern = new PatternClass("P1")
                {
                    PatternSetList = [["P1", "P2"]]
                },
                BurstPatterns =
                [
                    new() { Pattern = new PatternClass("P1") { RealPatternName = "P1" }, RegisterAssignment = "Reg_assign_A" },
                    new() { Pattern = new PatternClass("P2") { RealPatternName = "P2" }, RegisterAssignment = "Reg_assign_B" }
                ]
            };

            // Act
            bool isAddSweep = false;
            string result = hardIpPattern.ProcessSweepData(ref isAddSweep);

            // Assert
            Assert.IsFalse(isAddSweep);
            Assert.AreEqual("Reg_assign_A|Reg_assign_B", result);
            Assert.AreEqual("Reg_assign_A|Reg_assign_B", hardIpPattern.RegisterAssignment);
        }

        [TestMethod]
        public void AssignmentWithNestSweep_ShouldProduceSweepSrc()
        {
            var hardIpPattern = new HardIpPattern
            {
                RegisterAssignment = "reg1=nestsweep1(100,200)"
            };
            bool isAddSweep = false;
            _ = hardIpPattern.ProcessSweepData(ref isAddSweep);

            Assert.IsTrue(isAddSweep);
            CollectionAssert.Contains(hardIpPattern.DspFunction, "sweepsrc:sweep1=100@200");
        }

        [TestMethod]
        public void AssignmentWithSquareBracketSweep_ShouldProduceSweepSrc()
        {
            var hardIpPattern = new HardIpPattern
            {
                RegisterAssignment = "reg2=sweep[300]"
            };
            bool isAddSweep = false;
            _ = hardIpPattern.ProcessSweepData(ref isAddSweep);

            Assert.IsTrue(isAddSweep);
            CollectionAssert.Contains(hardIpPattern.DspFunction, "sweepsrc:sweep1=300");
        }

        [TestMethod]
        public void AssignmentWithSweepSingleLoop_ShouldCheckTableExists()
        {
            var hardIpPattern = new HardIpPattern
            {
                RegisterAssignment = "reg3=Sweep1_Singleloop::MyTable"
            };

            bool isAddSweep = false;
            _ = hardIpPattern.ProcessSweepData(ref isAddSweep);

            Assert.IsTrue(isAddSweep);
            string expected = "Missing sweep or single-loop table definition sheet MyTable in TestPlan";
            Assert.IsTrue(ErrorReportManager.GetErrorList().Exists(e => e.Message.Contains(expected)));
            CollectionAssert.Contains(hardIpPattern.DspFunction, "sweepsrc_singleloop:MyTable");
        }

        [TestMethod]
        public void AssignmentWithNormalRegister_ShouldStaySame()
        {
            var hardIpPattern = new HardIpPattern
            {
                RegisterAssignment = "reg4=abcd"
            };
            bool isAddSweep = false;
            string result = hardIpPattern.ProcessSweepData(ref isAddSweep);

            Assert.IsFalse(isAddSweep);
            Assert.AreEqual("reg4=abcd", result);
        }

        [TestMethod]
        public void NestSweepVoltageMerging_ShouldReorderCorrectly()
        {
            var hardIpPattern = new HardIpPattern
            {
                RegisterAssignment = "reg5=nestsweep1(111)",
                SweepVoltage =
                {
                    ["NestSweep"] =
                    [
                        new("1,5,1") { Order = "2", Start = "0.1", Stop = "0.2", Step = "0.01" }
                    ]
                }
            };
            bool isAddSweep = false;
            string result = hardIpPattern.ProcessSweepData(ref isAddSweep);

            Assert.IsTrue(isAddSweep);
            Assert.AreEqual("reg5=sweep1", result);
        }

        [TestMethod]
        public void NoSweepButVoltageExists_ShouldReturnVoltageOnly()
        {
            var hardIpPattern = new HardIpPattern
            {
                RegisterAssignment = "reg6=abcd",
                SweepVoltage =
                {
                    ["NestSweep"] =
                    [
                        new("1,5,1") { VoltageList = "1.0,1.1,1.2" }
                    ]
                }
            };
            bool isAddSweep = false;
            _ = hardIpPattern.ProcessSweepData(ref isAddSweep);

            Assert.IsFalse(isAddSweep);
            Assert.AreEqual("sweepvoltage=1.0,1.1,1.2", hardIpPattern.DspFunction.First());
        }

        [TestMethod]
        public void BurstMode_ShouldUpdateRegisterAssignment()
        {
            var hardIpPattern = new HardIpPattern
            {
                RegisterAssignment = "reg1=nestsweep(10)"
            };
            bool isAddSweep = false;
            string result = hardIpPattern.ProcessSweepData(ref isAddSweep, isBurst: true);

            Assert.IsTrue(isAddSweep);
            Assert.AreEqual(hardIpPattern.RegisterAssignment, result);
            StringAssert.Contains(result, "sweep1");
        }

        [TestMethod]
        public void DuplicateSweepSrc_ShouldBeDistinct()
        {
            var hardIpPattern = new HardIpPattern
            {
                RegisterAssignment = "reg1=nestsweep(10);reg2=nestsweep(10)"
            };
            bool isAddSweep = false;
            _ = hardIpPattern.ProcessSweepData(ref isAddSweep);

            Assert.IsTrue(isAddSweep);
            Assert.AreEqual("sweepsrc:sweep1=10;sweep2=10", hardIpPattern.DspFunction.First());
        }

        [TestMethod]
        public void NoMatches_ShouldReturnOriginalRegisterAssignment()
        {
            var hardIpPattern = new HardIpPattern
            {
                RegisterAssignment = "regX=something"
            };
            bool isAddSweep = false;
            string result = hardIpPattern.ProcessSweepData(ref isAddSweep);

            Assert.IsFalse(isAddSweep);
            Assert.AreEqual("regX=something", result);
        }
    }
}
