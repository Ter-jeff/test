using System.Collections.Generic;
using System.Data;
using System.IO;

using Automation.GenerateIgxl.Basic.Business.GenConti.Base;
using Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy;
using Automation.Reader;
using Automation.Static;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Basic;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class CRWTTableTests : FunctionTestBase
    {

        [TestMethod]
        [DataRow("UF")]
        [DataRow("UFP")]
        public void CRWTTest_ByPlatform(string platform)
        {
            string subName = "CRWT_Table";
            string outputPath = Path.Combine(OutputPath, "Common", subName, platform);
            string expectPath = Path.Combine(ExpectPath, "Common", subName, platform);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            List<ContiBase> contiBases = BuildContiBases();

            // Act
            var table = new CrwtTable(contiBases, platform);
            table.ExportToTxt(outputPath);

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail($"Unit Test Fail for platform = {platform} !!!");
            }
        }

        private static List<ContiBase> BuildContiBases()
        {
            var row1 = new DcTestContiRow { PinGroup = "A; B ; C" };
            var conti1 = new ContiPowerShort(row1);

            var row2 = new DcTestContiRow { PinGroup = " C ;  A ;  " };
            var conti2 = new ContiPowerShort(row2);

            return [conti1, conti2];
        }

        [TestMethod]
        public void ReturnDefault_When_JobInfoSheetIsNull()
        {
            // Arrange
            SetJobInfoSheet(null!);

            // Act
            string result = ContiMain.GetTesterTypeOrDefault("JOB1");

            // Assert
            Assert.AreEqual("UF", result);
        }

        private static void SetJobInfoSheet(JobInfoSheet jobInfoSheet)
        {
            TestPlanStatic._jobInfoSheet = jobInfoSheet;
        }

        [TestMethod]
        public void ReturnDefault_When_JobNotFound()
        {
            // Arrange
            _ = new JobInfoSheet
            {
                Rows =
                [
                    new() { JobName = "JOB1", TesterType = "UF" }
                ]
            };

            // Act
            string result = ContiMain.GetTesterTypeOrDefault("NOT_EXIST");

            // Assert
            Assert.AreEqual("UF", result);
        }

        [TestMethod]
        public void ReturnDefault_When_TesterTypeIsNull()
        {
            // Arrange
            TestPlanStatic.JobInfoSheet.Rows.Add(
                new JobInfoRow { JobName = "JOB1", TesterType = null! });

            // Act
            string result = ContiMain.GetTesterTypeOrDefault("JOB1");

            // Assert
            Assert.AreEqual("UF", result);
        }

        [TestMethod]
        public void ReturnDefault_When_TesterTypeIsWhitespace()
        {
            // Arrange
            TestPlanStatic.JobInfoSheet.Rows.Add(
                new JobInfoRow { JobName = "JOB1", TesterType = "   " });

            // Act
            string result = ContiMain.GetTesterTypeOrDefault("JOB1", "PLUS");

            // Assert
            Assert.AreEqual("PLUS", result);
        }

        [TestMethod]
        public void ReturnUpper_When_TesterTypeExists()
        {
            // Arrange

            TestPlanStatic.JobInfoSheet.Rows.Clear();
            TestPlanStatic.JobInfoSheet.Rows.Add(
                new JobInfoRow
                {
                    JobName = "JOB1",
                    TesterType = "plus"
                });

            // Act
            string result = ContiMain.GetTesterTypeOrDefault("job1");

            // Assert
            Assert.AreEqual("PLUS", result);

        }

        [TestMethod]
        public void CRWT_When_PinGroup_IsEmpty_Should_Be_Ignored()
        {
            // Arrange
            var row = new DcTestContiRow { PinGroup = "" };
            var conti = new ContiPowerShort(row);

            // Act
            var table = new CrwtTable([conti], "UF");

            // Assert
            Assert.AreEqual(0, table.Rows.Count);
        }
    }
}
