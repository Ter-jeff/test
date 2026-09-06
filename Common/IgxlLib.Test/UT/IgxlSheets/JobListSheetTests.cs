using System.Collections.Generic;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class JobListSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void JobListSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "JobList";

            // Act
            var jobListSheet = new JobListSheet(sheetName);

            // Assert
            Assert.IsNotNull(jobListSheet);
            Assert.AreEqual(sheetName, jobListSheet.Name);
            Assert.AreEqual("DTJobListSheet", jobListSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.JobList, jobListSheet.IgxlSheetName);
        }

        [TestMethod]
        public void JobListSheet_AddRow()
        {
            // Arrange
            var jobListSheet = new JobListSheet("JobList");
            var jobRow = new JobRow { JobName = "Job1", FlowTable = "Flow1" };

            // Act
            jobListSheet.AddRow(jobRow);

            // Assert
            Assert.AreEqual(1, jobListSheet.Rows.Count);
        }

        [TestMethod]
        public void JobListSheet_AddRows()
        {
            // Arrange
            var jobListSheet = new JobListSheet("JobList");
            var rows = new List<JobRow>
            {
                new() { JobName = "Job1", FlowTable = "Flow1" },
                new() { JobName = "Job2", FlowTable = "Flow2" },
                new() { JobName = "Job3", FlowTable = "Flow3" }
            };

            // Act
            jobListSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, jobListSheet.Rows.Count);
        }

        [TestMethod]
        public void JobListSheet_RemoveRow()
        {
            // Arrange
            var jobListSheet = new JobListSheet("JobList");
            var row1 = new JobRow { JobName = "Job1", FlowTable = "Flow1" };
            var row2 = new JobRow { JobName = "Job2", FlowTable = "Flow2" };
            jobListSheet.AddRow(row1);
            jobListSheet.AddRow(row2);

            // Act
            jobListSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, jobListSheet.Rows.Count);
            Assert.AreEqual("Job2", jobListSheet.Rows[0].JobName);
        }

        [TestMethod]
        public void JobListSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var jobListSheet = new JobListSheet("JobList");

            // Assert
            Assert.AreEqual("DTJobListSheet", jobListSheet.SheetType);
        }

        [TestMethod]
        public void JobListSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var jobListSheet = new JobListSheet("JobList");

            // Assert
            Assert.IsNotNull(jobListSheet.GetErrors());
            Assert.AreEqual(0, jobListSheet.GetErrors().Count);
        }

        [TestMethod]
        public void JobListSheet_Name_CanBeSet()
        {
            // Arrange
            var jobListSheet = new JobListSheet("JobList")
            {
                // Act
                Name = "NewJobListName"
            };

            // Assert
            Assert.AreEqual("NewJobListName", jobListSheet.Name);
        }
    }
}
