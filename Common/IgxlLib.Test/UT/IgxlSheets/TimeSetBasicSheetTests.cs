using System.Collections.Generic;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class TimeSetBasicSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void TimeSetBasicSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "TimeSet";

            // Act
            var timeSetSheet = new TimeSetBasicSheet(sheetName);

            // Assert
            Assert.IsNotNull(timeSetSheet);
            Assert.AreEqual(sheetName, timeSetSheet.Name);
            Assert.AreEqual("DTTimesetBasicSheet", timeSetSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.TimeSetsBasic, timeSetSheet.IgxlSheetName);
            Assert.AreEqual("", timeSetSheet.TimingMode);
        }

        [TestMethod]
        public void TimeSetBasicSheet_Constructor_WithSheetNameAndTimingMode()
        {
            // Arrange
            string sheetName = "TimeSet";
            string timingMode = "ASYNC";

            // Act
            var timeSetSheet = new TimeSetBasicSheet(sheetName, timingMode);

            // Assert
            Assert.IsNotNull(timeSetSheet);
            Assert.AreEqual(sheetName, timeSetSheet.Name);
            Assert.AreEqual(timingMode, timeSetSheet.TimingMode);
        }

        [TestMethod]
        public void TimeSetBasicSheet_AddRow()
        {
            // Arrange
            var timeSetSheet = new TimeSetBasicSheet("TimeSet");
            var tSet = new TSet { CyclePeriod = "10ns" };

            // Act
            timeSetSheet.AddRow(tSet);

            // Assert
            Assert.AreEqual(1, timeSetSheet.Rows.Count);
        }

        [TestMethod]
        public void TimeSetBasicSheet_AddRows()
        {
            // Arrange
            var timeSetSheet = new TimeSetBasicSheet("TimeSet");
            var rows = new List<TSet>
            {
                new() { CyclePeriod = "10ns" },
                new() { CyclePeriod = "20ns" }
            };

            // Act
            timeSetSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(2, timeSetSheet.Rows.Count);
        }

        [TestMethod]
        public void TimeSetBasicSheet_CommentVariable_Dictionary()
        {
            // Arrange
            var timeSetSheet = new TimeSetBasicSheet("TimeSet");

            // Act
            timeSetSheet.CommentVariable.Add("VAR1", 10.5);
            timeSetSheet.CommentVariable.Add("VAR2", 20.3);

            // Assert
            Assert.AreEqual(2, timeSetSheet.CommentVariable.Count);
            Assert.AreEqual(10.5, timeSetSheet.CommentVariable["VAR1"]);
            Assert.AreEqual(20.3, timeSetSheet.CommentVariable["VAR2"]);
        }

        [TestMethod]
        public void TimeSetBasicSheet_TimingMode_CanBeSet()
        {
            // Arrange
            var timeSetSheet = new TimeSetBasicSheet("TimeSet")
            {
                // Act
                TimingMode = "SYNC"
            };

            // Assert
            Assert.AreEqual("SYNC", timeSetSheet.TimingMode);
        }

        [TestMethod]
        public void TimeSetBasicSheet_MasterTimeSet_CanBeSet()
        {
            // Arrange
            var timeSetSheet = new TimeSetBasicSheet("TimeSet")
            {
                // Act
                MasterTimeSet = "MasterTimeSetName"
            };

            // Assert
            Assert.AreEqual("MasterTimeSetName", timeSetSheet.MasterTimeSet);
        }

        [TestMethod]
        public void TimeSetBasicSheet_TimeDomain_CanBeSet()
        {
            // Arrange
            var timeSetSheet = new TimeSetBasicSheet("TimeSet")
            {
                // Act
                TimeDomain = "DOMAIN1"
            };

            // Assert
            Assert.AreEqual("DOMAIN1", timeSetSheet.TimeDomain);
        }

        [TestMethod]
        public void TimeSetBasicSheet_StrobeRefSetup_CanBeSet()
        {
            // Arrange
            var timeSetSheet = new TimeSetBasicSheet("TimeSet")
            {
                // Act
                StrobeRefSetup = "SETUP1"
            };

            // Assert
            Assert.AreEqual("SETUP1", timeSetSheet.StrobeRefSetup);
        }

        [TestMethod]
        public void TimeSetBasicSheet_RemoveRow()
        {
            // Arrange
            var timeSetSheet = new TimeSetBasicSheet("TimeSet");
            var tSet1 = new TSet { CyclePeriod = "10ns" };
            var tSet2 = new TSet { CyclePeriod = "20ns" };
            timeSetSheet.AddRow(tSet1);
            timeSetSheet.AddRow(tSet2);

            // Act
            timeSetSheet.RemoveRow(tSet1);

            // Assert
            Assert.AreEqual(1, timeSetSheet.Rows.Count);
        }

        [TestMethod]
        public void TimeSetBasicSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var timeSetSheet = new TimeSetBasicSheet("TimeSet");

            // Assert
            Assert.IsNotNull(timeSetSheet.GetErrors());
            Assert.AreEqual(0, timeSetSheet.GetErrors().Count);
        }

        [TestMethod]
        public void TimeSetBasicSheet_Copy()
        {
            // Arrange
            var originalSheet = new TimeSetBasicSheet("TimeSet", "ASYNC")
            {
                TimingMode = "SYNC",
                MasterTimeSet = "Master",
                TimeDomain = "Domain1",
                StrobeRefSetup = "Setup1"
            };
            originalSheet.AddRow(new TSet { CyclePeriod = "10ns" });

            // Act
            var copiedSheet = new TimeSetBasicSheet(originalSheet);

            // Assert
            Assert.AreNotSame(originalSheet, copiedSheet);
            Assert.AreEqual(originalSheet.Name, copiedSheet.Name);
            Assert.AreEqual(originalSheet.TimingMode, copiedSheet.TimingMode);
            Assert.AreEqual(originalSheet.MasterTimeSet, copiedSheet.MasterTimeSet);
            Assert.AreEqual(originalSheet.TimeDomain, copiedSheet.TimeDomain);
            Assert.AreEqual(originalSheet.StrobeRefSetup, copiedSheet.StrobeRefSetup);
            Assert.AreEqual(originalSheet.Rows.Count, copiedSheet.Rows.Count);
        }

        [TestMethod]
        public void TimeSetBasicSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var timeSetSheet = new TimeSetBasicSheet("TimeSet");

            // Assert
            Assert.AreEqual("DTTimesetBasicSheet", timeSetSheet.SheetType);
        }
    }
}
