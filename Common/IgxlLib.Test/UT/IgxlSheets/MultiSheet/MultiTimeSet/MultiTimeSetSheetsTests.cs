using System.Collections.Generic;

using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets.MultiSheet.MultiTimeSet
{
    [TestClass]
    public class MultiTimeSetSheetsTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void MultiTimeSetSheets_Constructor()
        {
            // Arrange & Act
            var multiTimeSetSheets = new MultiTimeSetSheets();

            // Assert
            Assert.IsNotNull(multiTimeSetSheets);
            Assert.AreEqual(0, multiTimeSetSheets.TimeSetBasicSheetsList.Count);
        }

        [TestMethod]
        public void MultiTimeSetSheets_TimeSetBasicSheetsList_Initialized()
        {
            // Arrange & Act
            var multiTimeSetSheets = new MultiTimeSetSheets();

            // Assert
            Assert.IsInstanceOfType(multiTimeSetSheets.TimeSetBasicSheetsList, typeof(List<ComTimeSetBasicSheet>));
        }

        [TestMethod]
        public void MultiTimeSetSheets_AddTimeSetSheet()
        {
            // Arrange
            var multiTimeSetSheets = new MultiTimeSetSheets();
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet("TimeSet1");

            // Act
            multiTimeSetSheets.AddTimeSetSheet(comTimeSetBasicSheet);

            // Assert
            Assert.AreEqual(1, multiTimeSetSheets.TimeSetBasicSheetsList.Count);
            Assert.AreEqual("TimeSet1", multiTimeSetSheets.TimeSetBasicSheetsList[0].Name);
        }

        [TestMethod]
        public void MultiTimeSetSheets_AddMultipleTimeSetSheets()
        {
            // Arrange
            var multiTimeSetSheets = new MultiTimeSetSheets();
            var timeSet1 = new ComTimeSetBasicSheet("TimeSet1");
            var timeSet2 = new ComTimeSetBasicSheet("TimeSet2");
            var timeSet3 = new ComTimeSetBasicSheet("TimeSet3");

            // Act
            multiTimeSetSheets.AddTimeSetSheet(timeSet1);
            multiTimeSetSheets.AddTimeSetSheet(timeSet2);
            multiTimeSetSheets.AddTimeSetSheet(timeSet3);

            // Assert
            Assert.AreEqual(3, multiTimeSetSheets.TimeSetBasicSheetsList.Count);
            Assert.AreEqual("TimeSet1", multiTimeSetSheets.TimeSetBasicSheetsList[0].Name);
            Assert.AreEqual("TimeSet2", multiTimeSetSheets.TimeSetBasicSheetsList[1].Name);
            Assert.AreEqual("TimeSet3", multiTimeSetSheets.TimeSetBasicSheetsList[2].Name);
        }

        [TestMethod]
        public void MultiTimeSetSheets_AddTimeSetSheet_OrderPreserved()
        {
            // Arrange
            var multiTimeSetSheets = new MultiTimeSetSheets();

            // Act
            multiTimeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet("A"));
            multiTimeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet("B"));
            multiTimeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet("C"));

            // Assert
            Assert.AreEqual("A", multiTimeSetSheets.TimeSetBasicSheetsList[0].Name);
            Assert.AreEqual("B", multiTimeSetSheets.TimeSetBasicSheetsList[1].Name);
            Assert.AreEqual("C", multiTimeSetSheets.TimeSetBasicSheetsList[2].Name);
        }

        [TestMethod]
        public void MultiTimeSetSheets_GetTimeSetSheetByIndex()
        {
            // Arrange
            var multiTimeSetSheets = new MultiTimeSetSheets();
            var timeSet1 = new ComTimeSetBasicSheet("TimeSet1");
            var timeSet2 = new ComTimeSetBasicSheet("TimeSet2");
            multiTimeSetSheets.AddTimeSetSheet(timeSet1);
            multiTimeSetSheets.AddTimeSetSheet(timeSet2);

            // Act
            ComTimeSetBasicSheet retrievedTimeSet = multiTimeSetSheets.TimeSetBasicSheetsList[1];

            // Assert
            Assert.AreEqual("TimeSet2", retrievedTimeSet.Name);
        }

        [TestMethod]
        public void MultiTimeSetSheets_TimeSetBasicSheetsList_CanBeIterated()
        {
            // Arrange
            var multiTimeSetSheets = new MultiTimeSetSheets();
            multiTimeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet("TS1"));
            multiTimeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet("TS2"));
            multiTimeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet("TS3"));

            // Act
            int count = 0;
            foreach (ComTimeSetBasicSheet sheet in multiTimeSetSheets.TimeSetBasicSheetsList)
            {
                count++;
            }

            // Assert
            Assert.AreEqual(3, count);
        }

        [TestMethod]
        public void MultiTimeSetSheets_TimeSetBasicSheetsList_Count()
        {
            // Arrange
            var multiTimeSetSheets = new MultiTimeSetSheets();

            // Act & Assert
            Assert.AreEqual(0, multiTimeSetSheets.TimeSetBasicSheetsList.Count);

            multiTimeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet("TimeSet1"));
            Assert.AreEqual(1, multiTimeSetSheets.TimeSetBasicSheetsList.Count);

            multiTimeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet("TimeSet2"));
            Assert.AreEqual(2, multiTimeSetSheets.TimeSetBasicSheetsList.Count);
        }

        [TestMethod]
        public void MultiTimeSetSheets_WithDifferentTimingModes()
        {
            // Arrange
            var multiTimeSetSheets = new MultiTimeSetSheets();
            var asyncTimeSet = new ComTimeSetBasicSheet("AsyncTimeSet", "ASYNC");
            var syncTimeSet = new ComTimeSetBasicSheet("SyncTimeSet", "SYNC");

            // Act
            multiTimeSetSheets.AddTimeSetSheet(asyncTimeSet);
            multiTimeSetSheets.AddTimeSetSheet(syncTimeSet);

            // Assert
            Assert.AreEqual(2, multiTimeSetSheets.TimeSetBasicSheetsList.Count);
            Assert.AreEqual("ASYNC", multiTimeSetSheets.TimeSetBasicSheetsList[0].TimingMode);
            Assert.AreEqual("SYNC", multiTimeSetSheets.TimeSetBasicSheetsList[1].TimingMode);
        }
    }
}
