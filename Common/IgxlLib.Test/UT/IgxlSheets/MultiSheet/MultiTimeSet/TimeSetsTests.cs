using System.Collections.Generic;

using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets.MultiSheet.MultiTimeSet
{
    [TestClass]
    public class TimeSetsTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void TimeSetSheets_Constructor()
        {
            // Arrange & Act
            var timeSetSheets = new TimeSetSheets();

            // Assert
            Assert.IsNotNull(timeSetSheets);
            Assert.AreEqual(0, timeSetSheets.Count);
        }

        [TestMethod]
        public void TimeSetSheets_IsListOfComTimeSetBasicSheet()
        {
            // Arrange & Act
            var timeSetSheets = new TimeSetSheets();

            // Assert
            Assert.IsInstanceOfType(timeSetSheets, typeof(List<ComTimeSetBasicSheet>));
        }

        [TestMethod]
        public void TimeSetSheets_AddTimeSetSheet()
        {
            // Arrange
            var timeSetSheets = new TimeSetSheets();
            var comTimeSetBasicSheet = new ComTimeSetBasicSheet("TimeSet1");

            // Act
            timeSetSheets.AddTimeSetSheet(comTimeSetBasicSheet);

            // Assert
            Assert.AreEqual(1, timeSetSheets.Count);
            Assert.AreEqual("TimeSet1", timeSetSheets[0].Name);
        }

        [TestMethod]
        public void TimeSetSheets_AddMultipleTimeSetSheets()
        {
            // Arrange
            var timeSetSheets = new TimeSetSheets();
            var timeSet1 = new ComTimeSetBasicSheet("TimeSet1");
            var timeSet2 = new ComTimeSetBasicSheet("TimeSet2");
            var timeSet3 = new ComTimeSetBasicSheet("TimeSet3");

            // Act
            timeSetSheets.AddTimeSetSheet(timeSet1);
            timeSetSheets.AddTimeSetSheet(timeSet2);
            timeSetSheets.AddTimeSetSheet(timeSet3);

            // Assert
            Assert.AreEqual(3, timeSetSheets.Count);
            Assert.AreEqual("TimeSet1", timeSetSheets[0].Name);
            Assert.AreEqual("TimeSet2", timeSetSheets[1].Name);
            Assert.AreEqual("TimeSet3", timeSetSheets[2].Name);
        }

        [TestMethod]
        public void TimeSetSheets_UsesListIndexer()
        {
            // Arrange
            var timeSetSheets = new TimeSetSheets();
            var timeSet1 = new ComTimeSetBasicSheet("TimeSet1");
            var timeSet2 = new ComTimeSetBasicSheet("TimeSet2");

            // Act
            timeSetSheets.Add(timeSet1);
            timeSetSheets.Add(timeSet2);

            // Assert
            Assert.AreEqual("TimeSet1", timeSetSheets[0].Name);
            Assert.AreEqual("TimeSet2", timeSetSheets[1].Name);
        }

        [TestMethod]
        public void TimeSetSheets_Iteration()
        {
            // Arrange
            var timeSetSheets = new TimeSetSheets
            {
                new ComTimeSetBasicSheet("TS1"),
                new ComTimeSetBasicSheet("TS2"),
                new ComTimeSetBasicSheet("TS3")
            };

            // Act
            int count = 0;
            foreach (ComTimeSetBasicSheet sheet in timeSetSheets)
            {
                count++;
            }

            // Assert
            Assert.AreEqual(3, count);
        }

        [TestMethod]
        public void TimeSetSheets_Count()
        {
            // Arrange
            var timeSetSheets = new TimeSetSheets();

            // Act & Assert
            Assert.AreEqual(0, timeSetSheets.Count);

            timeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet("TimeSet1"));
            Assert.AreEqual(1, timeSetSheets.Count);

            timeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet("TimeSet2"));
            Assert.AreEqual(2, timeSetSheets.Count);
        }

        [TestMethod]
        public void TimeSetSheets_CanUseListMethods()
        {
            // Arrange
            var timeSetSheets = new TimeSetSheets();
            var timeSet1 = new ComTimeSetBasicSheet("TimeSet1");
            var timeSet2 = new ComTimeSetBasicSheet("TimeSet2");

            // Act
            timeSetSheets.Add(timeSet1);
            timeSetSheets.Add(timeSet2);

            bool contains = timeSetSheets.Contains(timeSet1);

            // Assert
            Assert.IsTrue(contains);
        }

        [TestMethod]
        public void TimeSetSheets_Remove()
        {
            // Arrange
            var timeSetSheets = new TimeSetSheets();
            var timeSet1 = new ComTimeSetBasicSheet("TimeSet1");
            var timeSet2 = new ComTimeSetBasicSheet("TimeSet2");
            timeSetSheets.Add(timeSet1);
            timeSetSheets.Add(timeSet2);

            // Act
            timeSetSheets.Remove(timeSet1);

            // Assert
            Assert.AreEqual(1, timeSetSheets.Count);
            Assert.AreEqual("TimeSet2", timeSetSheets[0].Name);
        }

        [TestMethod]
        public void TimeSetSheets_Clear()
        {
            // Arrange
            var timeSetSheets = new TimeSetSheets();
            timeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet("TimeSet1"));
            timeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet("TimeSet2"));

            // Act
            timeSetSheets.Clear();

            // Assert
            Assert.AreEqual(0, timeSetSheets.Count);
        }

        [TestMethod]
        public void TimeSetSheets_OrderPreserved()
        {
            // Arrange
            var timeSetSheets = new TimeSetSheets();

            // Act
            for (int i = 1; i <= 5; i++)
            {
                timeSetSheets.AddTimeSetSheet(new ComTimeSetBasicSheet($"TimeSet{i}"));
            }

            // Assert
            for (int i = 0; i < 5; i++)
            {
                Assert.AreEqual($"TimeSet{i + 1}", timeSetSheets[i].Name);
            }
        }
    }
}
