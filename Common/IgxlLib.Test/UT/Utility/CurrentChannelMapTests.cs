using System.Collections.Generic;
using System.IO;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.Utility.CurrentChannel;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

using Newtonsoft.Json;

namespace IgxlLib.Test.UT.Utility
{
    [TestClass]
    public class CurrentChannelMapTests
    {
        public string InputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        public string OutputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        public string ExpectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void CurrentChannelMapTest()
        {
            string subName = "CurrentChannelMap";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string file = Path.Combine(InputPath, "CurrentChannelMap.txt");
            CurrentChannelSheet currentChannel = CurrentChannelReader.ReadFile(file);

            Dictionary<string, string> dictionary = currentChannel.GetPogoMapping("DC-30");

            string json = JsonConvert.SerializeObject(currentChannel, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);
            string json1 = JsonConvert.SerializeObject(dictionary, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result1.json"), json1);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void CurrentChannelMapRow_DefaultConstructor_InitializesWithEmptyStrings()
        {
            // Arrange & Act
            var currentChannelMapRow = new CurrentChannelMapRow();

            // Assert
            Assert.AreEqual(string.Empty, currentChannelMapRow.TesterChannel);
            Assert.AreEqual(string.Empty, currentChannelMapRow.DibChannel);
            Assert.AreEqual(string.Empty, currentChannelMapRow.SignalName);
        }

        [TestMethod]
        public void CurrentChannelMapRow_SetProperties_UpdatesValues()
        {
            // Arrange
            var currentChannelMapRow = new CurrentChannelMapRow
            {
                // Act
                TesterChannel = "Channel1",
                DibChannel = "DibChan1",
                SignalName = "Signal1"
            };

            // Assert
            Assert.AreEqual("Channel1", currentChannelMapRow.TesterChannel);
            Assert.AreEqual("DibChan1", currentChannelMapRow.DibChannel);
            Assert.AreEqual("Signal1", currentChannelMapRow.SignalName);
        }

        [TestMethod]
        public void CurrentChannelMapRow_CopyConstructor_CopiesAllProperties()
        {
            // Arrange
            var original = new CurrentChannelMapRow
            {
                TesterChannel = "TesterChan1",
                DibChannel = "DibChan1",
                SignalName = "Signal1",
                SheetName = "Sheet1",
                RowNum = 5
            };

            // Act
            var copy = new CurrentChannelMapRow(original);

            // Assert
            Assert.AreEqual(original.TesterChannel, copy.TesterChannel);
            Assert.AreEqual(original.DibChannel, copy.DibChannel);
            Assert.AreEqual(original.SignalName, copy.SignalName);
            Assert.AreEqual(original.SheetName, copy.SheetName);
        }

        [TestMethod]
        public void CurrentChannelMapRow_Copy_CreatesIndependentCopy()
        {
            // Arrange
            var original = new CurrentChannelMapRow
            {
                TesterChannel = "Channel1",
                DibChannel = "DibChannel1",
                SignalName = "Signal1"
            };

            // Act
            CurrentChannelMapRow copied = original.Copy();

            // Assert
            Assert.AreEqual(original.TesterChannel, copied.TesterChannel);
            Assert.AreNotSame(original, copied);

            // Verify independence
            original.TesterChannel = "ModifiedChannel";
            Assert.AreEqual("Channel1", copied.TesterChannel);
        }

        [TestMethod]
        public void CurrentChannelMapRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var currentChannelMapRow = new CurrentChannelMapRow();

            // Assert
            Assert.IsInstanceOfType(currentChannelMapRow, typeof(IgxlRow));
        }
    }
}
