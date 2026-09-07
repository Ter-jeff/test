using System.Collections.Generic;
using System.Linq;

using Automation.Reader.ConfigFile.RtosTable;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class GenerateIgxlMainTests
    {
        private static IGrouping<string, RtosTableArgRow> NewGrouping(params RtosTableArgRow[] rows)
        {
            return rows.GroupBy(x => x.SheetName).First();
        }

        [DataTestMethod]
        [DataRow("DC Category", DisplayName = "DcCategorySpaced")]
        [DataRow("DCCategory", DisplayName = "DcCategoryCompact")]
        public void PopulateDc_DcCategoryArg_SetsDcCategory(string argName)
        {
            // Arrange
            var instanceRow = new InstanceRow { ArgList = "A,B,C", Args = ["", "", ""] };
            IGrouping<string, RtosTableArgRow> grouping = NewGrouping(new RtosTableArgRow("Sheet1", "Func1", argName, "CAT1"));

            // Act
            GenerateIgxlMain.PopulateDc(grouping, instanceRow);

            // Assert
            Assert.AreEqual("CAT1", instanceRow.DcCategory);
        }

        [DataTestMethod]
        [DataRow("DC Selector", DisplayName = "DcSelectorSpaced")]
        [DataRow("DCSelector", DisplayName = "DcSelectorCompact")]
        public void PopulateDc_DcSelectorArg_SetsDcSelector(string argName)
        {
            // Arrange
            var instanceRow = new InstanceRow { ArgList = "A,B,C", Args = ["", "", ""] };
            IGrouping<string, RtosTableArgRow> grouping = NewGrouping(new RtosTableArgRow("Sheet1", "Func1", argName, "SEL1"));

            // Act
            GenerateIgxlMain.PopulateDc(grouping, instanceRow);

            // Assert
            Assert.AreEqual("SEL1", instanceRow.DcSelector);
        }

        [DataTestMethod]
        [DataRow("AC Category", DisplayName = "AcCategorySpaced")]
        [DataRow("ACCategory", DisplayName = "AcCategoryCompact")]
        public void PopulateDc_AcCategoryArg_SetsAcCategory(string argName)
        {
            // Arrange
            var instanceRow = new InstanceRow { ArgList = "A,B,C", Args = ["", "", ""] };
            IGrouping<string, RtosTableArgRow> grouping = NewGrouping(new RtosTableArgRow("Sheet1", "Func1", argName, "ACAT1"));

            // Act
            GenerateIgxlMain.PopulateDc(grouping, instanceRow);

            // Assert
            Assert.AreEqual("ACAT1", instanceRow.AcCategory);
        }

        [DataTestMethod]
        [DataRow("AC Selector", DisplayName = "AcSelectorSpaced")]
        [DataRow("ACSelector", DisplayName = "AcSelectorCompact")]
        public void PopulateDc_AcSelectorArg_SetsAcSelector(string argName)
        {
            // Arrange
            var instanceRow = new InstanceRow { ArgList = "A,B,C", Args = ["", "", ""] };
            IGrouping<string, RtosTableArgRow> grouping = NewGrouping(new RtosTableArgRow("Sheet1", "Func1", argName, "ASEL1"));

            // Act
            GenerateIgxlMain.PopulateDc(grouping, instanceRow);

            // Assert
            Assert.AreEqual("ASEL1", instanceRow.AcSelector);
        }

        [DataTestMethod]
        [DataRow("Time Sets", DisplayName = "TimeSetsSpaced")]
        [DataRow("TimeSets", DisplayName = "TimeSetsCompact")]
        public void PopulateDc_TimeSetsArg_SetsTimeSets(string argName)
        {
            // Arrange
            var instanceRow = new InstanceRow { ArgList = "A,B,C", Args = ["", "", ""] };
            IGrouping<string, RtosTableArgRow> grouping = NewGrouping(new RtosTableArgRow("Sheet1", "Func1", argName, "TS1"));

            // Act
            GenerateIgxlMain.PopulateDc(grouping, instanceRow);

            // Assert
            Assert.AreEqual("TS1", instanceRow.TimeSets);
        }

        [DataTestMethod]
        [DataRow("Pin Levels", DisplayName = "PinLevelsSpaced")]
        [DataRow("PinLevels", DisplayName = "PinLevelsCompact")]
        public void PopulateDc_PinLevelsArg_SetsPinLevels(string argName)
        {
            // Arrange
            var instanceRow = new InstanceRow { ArgList = "A,B,C", Args = ["", "", ""] };
            IGrouping<string, RtosTableArgRow> grouping = NewGrouping(new RtosTableArgRow("Sheet1", "Func1", argName, "PL1"));

            // Act
            GenerateIgxlMain.PopulateDc(grouping, instanceRow);

            // Assert
            Assert.AreEqual("PL1", instanceRow.PinLevels);
        }

        [TestMethod]
        public void PopulateDc_CustomArgNameMatchesArgListPosition_UpdatesArgsAtThatIndex()
        {
            // Arrange
            var instanceRow = new InstanceRow { ArgList = "Voltage,Freq,Custom1", Args = ["V0", "F0", "C0"] };
            IGrouping<string, RtosTableArgRow> grouping = NewGrouping(new RtosTableArgRow("Sheet1", "Func1", "Custom1", "NEWVAL"));

            // Act
            GenerateIgxlMain.PopulateDc(grouping, instanceRow);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "V0", "F0", "NEWVAL" }, instanceRow.Args);
        }

        [TestMethod]
        public void PopulateDc_ArgNameNotFoundInArgList_LeavesArgsUnchanged()
        {
            // Arrange
            var instanceRow = new InstanceRow { ArgList = "Voltage,Freq", Args = ["V0", "F0"] };
            IGrouping<string, RtosTableArgRow> grouping = NewGrouping(new RtosTableArgRow("Sheet1", "Func1", "NotFound", "NEWVAL"));

            // Act
            GenerateIgxlMain.PopulateDc(grouping, instanceRow);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "V0", "F0" }, instanceRow.Args);
        }
    }
}
