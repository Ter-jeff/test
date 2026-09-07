
using System.Collections.Generic;

using Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy;
using Automation.Static;

using IgxlLib;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using static Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy.ContiBase;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class ContiPowerShortTests : FunctionTestBase
    {
        [TestInitialize]
        public void Init()
        {
            ExcelWorksheet ws = EpWorkbook.TestPlanWorkbook.Worksheets["IO_PinMap"];
            var pinMapSheet = new PinMapSheet(ws);

            var pinGroup = new PinGroup("All_DiffPairs", "I/O");
            pinGroup.AddPin(new Pin("TX_P", "I/O"));
            pinGroup.AddPin(new Pin("TX_N", "I/O"));

            pinMapSheet.AddGroup(pinGroup);

            TestProgram.IgxlWorkBk = new IgxlWorkBook
            {
                PinMapPair = new KeyValuePair<string, PinMapSheet>(
                    "PINMAP",
                    pinMapSheet)
            };
        }

        [TestMethod]
        public void GetHighScale_uA_Returns_u()
        {
            // Arrange
            string limitUnit = "uA";

            // Act
            string result = ScaleHelper.GetHighScale(limitUnit);

            // Assert
            Assert.AreEqual("u", result);
        }

        [TestMethod]
        public void GetHighScale_CaseInsensitive_Returns_u()
        {
            // Arrange
            string limitUnit = "Ua";

            // Act
            string result = ScaleHelper.GetHighScale(limitUnit);

            // Assert
            Assert.AreEqual("u", result);
        }

        [TestMethod]
        public void GetHighScale_UnknownUnit_Returns_m()
        {
            // Arrange
            string limitUnit = "mA";

            // Act
            string result = ScaleHelper.GetHighScale(limitUnit);

            // Assert
            Assert.AreEqual("m", result);
        }

        [TestMethod]
        public void GetHighScale_NullUnit_Returns_m()
        {
            // Arrange
            string? limitUnit = null;

            // Act
            string result = ScaleHelper.GetHighScale(limitUnit);

            // Assert
            Assert.AreEqual("m", result);
        }

        [TestMethod]
        public void BuildLimitPins_WhenPinGroupIsNotNull_ShouldReturnPinNames()
        {
            // Arrange
            var pinGrp = new PinGroup("All_DiffPairs", "I/O");
            pinGrp.PinList.Add(new Pin { PinName = "All_DiffPairs" });
            pinGrp.PinList.Add(new Pin { PinName = "I/O" });

            // Act
            List<string> result = ContiPowerShort.BuildLimitPins(pinGrp);

            // Assert
            CollectionAssert.AreEqual(
                new List<string> { "All_DiffPairs", "I/O" },
                result
            );
        }

        [TestMethod]
        public void BuildLimitPins_WhenPinGroupIsNull_ShouldReturnEmptyList()
        {
            // Act
            List<string> result = ContiPowerShort.BuildLimitPins(null);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(0, result.Count);
        }
        private static readonly string[] _expected =
                [
                    "FAIL1",
                    "FAIL2",
                    "FAIL3",
                    "FAIL4"
                ];

        [TestMethod]
        public void GetBinFlags_Should_Remove_FilteredFlags_TrimPrefix_And_Distinct()
        {
            // Arrange
            var flowRows = new List<FlowRow>
            {
                new() { FailAction = "AAA,F_FAIL1,F_FAIL2" },
                new() { FailAction = "fail2,FAIL3" },
                new() { FailAction = " BBB , F_FAIL4 " },
                new() { FailAction = "" }
            };

            // Act
            List<string> result = ContiPowerShort.GetBinFlags(flowRows, ["AAA", "BBB"]);

            // Assert
            CollectionAssert.AreEquivalent(_expected, result);
        }
    }
}
