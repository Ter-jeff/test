using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.PostAction.UpdatePatternSetTimeDomain;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class UpdateTimeDomainMainTests
    {
        [TestMethod]
        public void UpdateTimeDomainForNonUltraFlexTest()
        {
            //Arrange
            PatSetSheet patSetsAllSheet = ArrangePatSetsAllSheet();
            IEnumerable<PatSetSheet> patSetSheets = ArrangePatSetSheets();
            IEnumerable<TimeSetBasicSheet> timeSetSheets = ArrangeTimeSetSheets();
            List<EnumEquipment> equipmentList = [EnumEquipment.UltraFlexPlus];

            // Act
            var updateTimeDomainMain = new UpdateTimeDomainMain(equipmentList);
            updateTimeDomainMain.UpdatePatternSetSheets(ref patSetsAllSheet, ref patSetSheets);
            updateTimeDomainMain.UpdateTimeSetSheets(ref timeSetSheets);

            // Assert
            Assert.IsTrue(patSetSheets.SelectMany(x => x.Rows).SelectMany(x => x.PatSetRows).Where(x => x.File.EqualsIgnoreCase("A123")).All(x => x.TimeDomain.EqualsIgnoreCase("domain1,domain2,domain5,domain6")));
            Assert.IsTrue(patSetSheets.SelectMany(x => x.Rows).SelectMany(x => x.PatSetRows).Where(x => x.File.EqualsIgnoreCase("B123")).All(x => x.TimeDomain.EqualsIgnoreCase("domain3,domain4")));
            TimeSetBasicSheet? timeSetA = timeSetSheets.FirstOrDefault(x => x.Name.EqualsIgnoreCase("TIMESET_A"));
            Assert.AreNotEqual(null, timeSetA);
            Assert.IsTrue(timeSetA!.TimeDomain == "TimeDomainA");
            Assert.IsTrue(string.IsNullOrEmpty(timeSetSheets.FirstOrDefault(x => x.Name.EqualsIgnoreCase("TIMESET_B"))?.TimeDomain));
            TimeSetBasicSheet? timeSetC = timeSetSheets.FirstOrDefault(x => x.Name.EqualsIgnoreCase("TIMESET_C"));
            Assert.AreNotEqual(null, timeSetC);
            Assert.IsTrue(timeSetC!.TimeDomain == "TimeDomainC");
        }

        [TestMethod]
        public void UpdateTimeDomainForUltraFlexTest()
        {
            //Arrange
            PatSetSheet patSetsAllSheet = ArrangePatSetsAllSheet();
            IEnumerable<PatSetSheet> patSetSheets = ArrangePatSetSheets();
            IEnumerable<TimeSetBasicSheet> timeSetSheets = ArrangeTimeSetSheets();
            List<EnumEquipment> equipmentList = [EnumEquipment.UltraFlex, EnumEquipment.UltraFlexPlus];

            // Act
            var updateTimeDomainMain = new UpdateTimeDomainMain(equipmentList);
            updateTimeDomainMain.UpdatePatternSetSheets(ref patSetsAllSheet, ref patSetSheets);
            updateTimeDomainMain.UpdateTimeSetSheets(ref timeSetSheets);

            // Assert
            Assert.IsTrue(patSetsAllSheet.Rows.SelectMany(x => x.PatSetRows).All(x => string.IsNullOrEmpty(x.TimeDomain)));
            Assert.IsTrue(patSetSheets.SelectMany(x => x.Rows).SelectMany(x => x.PatSetRows).Where(x => x.File.EqualsIgnoreCase("A123")).All(x => string.IsNullOrEmpty(x.TimeDomain)));
            Assert.IsTrue(patSetSheets.SelectMany(x => x.Rows).SelectMany(x => x.PatSetRows).Where(x => x.File.EqualsIgnoreCase("B123")).All(x => string.IsNullOrEmpty(x.TimeDomain)));
            Assert.IsTrue(string.IsNullOrEmpty(timeSetSheets.FirstOrDefault(x => x.Name.EqualsIgnoreCase("TIMESET_A"))?.TimeDomain));
            Assert.IsTrue(string.IsNullOrEmpty(timeSetSheets.FirstOrDefault(x => x.Name.EqualsIgnoreCase("TIMESET_A"))?.TimeDomain));
            Assert.IsTrue(string.IsNullOrEmpty(timeSetSheets.FirstOrDefault(x => x.Name.EqualsIgnoreCase("TIMESET_A"))?.TimeDomain));
        }

        private static PatSetSheet ArrangePatSetsAllSheet()
        {
            PatSetSheet patSetsAll = new PatSetSheet("PatSets_All")
            {
                Rows =
                [
                    new()
                    {
                        PatSetName = "A123",
                        PatSetRows =
                        [
                            new() { TimeDomain = "domain1,domain2" }
                        ]
                    },
                    new()
                    {
                        PatSetName = "B123",
                        PatSetRows =
                        [
                            new() { TimeDomain = "domain3" },
                            new() { TimeDomain = "domain4" }
                        ]
                    },
                    new()
                    {
                        PatSetName = "A123",
                        PatSetRows =
                        [
                            new() { TimeDomain = "domain5,domain6" }
                        ]
                    }
                ]
            };
            return patSetsAll;
        }

        private static IEnumerable<PatSetSheet> ArrangePatSetSheets()
        {
            IEnumerable<PatSetSheet> patSetSheets =
            [
                new("PatSet_Scan")
                {
                    Rows =
                    [
                        new()
                        {
                            PatSetName = "A",
                            PatSetRows =
                            [
                                new() { File = "A123" },
                                new() { File = "B456" }
                            ]
                        },
                    ]
                },
                new("PatSet_HIP")
                {
                    Rows =
                    [
                        new()
                        {
                            PatSetName = "B",
                            PatSetRows =
                            [
                                new() { File = "b123" },
                                new() { File = "a123" },
                                new() { File = "A123" },
                            ]
                        },
                    ]
                },
            ];
            return patSetSheets;
        }

        private static IEnumerable<TimeSetBasicSheet> ArrangeTimeSetSheets()
        {
            IEnumerable<TimeSetBasicSheet> timeSetSheets =
            [
                new("TIMESET_A")
                {
                    TimeDomain = "TimeDomainA"
                },
                new("TIMESET_B")
                {
                    TimeDomain = ""
                },
                new("TIMESET_C")
                {
                    TimeDomain = "TimeDomainC"
                }
            ];
            return timeSetSheets;
        }
    }
}
