using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.PostAction.TempMon.Data;
using Automation.GenerateIgxl.PostAction.TempMon.Enums;
using Automation.Utility.Basic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Utility
{
    [TestClass]
    public class TempMonServiceTests
    {
        // -----------------------------
        // TryGetTempMonSyntax - Valid
        // -----------------------------

        [DataTestMethod]
        [DataRow("TempMon_ENG:Enable", "ENG", EnumCondition.Include)]
        [DataRow("TempMon_ENG:Disable", "ENG", EnumCondition.Exclude)]
        [DataRow("TempMon_eng:enable", "ENG", EnumCondition.Include)]
        [DataRow("TempMon_TEST:Enable ", "TEST", EnumCondition.Include)]
        [DataRow("TempMon_XYZ:Disable;", "XYZ", EnumCondition.Exclude)]
        [DataRow("prefix;TempMon_ABC:Enable,", "ABC", EnumCondition.Include)]
        public void TryGetTempMonSyntax_Valid(string input, string expectedMode, EnumCondition enumCondition)
        {
            bool result = TempMonService.TryGetTempMonSyntax(input, out string mode, out EnumCondition condition);

            Assert.IsTrue(result);
            Assert.AreEqual(expectedMode, mode);
            Assert.AreEqual(enumCondition, condition);
        }

        // -----------------------------
        // TryGetTempMonSyntax - Invalid
        // -----------------------------

        [DataTestMethod]
        [DataRow(null)]
        [DataRow("")]
        [DataRow("   ")]
        [DataRow("TempMon_ENG:EnableENG")]   // should NOT match
        [DataRow("TempMon_ENG:Enab")]
        [DataRow("TempMon_:Enable")]
        [DataRow("RandomText")]
        public void TryGetTempMonSyntax_Invalid(string input)
        {
            bool result = TempMonService.TryGetTempMonSyntax(input, out string mode, out EnumCondition condition);

            Assert.IsFalse(result);
            Assert.AreEqual(string.Empty, mode);
            Assert.AreEqual(EnumCondition.Unknown, condition);
        }

        // -----------------------------
        // TryGetTempMonSyntax - Boundary check
        // -----------------------------

        [TestMethod]
        public void TryGetTempMonSyntax_ShouldNotOverMatch()
        {
            // invalid tail
            string input = "TempMon_ENG:EnableXYZ";

            bool result = TempMonService.TryGetTempMonSyntax(input, out _, out _);

            Assert.IsFalse(result);
        }

        // -----------------------------
        // TrySetTempMon
        // -----------------------------

        [TestMethod]
        public void TrySetTempMon_Valid_AddsData()
        {
            var set = new HashSet<TempMonData>();

            bool check = TempMonService.TrySetTempMon(set, "TempMon_ENG:Enable", "Item1", EnumType.Instance);
            TempMonService.TrySetTempMon(set, "TempMon_ENG:Enable", "Item2", EnumType.Flow);

            Assert.IsTrue(check);
            Assert.AreEqual(2, set.Count);

            TempMonData data = set.Last();

            Assert.AreEqual("ENG", data.Mode);
            Assert.AreEqual(EnumCondition.Include, data.Condition);
            Assert.AreEqual("Item2", data.Item);
            Assert.AreEqual(EnumType.Flow, data.Type);
        }

        [TestMethod]
        public void TrySetTempMon_Invalid_DoesNothing()
        {
            var set = new HashSet<TempMonData>();

            bool result = TempMonService.TrySetTempMon(set, "InvalidSyntax", "Item1", EnumType.Flow);

            Assert.IsFalse(result);
            Assert.AreEqual(0, set.Count);
        }

        // -----------------------------
        // SetTempMon
        // -----------------------------

        [TestMethod]
        public void SetTempMon_AddsEntry()
        {
            var set = new HashSet<TempMonData>();

            TempMonService.SetTempMon(set, "ENG", EnumCondition.Include, EnumType.Instance, "ItemX");
            TempMonService.SetTempMon(set, "ENG", EnumCondition.Include, EnumType.Flow, "ItemX");

            Assert.AreEqual(2, set.Count);

            TempMonData data = set.First();

            Assert.AreEqual("ENG", data.Mode);
            Assert.AreEqual(EnumCondition.Include, data.Condition);
            Assert.AreEqual(EnumType.Instance, data.Type);
            Assert.AreEqual("ItemX", data.Item);
        }

        // -----------------------------
        // HashSet behavior
        // -----------------------------

        [TestMethod]
        public void TrySetTempMon_Duplicate_ShouldNotDuplicate()
        {
            var set = new HashSet<TempMonData>();

            TempMonService.TrySetTempMon(set, "TempMon_ENG:Enable", "Item1", EnumType.Instance);
            TempMonService.TrySetTempMon(set, "TempMon_ENG:Enable", "Item1", EnumType.Instance);

            Assert.AreEqual(1, set.Count);
        }

        [TestMethod]
        public void TrySetTempMon_DifferentMode_ShouldAddBoth()
        {
            var set = new HashSet<TempMonData>();

            TempMonService.TrySetTempMon(set, "TempMon_ENG:Enable", "Item1", EnumType.Instance);
            TempMonService.TrySetTempMon(set, "TempMon_XYZ:Enable", "Item1", EnumType.Instance);

            Assert.AreEqual(2, set.Count);
        }
    }
}
