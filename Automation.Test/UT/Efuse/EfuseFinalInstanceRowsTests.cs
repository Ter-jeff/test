using Automation.GenerateIgxl.EFuse.Business;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Efuse
{
    [TestClass]
    public class EfuseFinalInstanceRowsTests
    {
        #region ReTestNameDuplicateRows

        [TestMethod]
        public void ReTestNameDuplicateRows_DifferentPayloadLists_AppendsRowNumSuffix()
        {
            var row1 = new EfuseFinalInstanceRow { TestName = "T1", TestNameTemp = "T1" };
            row1.EfusePatternRow.PayloadList = ["A", "B"];
            var row2 = new EfuseFinalInstanceRow { TestName = "T1", TestNameTemp = "T1" };
            row2.EfusePatternRow.PayloadList = ["A", "C"];
            row2.EfusePatternRow.RowNum = 5;

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.ReTestNameDuplicateRows();

            Assert.AreEqual("T1", row1.TestName);
            Assert.AreEqual("T1_RowNum5", row2.TestName);
        }

        [TestMethod]
        public void ReTestNameDuplicateRows_SamePayloadLists_PropagatesFirstRowTempName()
        {
            var row1 = new EfuseFinalInstanceRow { TestName = "T1", TestNameTemp = "T1_CANONICAL" };
            row1.EfusePatternRow.PayloadList = ["A", "B"];
            var row2 = new EfuseFinalInstanceRow { TestName = "T1", TestNameTemp = "T1_TEMP" };
            row2.EfusePatternRow.PayloadList = ["A", "B"];

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.ReTestNameDuplicateRows();

            Assert.AreEqual("T1_CANONICAL", row1.TestName);
            Assert.AreEqual("T1_CANONICAL", row2.TestName);
        }

        [TestMethod]
        public void ReTestNameDuplicateRows_InitItemsDifferentInitLists_AppendsRowNumSuffix()
        {
            var row1 = new EfuseFinalInstanceRow { TestName = "T1", TestNameTemp = "T1", InitPatName = "INIT1" };
            row1.EfusePatternRow.InitList = ["I1", "I2"];
            var row2 = new EfuseFinalInstanceRow { TestName = "T1", TestNameTemp = "T1", InitPatName = "INIT2" };
            row2.EfusePatternRow.InitList = ["I1", "I3"];
            row2.EfusePatternRow.RowNum = 7;

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.ReTestNameDuplicateRows();

            Assert.AreEqual("T1_RowNum7", row2.TestName);
        }

        [TestMethod]
        public void ReTestNameDuplicateRows_InitItemsSameInitLists_PropagatesFirstRowTempName()
        {
            var row1 = new EfuseFinalInstanceRow { TestName = "T1", TestNameTemp = "T1_CANONICAL", InitPatName = "INIT1" };
            row1.EfusePatternRow.InitList = ["I1", "I2"];
            var row2 = new EfuseFinalInstanceRow { TestName = "T1", TestNameTemp = "T1_TEMP", InitPatName = "INIT2" };
            row2.EfusePatternRow.InitList = ["I1", "I2"];

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.ReTestNameDuplicateRows();

            Assert.AreEqual("T1_CANONICAL", row2.TestName);
        }

        [TestMethod]
        public void ReTestNameDuplicateRows_OnlyOneRowHasInitPatName_TreatedAsPayloadItem()
        {
            // IsInitItem requires BOTH rows to have a non-empty InitPatName (logical AND); guards against
            // a mutation to OR, which would misroute this pair into the init-comparison branch instead.
            var row1 = new EfuseFinalInstanceRow { TestName = "T1", TestNameTemp = "T1", InitPatName = "INIT1" };
            row1.EfusePatternRow.PayloadList = ["A", "B"];
            var row2 = new EfuseFinalInstanceRow { TestName = "T1", TestNameTemp = "T1", InitPatName = "" };
            row2.EfusePatternRow.PayloadList = ["A", "C"];
            row2.EfusePatternRow.RowNum = 5;

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.ReTestNameDuplicateRows();

            Assert.AreEqual("T1_RowNum5", row2.TestName);
        }

        [TestMethod]
        public void ReTestNameDuplicateRows_DifferentTestNames_NotGroupedTogether()
        {
            var row1 = new EfuseFinalInstanceRow { TestName = "T1", TestNameTemp = "T1" };
            row1.EfusePatternRow.PayloadList = ["A"];
            var row2 = new EfuseFinalInstanceRow { TestName = "T2", TestNameTemp = "T2" };
            row2.EfusePatternRow.PayloadList = ["B"];

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.ReTestNameDuplicateRows();

            Assert.AreEqual("T1", row1.TestName);
            Assert.AreEqual("T2", row2.TestName);
        }

        [TestMethod]
        public void ReTestNameDuplicateRows_EmptyTestName_ExcludedFromGrouping()
        {
            var row1 = new EfuseFinalInstanceRow { TestName = "", TestNameTemp = "" };
            row1.EfusePatternRow.PayloadList = ["A"];
            var row2 = new EfuseFinalInstanceRow { TestName = "", TestNameTemp = "" };
            row2.EfusePatternRow.PayloadList = ["B"];
            row2.EfusePatternRow.RowNum = 7;

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.ReTestNameDuplicateRows();

            Assert.AreEqual("", row2.TestName);
        }

        #endregion

        #region RePatternNameDuplicateRows

        [TestMethod]
        public void RePatternNameDuplicateRows_PayloadDiffersAndNotInitAll_AppendsRowNumSuffix()
        {
            var row1 = new EfuseFinalInstanceRow { PatSetName = "SET_A", PatSetNameTemp = "SET_A" };
            row1.EfusePatternRow.PayloadList = ["P1", "P2"];
            var row2 = new EfuseFinalInstanceRow { PatSetName = "SET_A", PatSetNameTemp = "SET_A" };
            row2.EfusePatternRow.PayloadList = ["P1", "P3"];
            row2.EfusePatternRow.RowNum = 9;

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.RePatternNameDuplicateRows();

            Assert.AreEqual("SET_A", row1.PatSetName);
            Assert.AreEqual("SET_A_RowNum9", row2.PatSetName);
        }

        [TestMethod]
        public void RePatternNameDuplicateRows_PayloadSame_PropagatesFirstRowPatSetNameTemp()
        {
            var row1 = new EfuseFinalInstanceRow { PatSetName = "SET_B", PatSetNameTemp = "SET_B_CANONICAL" };
            row1.EfusePatternRow.PayloadList = ["P1", "P2"];
            var row2 = new EfuseFinalInstanceRow { PatSetName = "SET_B", PatSetNameTemp = "SET_B_TEMP" };
            row2.EfusePatternRow.PayloadList = ["P1", "P2"];

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.RePatternNameDuplicateRows();

            Assert.AreEqual("SET_B_CANONICAL", row1.PatSetName);
            Assert.AreEqual("SET_B_CANONICAL", row2.PatSetName);
        }

        [TestMethod]
        public void RePatternNameDuplicateRows_PatSetNameContainsInitAll_SkipsRenaming()
        {
            var row1 = new EfuseFinalInstanceRow { PatSetName = "Init_ALL_1", PatSetNameTemp = "Init_ALL_1" };
            row1.EfusePatternRow.PayloadList = ["P1", "P2"];
            var row2 = new EfuseFinalInstanceRow { PatSetName = "Init_ALL_1", PatSetNameTemp = "Init_ALL_1" };
            row2.EfusePatternRow.PayloadList = ["P1", "P3"];
            row2.EfusePatternRow.RowNum = 3;

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.RePatternNameDuplicateRows();

            Assert.AreEqual("Init_ALL_1", row2.PatSetName);
        }

        [TestMethod]
        public void RePatternNameDuplicateRows_InitItemsDiffer_AppendsRowNumSuffix()
        {
            var row1 = new EfuseFinalInstanceRow { PatSetName = "SET_C", PatSetNameTemp = "SET_C", InitPatName = "INIT1" };
            row1.EfusePatternRow.InitList = ["I1", "I2"];
            row1.EfusePatternRow.PayloadList = ["P1", "P2"];
            var row2 = new EfuseFinalInstanceRow { PatSetName = "SET_C", PatSetNameTemp = "SET_C", InitPatName = "INIT2" };
            row2.EfusePatternRow.InitList = ["I1", "I3"];
            row2.EfusePatternRow.PayloadList = ["P1", "P2"];
            row2.EfusePatternRow.RowNum = 4;

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.RePatternNameDuplicateRows();

            Assert.AreEqual("SET_C_RowNum4", row2.PatSetName);
        }

        [TestMethod]
        public void RePatternNameDuplicateRows_InitItemsDifferAndPatSetNameContainsInitAll_SkipsRenaming()
        {
            // Distinguishes && from || at line 79: PatSetName-equal is true but !Contains("Init_ALL") is
            // false, so only the AND form correctly skips renaming.
            var row1 = new EfuseFinalInstanceRow { PatSetName = "Init_ALL_1", PatSetNameTemp = "Init_ALL_1", InitPatName = "INIT1" };
            row1.EfusePatternRow.InitList = ["I1", "I2"];
            row1.EfusePatternRow.PayloadList = ["P1", "P2"];
            var row2 = new EfuseFinalInstanceRow { PatSetName = "Init_ALL_1", PatSetNameTemp = "Init_ALL_1", InitPatName = "INIT2" };
            row2.EfusePatternRow.InitList = ["I1", "I3"];
            row2.EfusePatternRow.PayloadList = ["P1", "P2"];
            row2.EfusePatternRow.RowNum = 3;

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.RePatternNameDuplicateRows();

            Assert.AreEqual("Init_ALL_1", row2.PatSetName);
        }

        [TestMethod]
        public void RePatternNameDuplicateRows_EmptyPatSetName_ExcludedFromGrouping()
        {
            // EfusePatternRow lists left empty, so PatSet.PatSetName stays "" and the row is filtered out
            // of grouping entirely. If the filter were bypassed, both rows would still group under the
            // same "" key and (with trivially-matching empty pattern lists) row1's temp name would
            // propagate to row2 - so a distinct sentinel value on row1 makes that leak observable.
            var row1 = new EfuseFinalInstanceRow { PatSetName = "", PatSetNameTemp = "SHOULD_NOT_PROPAGATE" };
            var row2 = new EfuseFinalInstanceRow { PatSetName = "", PatSetNameTemp = "" };

            var rows = new EfuseFinalInstanceRows { row1, row2 };
            rows.RePatternNameDuplicateRows();

            Assert.AreEqual("", row2.PatSetName);
        }

        #endregion
    }
}
