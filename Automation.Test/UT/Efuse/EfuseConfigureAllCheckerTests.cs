using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.EFuse.InputChecker;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Efuse.Input;

namespace Automation.Test.UT.Efuse
{
    [TestClass]
    public class EfuseConfigureAllCheckerTests
    {
        private readonly EfuseConfigureAllChecker _checker = new();

        [TestInitialize]
        public void Setup()
        {
            ErrorReportManager.ClearErrors();
        }

        [TestCleanup]
        public void Cleanup()
        {
            ErrorReportManager.ClearErrors();
        }

        private static BitDefTable NewTable(string blockName)
        {
            return new BitDefTable
            {
                SheetName = "BdfSheet",
                BlockName = blockName,
                HeaderRowNum = 1,
                BankEfuseBitDefIdx = 0,
                MsbBitIdx = 1,
                LsbBitIdx = 2,
                BitWidthIdx = 3,
                ProgrammingStageIdx = 4,
                DefaultOrRealIdx = 5,
                LowLimitIdx = 6,
                HighLimitIdx = 7,
                DefaultValueIdx = 8,
                AlgorithmIdx = 9,
                DescriptionIdx = 10,
                ResolutionIdx = 11
            };
        }

        private static BitDefRow NewBdfRow(SortedDictionary<int, string> values)
        {
            var row = new BitDefRow { SheetName = "BdfSheet", RowNum = 2 };
            int max = -1;
            foreach (int key in values.Keys)
            {
                if (key > max)
                {
                    max = key;
                }
            }

            for (int i = 0; i <= max; i++)
            {
                row.RowData.Add(values.TryGetValue(i, out string? value) ? value : "");
            }

            return row;
        }

        private static EfuseConfigMainSheet NewSheet(string sheetName, params EfuseConfigMainRow[] rows)
        {
            var sheet = new EfuseConfigMainSheet(sheetName);
            sheet.Rows.AddRange(rows);
            return sheet;
        }

        private static List<Error> Errors()
        {
            return ErrorReportManager.GetErrorList();
        }

        [TestMethod]
        public void WorkFlow_ComputesMaxBitsAndFlagsSheetWithLowerConditionMsb()
        {
            BitDefTable bdfTable = NewTable("CFG");
            bdfTable.Rows.Add(NewBdfRow(new SortedDictionary<int, string> { [1] = "3", [9] = "cond" }));
            bdfTable.Rows.Add(NewBdfRow(new SortedDictionary<int, string> { [1] = "7", [9] = "cond" }));
            bdfTable.Rows.Add(NewBdfRow(new SortedDictionary<int, string> { [1] = "99", [9] = "other" }));

            EfuseConfigMainSheet sheetGood = NewSheet("SheetGood", new EfuseConfigMainRow { Description = "CFG_condition_X", Msb = 7 });
            EfuseConfigMainSheet sheetBad = NewSheet("SheetBad",
                new EfuseConfigMainRow { Description = "CFG_condition_Y", Msb = 2 },
                new EfuseConfigMainRow { Description = "CFG_condition_Z", Msb = 5 });

            _checker.WorkFlow([sheetGood, sheetBad], [bdfTable]);

            List<Error> maxBitsErrors = [.. Errors().Where(e => e.ErrorCode.FullCode == EFuseErrorType.E_InvalidMaximumBits_01.FullCode)];
            Assert.AreEqual(1, maxBitsErrors.Count);
            Assert.AreEqual("SheetBad", maxBitsErrors[0].SheetName);
            // The reported MSB must be the maximum among the CFG_condition_ rows (5), not the minimum (2).
            CollectionAssert.AreEqual(new[] { "5", "7" }, maxBitsErrors[0].MessageArgs);
        }

        [TestMethod]
        public void WorkFlow_NonConditionRowHasHighMsb_StillFlagsSheetBasedOnConditionRowsOnly()
        {
            // Guards the "CFG_condition_" prefix check inside the Exists/Any predicates: a mutation to an
            // empty-string prefix would let this high-MSB, wrongly-named row satisfy the "already OK"
            // check and incorrectly suppress the error.
            BitDefTable bdfTable = NewTable("CFG");
            bdfTable.Rows.Add(NewBdfRow(new SortedDictionary<int, string> { [1] = "7", [9] = "cond" }));

            EfuseConfigMainSheet sheet = NewSheet("SheetBad",
                new EfuseConfigMainRow { Description = "CFG_condition_X", Msb = 3 },
                new EfuseConfigMainRow { Description = "UnrelatedRow", Msb = 100 });

            _checker.WorkFlow([sheet], [bdfTable]);

            Error error = Errors().Single(e => e.ErrorCode.FullCode == EFuseErrorType.E_InvalidMaximumBits_01.FullCode);
            CollectionAssert.AreEqual(new[] { "3", "7" }, error.MessageArgs);
        }

        [TestMethod]
        public void WorkFlow_NoMatchingCfgBdfTable_MaxBitsDefaultsToZero_NonNegativeMsbSatisfiesCheck()
        {
            BitDefTable bdfTable = NewTable("ECID");
            bdfTable.Rows.Add(NewBdfRow(new SortedDictionary<int, string> { [1] = "5", [9] = "cond" }));

            EfuseConfigMainSheet sheet = NewSheet("Sheet1", new EfuseConfigMainRow { Description = "CFG_condition_X", Msb = 0 });

            _checker.WorkFlow([sheet], [bdfTable]);

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == EFuseErrorType.E_InvalidMaximumBits_01.FullCode));
        }

        [TestMethod]
        public void WorkFlow_SheetWithNoConditionRows_NoConditionMsbError()
        {
            BitDefTable bdfTable = NewTable("CFG");
            bdfTable.Rows.Add(NewBdfRow(new SortedDictionary<int, string> { [1] = "5", [9] = "cond" }));

            EfuseConfigMainSheet sheet = NewSheet("Sheet1", new EfuseConfigMainRow { Description = "SomeOtherDescription", Msb = 1 });

            _checker.WorkFlow([sheet], [bdfTable]);

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == EFuseErrorType.E_InvalidMaximumBits_01.FullCode));
        }

        [TestMethod]
        public void WorkFlow_SheetsWithMismatchedRowCounts_AddsMismatchRowError()
        {
            EfuseConfigMainSheet sheetA = NewSheet("SheetA", new EfuseConfigMainRow(), new EfuseConfigMainRow());
            EfuseConfigMainSheet sheetB = NewSheet("SheetB", new EfuseConfigMainRow(), new EfuseConfigMainRow(), new EfuseConfigMainRow());

            _checker.WorkFlow([sheetA, sheetB], []);

            Error error = Errors().Single(e => e.ErrorCode.FullCode == EFuseErrorType.E_MismatchRow_01.FullCode);
            CollectionAssert.AreEqual(new[] { "SheetA", "2", "SheetB", "3" }, error.MessageArgs);
            // The mismatch is also recorded on the sheet's own local error list, not just the global one.
            Assert.IsTrue(sheetA.GetErrors().Any(e => e.ErrorCode.FullCode == EFuseErrorType.E_MismatchRow_01.FullCode));
        }

        [TestMethod]
        public void WorkFlow_SheetsWithSameRowCounts_NoMismatchRowError()
        {
            EfuseConfigMainSheet sheetA = NewSheet("SheetA", new EfuseConfigMainRow(), new EfuseConfigMainRow());
            EfuseConfigMainSheet sheetB = NewSheet("SheetB", new EfuseConfigMainRow(), new EfuseConfigMainRow());

            _checker.WorkFlow([sheetA, sheetB], []);

            Assert.IsFalse(Errors().Any(e => e.ErrorCode.FullCode == EFuseErrorType.E_MismatchRow_01.FullCode));
        }

        [TestMethod]
        public void WorkFlow_ThreeSheetsMixedCounts_AddsErrorForEachMismatchedPairOnly()
        {
            EfuseConfigMainSheet sheetA = NewSheet("SheetA", new EfuseConfigMainRow());
            EfuseConfigMainSheet sheetB = NewSheet("SheetB", new EfuseConfigMainRow());
            EfuseConfigMainSheet sheetC = NewSheet("SheetC", new EfuseConfigMainRow(), new EfuseConfigMainRow());

            _checker.WorkFlow([sheetA, sheetB, sheetC], []);

            int count = Errors().Count(e => e.ErrorCode.FullCode == EFuseErrorType.E_MismatchRow_01.FullCode);
            Assert.AreEqual(2, count);
        }
    }
}
