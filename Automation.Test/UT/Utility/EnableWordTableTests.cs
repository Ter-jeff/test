using System;
using System.Collections.Generic;

using Automation.Utility.TpUpdate.HardIPEnableWordsUpdate;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Utility
{
    [TestClass]
    public class EnableWordTableTests
    {
        private static EnableWordTableRow MakeRow(string name, Dictionary<string, List<string>>? enableWords = null)
        {
            var headerMetadata = new Dictionary<string, TtrColumnMetadata>
            {
                ["Pattern"] = new TtrColumnMetadata("Pattern", 1),
                ["Category"] = new TtrColumnMetadata("Category", 2),
                ["SubBlock"] = new TtrColumnMetadata("SubBlock", 3)
            };
            var dataset = new Dictionary<int, string> { [1] = name, [2] = "CAT", [3] = "SUB" };
            var row = new EnableWordTableRow("Pattern", "Category", "SubBlock", "HIP_FailRun", headerMetadata, dataset);
            if (enableWords != null)
            {
                row.EnableWords = new Dictionary<string, List<string>>(enableWords, StringComparer.OrdinalIgnoreCase);
            }
            return row;
        }

        #region ReadTable

        [TestMethod]
        public void ReadTable_NoDimension_LeavesIsValidFalse()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            var table = new EnableWordTable();

            table.ReadTable(sheet);

            Assert.IsFalse(table.IsValid);
            Assert.AreEqual(0, table.Rows.Count);
        }

        [TestMethod]
        public void ReadTable_NoProposalForTtrCell_LeavesIsValidFalse()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "SomeUnrelatedHeader";
            var table = new EnableWordTable();

            table.ReadTable(sheet);

            Assert.IsFalse(table.IsValid);
        }

        [TestMethod]
        public void ReadTable_ProposalFoundButNoValidHeaderRow_LeavesIsValidFalse()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Proposal for TTR";
            sheet.Cells[2, 1].Value = "NotAHeaderRow";

            var table = new EnableWordTable();
            table.ReadTable(sheet);

            Assert.IsFalse(table.IsValid);
        }

        [TestMethod]
        public void ReadTable_ValidHeaderAndData_SetsIsValidTrueAndPopulatesRows()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Proposal for TTR";
            sheet.Cells[2, 1].Value = "Pattern";
            sheet.Cells[2, 2].Value = "Category";
            sheet.Cells[2, 3].Value = "SubBlock";
            sheet.Cells[3, 1].Value = "PAT1";
            sheet.Cells[3, 2].Value = "CAT1";
            sheet.Cells[3, 3].Value = "SUB1";

            var table = new EnableWordTable();
            table.ReadTable(sheet);

            Assert.IsTrue(table.IsValid);
            Assert.AreEqual(1, table.Rows.Count);
            Assert.AreEqual("PAT1", table.Rows[0].Name);
            Assert.AreEqual(4, table.Rows[0].RowNum);
            Assert.IsTrue(table.IsContainsSubBlock());
        }

        [TestMethod]
        public void ReadTable_BlankSubBlockValue_SkipsRow()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Proposal for TTR";
            sheet.Cells[2, 1].Value = "Pattern";
            sheet.Cells[2, 2].Value = "Category";
            sheet.Cells[2, 3].Value = "SubBlock";
            sheet.Cells[3, 1].Value = "PAT1";
            sheet.Cells[3, 2].Value = "CAT1";
            sheet.Cells[3, 3].Value = "";

            var table = new EnableWordTable();
            table.ReadTable(sheet);

            Assert.IsTrue(table.IsValid);
            Assert.AreEqual(0, table.Rows.Count);
        }

        [TestMethod]
        public void ReadTable_DuplicateNameAndSubBlock_LaterRowReplacesEarlierOne()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Proposal for TTR";
            sheet.Cells[2, 1].Value = "Pattern";
            sheet.Cells[2, 2].Value = "Category";
            sheet.Cells[2, 3].Value = "SubBlock";
            sheet.Cells[3, 1].Value = "PAT1";
            sheet.Cells[3, 2].Value = "CAT_FIRST";
            sheet.Cells[3, 3].Value = "SUB1";
            sheet.Cells[4, 1].Value = "PAT1";
            sheet.Cells[4, 2].Value = "CAT_SECOND";
            sheet.Cells[4, 3].Value = "SUB1";

            var table = new EnableWordTable();
            table.ReadTable(sheet);

            Assert.AreEqual(1, table.Rows.Count);
            Assert.AreEqual("CAT_SECOND", table.Rows[0].Category);
        }

        [TestMethod]
        public void ReadTable_SameNameDifferentSubBlock_BothRowsRetained()
        {
            // The duplicate lookup requires Name AND SubBlock to match (logical AND); guards against
            // a mutation to OR, which would incorrectly treat same-Name-different-SubBlock as a duplicate.
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Proposal for TTR";
            sheet.Cells[2, 1].Value = "Pattern";
            sheet.Cells[2, 2].Value = "Category";
            sheet.Cells[2, 3].Value = "SubBlock";
            sheet.Cells[3, 1].Value = "PAT1";
            sheet.Cells[3, 2].Value = "CAT_FIRST";
            sheet.Cells[3, 3].Value = "SUB1";
            sheet.Cells[4, 1].Value = "PAT1";
            sheet.Cells[4, 2].Value = "CAT_SECOND";
            sheet.Cells[4, 3].Value = "SUB2";

            var table = new EnableWordTable();
            table.ReadTable(sheet);

            Assert.AreEqual(2, table.Rows.Count);
        }

        [TestMethod]
        public void ReadTable_MissingCategoryHeader_LeavesIsValidFalse()
        {
            // Category is a required header; guards against a mutation that marks it optional.
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Proposal for TTR";
            sheet.Cells[2, 1].Value = "Pattern";
            sheet.Cells[2, 2].Value = "SubBlock";
            sheet.Cells[3, 1].Value = "PAT1";
            sheet.Cells[3, 2].Value = "SUB1";

            var table = new EnableWordTable();
            table.ReadTable(sheet);

            Assert.IsFalse(table.IsValid);
        }

        [TestMethod]
        public void ReadTable_MissingSubBlockHeader_LeavesIsValidFalse()
        {
            // SubBlock is a required header; guards against a mutation that marks it optional.
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Proposal for TTR";
            sheet.Cells[2, 1].Value = "Pattern";
            sheet.Cells[2, 2].Value = "Category";
            sheet.Cells[3, 1].Value = "PAT1";
            sheet.Cells[3, 2].Value = "CAT1";

            var table = new EnableWordTable();
            table.ReadTable(sheet);

            Assert.IsFalse(table.IsValid);
        }

        [TestMethod]
        public void ReadTable_BlankSubBlockRowFollowedByValidRow_ContinuesToProcessNextRow()
        {
            // The blank-subblock row must be skipped via "continue", not treated as the terminating
            // fully-blank row (which would "break" and never reach the row below it).
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Proposal for TTR";
            sheet.Cells[2, 1].Value = "Pattern";
            sheet.Cells[2, 2].Value = "Category";
            sheet.Cells[2, 3].Value = "SubBlock";
            sheet.Cells[3, 1].Value = "PAT1";
            sheet.Cells[3, 2].Value = "CAT1";
            sheet.Cells[3, 3].Value = "";
            sheet.Cells[4, 1].Value = "PAT2";
            sheet.Cells[4, 2].Value = "CAT2";
            sheet.Cells[4, 3].Value = "SUB2";

            var table = new EnableWordTable();
            table.ReadTable(sheet);

            Assert.AreEqual(1, table.Rows.Count);
            Assert.AreEqual("PAT2", table.Rows[0].Name);
        }

        [TestMethod]
        public void ReadTable_FullyBlankRow_StopsProcessingSubsequentRows()
        {
            // A row where every cell is genuinely blank must terminate the loop via "break"; guards
            // against a mutation to the blank-check that would instead fall through to the separate
            // blank-SubBlock "continue" and keep processing rows below it.
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Proposal for TTR";
            sheet.Cells[2, 1].Value = "Pattern";
            sheet.Cells[2, 2].Value = "Category";
            sheet.Cells[2, 3].Value = "SubBlock";
            sheet.Cells[3, 1].Value = "";
            sheet.Cells[3, 2].Value = "";
            sheet.Cells[3, 3].Value = "";
            sheet.Cells[4, 1].Value = "PAT2";
            sheet.Cells[4, 2].Value = "CAT2";
            sheet.Cells[4, 3].Value = "SUB2";

            var table = new EnableWordTable();
            table.ReadTable(sheet);

            Assert.AreEqual(0, table.Rows.Count);
        }

        [TestMethod]
        public void ReadTable_MultipleRows_OrderedByNameDescending()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Proposal for TTR";
            sheet.Cells[2, 1].Value = "Pattern";
            sheet.Cells[2, 2].Value = "Category";
            sheet.Cells[2, 3].Value = "SubBlock";
            sheet.Cells[3, 1].Value = "AAA";
            sheet.Cells[3, 2].Value = "CAT1";
            sheet.Cells[3, 3].Value = "SUB1";
            sheet.Cells[4, 1].Value = "ZZZ";
            sheet.Cells[4, 2].Value = "CAT2";
            sheet.Cells[4, 3].Value = "SUB2";

            var table = new EnableWordTable();
            table.ReadTable(sheet);

            Assert.AreEqual("ZZZ", table.Rows[0].Name);
            Assert.AreEqual("AAA", table.Rows[1].Name);
        }

        [TestMethod]
        public void ReadTable_ValidHeaderIsLastRowInSheet_StillFoundAndValid()
        {
            // The header-row scan must include sheet.Dimension.End.Row itself (<=, not <); guards
            // against an off-by-one that would miss a header row sitting exactly on the last row.
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Proposal for TTR";
            sheet.Cells[2, 1].Value = "Pattern";
            sheet.Cells[2, 2].Value = "Category";
            sheet.Cells[2, 3].Value = "SubBlock";

            var table = new EnableWordTable();
            table.ReadTable(sheet);

            Assert.IsTrue(table.IsValid);
        }

        [TestMethod]
        public void ReadTable_NoProposalCellButValidHeaderBelow_StillLeavesIsValidFalse()
        {
            // "Proposal for TTR" must be an actual substring match; guards against a mutation
            // to an empty pattern, which would match any cell text and false-positive on row 1.
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "SomeUnrelatedHeader";
            sheet.Cells[2, 1].Value = "Pattern";
            sheet.Cells[2, 2].Value = "Category";
            sheet.Cells[2, 3].Value = "SubBlock";
            sheet.Cells[3, 1].Value = "PAT1";
            sheet.Cells[3, 2].Value = "CAT1";
            sheet.Cells[3, 3].Value = "SUB1";

            var table = new EnableWordTable();
            table.ReadTable(sheet);

            Assert.IsFalse(table.IsValid);
        }

        [TestMethod]
        public void GetValidStartingRowIndex_ProposalCellOnLastRowOfSheet_StillFound()
        {
            // The scan must include sheet.Dimension.End.Row itself (<=, not <); guards against an
            // off-by-one that would miss the "Proposal for TTR" cell sitting exactly on the last row.
            // Tested directly since a valid header can never immediately follow the true last row.
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Proposal for TTR";

            int result = EnableWordTable.GetValidStartingRowIndex(sheet);

            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void GetBodyData_HeaderMetadataMissingPatternColumn_ThrowsInvalidOperationException()
        {
            // Guards Single() vs SingleOrDefault(): with zero matches, Single() throws
            // InvalidOperationException, while SingleOrDefault() would return null and defer the
            // failure to a NullReferenceException on the subsequent .Index access.
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 2].Value = "CAT1";
            sheet.Cells[1, 3].Value = "SUB1";
            var headerMetadata = new List<TtrColumnMetadata>
            {
                new("Category", 2),
                new("SubBlock", 3)
            };

            Assert.ThrowsException<InvalidOperationException>(() =>
                EnableWordTable.GetBodyData(sheet, 1, headerMetadata));
        }

        [TestMethod]
        public void GetBodyData_HeaderMetadataMissingSubBlockColumn_ThrowsInvalidOperationException()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "PAT1";
            sheet.Cells[1, 2].Value = "CAT1";
            var headerMetadata = new List<TtrColumnMetadata>
            {
                new("Pattern", 1),
                new("Category", 2)
            };

            Assert.ThrowsException<InvalidOperationException>(() =>
                EnableWordTable.GetBodyData(sheet, 1, headerMetadata));
        }

        #endregion

        #region EnableWordTableRow construction

        [TestMethod]
        public void Ctor_ExtractsNameCategorySubBlockFromDataset()
        {
            EnableWordTableRow row = MakeRow("PAT1");
            Assert.AreEqual("PAT1", row.Name);
            Assert.AreEqual("CAT", row.Category);
            Assert.AreEqual("SUB", row.SubBlock);
        }

        [TestMethod]
        public void Ctor_FailControlColumnAbsent_SetsFailControlEmpty()
        {
            EnableWordTableRow row = MakeRow("PAT1");
            Assert.AreEqual("", row.FailControl);
        }

        [TestMethod]
        public void Ctor_FailControlColumnPresent_SetsFailControl()
        {
            var headerMetadata = new Dictionary<string, TtrColumnMetadata>
            {
                ["Pattern"] = new TtrColumnMetadata("Pattern", 1),
                ["Category"] = new TtrColumnMetadata("Category", 2),
                ["SubBlock"] = new TtrColumnMetadata("SubBlock", 3),
                ["HIP_FailRun"] = new TtrColumnMetadata("HIP_FailRun", 4)
            };
            var dataset = new Dictionary<int, string> { [1] = "PAT1", [2] = "CAT", [3] = "SUB", [4] = "FailValue" };

            var row = new EnableWordTableRow("Pattern", "Category", "SubBlock", "HIP_FailRun", headerMetadata, dataset);

            Assert.AreEqual("FailValue", row.FailControl);
        }

        [TestMethod]
        public void IsOther_SubBlockEqualsOtherCaseInsensitive_ReturnsTrue()
        {
            EnableWordTableRow row = MakeRow("PAT1");
            row.SubBlock = "OTHER";
            Assert.IsTrue(row.IsOther);
        }

        [TestMethod]
        public void IsOther_SubBlockNotOther_ReturnsFalse()
        {
            EnableWordTableRow row = MakeRow("PAT1");
            row.SubBlock = "SUB1";
            Assert.IsFalse(row.IsOther);
        }

        [TestMethod]
        public void Ctor_PatternAndCategoryColumns_ExcludedFromTotalEnableWords()
        {
            EnableWordTableRow row = MakeRow("PAT1");
            CollectionAssert.DoesNotContain(row.TotalEnableWords, "Pattern");
            CollectionAssert.DoesNotContain(row.TotalEnableWords, "Category");
        }

        [TestMethod]
        public void Ctor_DataColumnWithCommaSeparatedValues_AddsEachAsEnableWordKey()
        {
            var headerMetadata = new Dictionary<string, TtrColumnMetadata>
            {
                ["Pattern"] = new TtrColumnMetadata("Pattern", 1),
                ["Category"] = new TtrColumnMetadata("Category", 2),
                ["SubBlock"] = new TtrColumnMetadata("SubBlock", 3),
                ["Test1"] = new TtrColumnMetadata("Test1", 4)
            };
            var dataset = new Dictionary<int, string> { [1] = "PAT1", [2] = "CAT", [3] = "SUB", [4] = "HV, LV" };

            var row = new EnableWordTableRow("Pattern", "Category", "SubBlock", "HIP_FailRun", headerMetadata, dataset);

            Assert.IsTrue(row.EnableWords.ContainsKey("HV"));
            Assert.IsTrue(row.EnableWords.ContainsKey("LV"));
            CollectionAssert.Contains(row.EnableWords["HV"], "Test1");
            CollectionAssert.Contains(row.EnableWords["LV"], "Test1");
        }

        [TestMethod]
        public void Ctor_HeaderNameWithSpace_IndexKeyIsUnderscored()
        {
            var headerMetadata = new Dictionary<string, TtrColumnMetadata>
            {
                ["Pattern"] = new TtrColumnMetadata("Pattern", 1),
                ["Category"] = new TtrColumnMetadata("Category", 2),
                ["SubBlock"] = new TtrColumnMetadata("SubBlock", 3),
                ["Test One"] = new TtrColumnMetadata("Test One", 4)
            };
            var dataset = new Dictionary<int, string> { [1] = "PAT1", [2] = "CAT", [3] = "SUB", [4] = "HV" };

            var row = new EnableWordTableRow("Pattern", "Category", "SubBlock", "HIP_FailRun", headerMetadata, dataset);

            CollectionAssert.Contains(row.TotalEnableWords, "Test_One");
            CollectionAssert.Contains(row.EnableWords["HV"], "Test_One");
        }

        [TestMethod]
        public void Ctor_SameEnableWordKeyFromMultipleColumns_AppendsIndexKeyToExistingList()
        {
            var headerMetadata = new Dictionary<string, TtrColumnMetadata>
            {
                ["Pattern"] = new TtrColumnMetadata("Pattern", 1),
                ["Category"] = new TtrColumnMetadata("Category", 2),
                ["SubBlock"] = new TtrColumnMetadata("SubBlock", 3),
                ["Test1"] = new TtrColumnMetadata("Test1", 4),
                ["Test2"] = new TtrColumnMetadata("Test2", 5)
            };
            var dataset = new Dictionary<int, string> { [1] = "PAT1", [2] = "CAT", [3] = "SUB", [4] = "HV", [5] = "HV" };

            var row = new EnableWordTableRow("Pattern", "Category", "SubBlock", "HIP_FailRun", headerMetadata, dataset);

            Assert.AreEqual(2, row.EnableWords["HV"].Count);
            CollectionAssert.Contains(row.EnableWords["HV"], "Test1");
            CollectionAssert.Contains(row.EnableWords["HV"], "Test2");
        }

        [TestMethod]
        public void Ctor_WhitespaceOnlyColumnValue_SkippedWithoutAddingEnableWordEntry()
        {
            // A value of a single space is non-empty before stripping (so the first guard doesn't
            // trigger) but becomes "" after Replace(" ", ""); guards against mutating that second check.
            var headerMetadata = new Dictionary<string, TtrColumnMetadata>
            {
                ["Pattern"] = new TtrColumnMetadata("Pattern", 1),
                ["Category"] = new TtrColumnMetadata("Category", 2),
                ["SubBlock"] = new TtrColumnMetadata("SubBlock", 3),
                ["Test1"] = new TtrColumnMetadata("Test1", 4)
            };
            var dataset = new Dictionary<int, string> { [1] = "PAT1", [2] = "CAT", [3] = "SUB", [4] = " " };

            var row = new EnableWordTableRow("Pattern", "Category", "SubBlock", "HIP_FailRun", headerMetadata, dataset);

            CollectionAssert.Contains(row.TotalEnableWords, "Test1");
            Assert.IsFalse(row.EnableWords.ContainsKey(""));
        }

        [TestMethod]
        public void Ctor_NonEmptyColumnValue_IsProcessedWithoutBeingTreatedAsEmpty()
        {
            // A non-empty, non-HLN-syntax value must still be parsed into EnableWords - guards against
            // a mutation of the empty-string literal compared against in the first guard, which (for a
            // value that happens to fail IsValidHlnSyntax) would incorrectly skip a real value.
            var headerMetadata = new Dictionary<string, TtrColumnMetadata>
            {
                ["Pattern"] = new TtrColumnMetadata("Pattern", 1),
                ["Category"] = new TtrColumnMetadata("Category", 2),
                ["SubBlock"] = new TtrColumnMetadata("SubBlock", 3),
                ["Test1"] = new TtrColumnMetadata("Test1", 4)
            };
            var dataset = new Dictionary<int, string> { [1] = "PAT1", [2] = "CAT", [3] = "SUB", [4] = "Stryker was here!" };

            var row = new EnableWordTableRow("Pattern", "Category", "SubBlock", "HIP_FailRun", headerMetadata, dataset);

            Assert.IsTrue(row.EnableWords.ContainsKey("Strykerwashere!"));
        }

        #endregion

        #region IsValidHlnSyntax

        [TestMethod]
        public void IsValidHlnSyntax_HlnToken_ReturnsTrue()
        {
            EnableWordTableRow row = MakeRow("PAT1");
            Assert.IsTrue(row.IsValidHlnSyntax("HV"));
        }

        [TestMethod]
        public void IsValidHlnSyntax_BvAndHbvTokens_ReturnTrue()
        {
            EnableWordTableRow row = MakeRow("PAT1");
            Assert.IsTrue(row.IsValidHlnSyntax("BV"));
            Assert.IsTrue(row.IsValidHlnSyntax("HBV"));
        }

        [TestMethod]
        public void IsValidHlnSyntax_PbCutToken_ReturnsTrue()
        {
            EnableWordTableRow row = MakeRow("PAT1");
            Assert.IsTrue(row.IsValidHlnSyntax("LV_PBCut"));
        }

        [TestMethod]
        public void IsValidHlnSyntax_ChkToken_ReturnsTrue()
        {
            EnableWordTableRow row = MakeRow("PAT1");
            Assert.IsTrue(row.IsValidHlnSyntax("HBV_Chk"));
        }

        [TestMethod]
        public void IsValidHlnSyntax_InvalidToken_ReturnsFalse()
        {
            EnableWordTableRow row = MakeRow("PAT1");
            Assert.IsFalse(row.IsValidHlnSyntax("NOT_A_VALID_TOKEN"));
        }

        [TestMethod]
        public void IsValidHlnSyntax_MultipleCommaSeparatedAllValid_ReturnsTrue()
        {
            EnableWordTableRow row = MakeRow("PAT1");
            Assert.IsTrue(row.IsValidHlnSyntax("HV,LV"));
        }

        [TestMethod]
        public void IsValidHlnSyntax_EmbeddedSpace_IsStrippedBeforeValidation()
        {
            // Guards Replace(" ", "") vs a mutation that inserts text instead of removing spaces;
            // a space-separated "HV, LV" must still validate as two clean tokens.
            EnableWordTableRow row = MakeRow("PAT1");
            Assert.IsTrue(row.IsValidHlnSyntax("HV, LV"));
        }

        [TestMethod]
        public void IsValidHlnSyntax_MultipleCommaSeparatedOneInvalid_ReturnsFalse()
        {
            EnableWordTableRow row = MakeRow("PAT1");
            Assert.IsFalse(row.IsValidHlnSyntax("HV,INVALID"));
        }

        #endregion

        #region GetBincutOverlay (private static, via reflection)

        [TestMethod]
        public void GetBincutOverlay_NoRowsNeedOverlay_ReturnsEmptyDictionary()
        {
            var rows = new List<EnableWordTableRow> { MakeRow("A_B_C_D_E_F_G_H_I_J", []) };

            Dictionary<string, List<string>> result = EnableWordTable.GetBincutOverlay(rows);

            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetBincutOverlay_NameHasNineOrFewerTokens_SkipsRow()
        {
            var rows = new List<EnableWordTableRow>
            {
                MakeRow("A_B_C", new Dictionary<string, List<string>> { ["LV_PBcut"] = ["X"] })
            };

            Dictionary<string, List<string>> result = EnableWordTable.GetBincutOverlay(rows);

            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetBincutOverlay_NameHasExactlyNineTokens_SkipsRowWithoutThrowing()
        {
            // Boundary case for "<= 9": a mutation to "< 9" would let a 9-token name (valid indices
            // 0-8) fall through to index [9], throwing IndexOutOfRangeException instead of skipping.
            var rows = new List<EnableWordTableRow>
            {
                MakeRow("A_B_C_D_E_F_G_H_I", new Dictionary<string, List<string>> { ["LV_PBcut"] = ["X"] })
            };

            Dictionary<string, List<string>> result = EnableWordTable.GetBincutOverlay(rows);

            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetBincutOverlay_NameHasTenTokens_ParsesModeFromTenthTokenStrippingX()
        {
            var rows = new List<EnableWordTableRow>
            {
                MakeRow("A_B_C_D_E_F_G_H_I_MODEX", new Dictionary<string, List<string>> { ["LV_PBcut"] = ["V1", "V2"] })
            };

            Dictionary<string, List<string>> result = EnableWordTable.GetBincutOverlay(rows);

            Assert.IsTrue(result.ContainsKey("MODE"));
            CollectionAssert.AreEquivalent(new List<string> { "V1", "V2" }, result["MODE"]);
        }

        [TestMethod]
        public void GetBincutOverlay_BothLvAndHvPresent_MergesAndDedupes()
        {
            var rows = new List<EnableWordTableRow>
            {
                MakeRow("A_B_C_D_E_F_G_H_I_MODE", new Dictionary<string, List<string>>
                {
                    ["LV_PBcut"] = ["V1", "V2"],
                    ["HV_PBCut"] = ["V2", "V3"]
                })
            };

            Dictionary<string, List<string>> result = EnableWordTable.GetBincutOverlay(rows);

            CollectionAssert.AreEquivalent(new List<string> { "V1", "V2", "V3" }, result["MODE"]);
        }

        [TestMethod]
        public void GetBincutOverlay_TwoRowsSameMode_MergesAcrossRows()
        {
            var rows = new List<EnableWordTableRow>
            {
                MakeRow("A_B_C_D_E_F_G_H_I_MODE", new Dictionary<string, List<string>> { ["LV_PBcut"] = ["V1"] }),
                MakeRow("Z_Y_X_W_V_U_T_S_R_MODE", new Dictionary<string, List<string>> { ["HV_PBCut"] = ["V2"] })
            };

            Dictionary<string, List<string>> result = EnableWordTable.GetBincutOverlay(rows);

            CollectionAssert.AreEquivalent(new List<string> { "V1", "V2" }, result["MODE"]);
        }

        #endregion

        #region GenJobAllEnableWord (private static, via reflection)

        [TestMethod]
        public void GenJobAllEnableWord_MatchingEnableWord_AddsAllSuffixedVariant()
        {
            var input = new List<string> { "ABC_cp" };

            List<string> result = EnableWordTable.GenJobAllEnableWord(input);

            CollectionAssert.Contains(result, "ABC_cp");
            CollectionAssert.Contains(result, "ABC_All");
        }

        [TestMethod]
        public void GenJobAllEnableWord_MultipleItemsSameEnable_AddsUniqueAllSuffixOnce()
        {
            var input = new List<string> { "ABC_cp", "ABC_ft" };

            List<string> result = EnableWordTable.GenJobAllEnableWord(input);

            int allSuffixCount = result.FindAll(x => x == "ABC_All").Count;
            Assert.AreEqual(1, allSuffixCount);
        }

        [TestMethod]
        public void GenJobAllEnableWord_NoMatchingItems_ReturnsUnchanged()
        {
            var input = new List<string> { "12345" };

            List<string> result = EnableWordTable.GenJobAllEnableWord(input);

            CollectionAssert.AreEqual(new List<string> { "12345" }, result);
        }

        #endregion

        #region GetHeaderMetadata (private static, via reflection)

        [TestMethod]
        public void GetHeaderMetadata_FiltersHvLvNvContainingHeaders()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Pattern";
            sheet.Cells[1, 2].Value = "LV_Test";
            sheet.Cells[1, 3].Value = "Category";
            sheet.Cells[1, 4].Value = "HV_X";
            sheet.Cells[1, 5].Value = "NV_Y";
            sheet.Cells[1, 6].Value = "Normal";

            IEnumerable<TtrColumnMetadata> result = EnableWordTable.GetHeaderMetadata(sheet, 1);
            var names = new List<string>();
            foreach (TtrColumnMetadata m in result)
            {
                names.Add(m.Name);
            }

            CollectionAssert.Contains(names, "Pattern");
            CollectionAssert.Contains(names, "Category");
            CollectionAssert.Contains(names, "Normal");
            CollectionAssert.DoesNotContain(names, "LV_Test");
            CollectionAssert.DoesNotContain(names, "HV_X");
            CollectionAssert.DoesNotContain(names, "NV_Y");
        }

        #endregion

        #region IsContainsSubBlock

        [TestMethod]
        public void IsContainsSubBlock_HeaderMetadataNeverSet_ReturnsFalse()
        {
            var table = new EnableWordTable();
            Assert.IsFalse(table.IsContainsSubBlock());
        }

        [TestMethod]
        public void IsContainsSubBlock_HeaderMetadataContainsSubBlock_ReturnsTrue()
        {
            var table = new EnableWordTable
            {
                _headerMetadata = [new("SubBlock", 1)]
            };

            Assert.IsTrue(table.IsContainsSubBlock());
        }

        [TestMethod]
        public void IsContainsSubBlock_HeaderMetadataWithoutSubBlock_ReturnsFalse()
        {
            var table = new EnableWordTable
            {
                _headerMetadata = [new("Pattern", 1)]
            };

            Assert.IsFalse(table.IsContainsSubBlock());
        }

        #endregion
    }
}
