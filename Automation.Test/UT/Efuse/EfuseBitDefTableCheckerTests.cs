using System.Collections.Generic;

using Automation.GenerateIgxl.EFuse.InputChecker;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Efuse.Input;

namespace Automation.Test.UT.Efuse
{
    [TestClass]
    public class EfuseBitDefTableCheckerTests
    {
        private readonly EfuseBitDefTableChecker _checker = new();

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

        private static BitDefTable NewTable(string blockName = "ECID")
        {
            return new BitDefTable
            {
                SheetName = "Sheet1",
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

        private static BitDefRow NewRow(SortedDictionary<int, string> values)
        {
            var row = new BitDefRow { SheetName = "Sheet1", RowNum = 2 };
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
                row.RowData.Add(values.TryGetValue(i, out string value) ? value : "");
            }

            return row;
        }

        [TestMethod]
        public void CheckBankType_UnknownBlockName_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("NotABank");

            // Act
            _checker.CheckBankType(table);

            // Assert
            Assert.IsTrue(ErrorReportManager.Contains(EFuseErrorType.W_RuleViolationBank_01.FormatMessage("NotABank")));
        }

        [TestMethod]
        public void CheckBankType_KnownBlockName_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");

            // Act
            _checker.CheckBankType(table);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckNonCpStageInEcid_EcidBankWithNonCpStage_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [4] = "FTF" });

            // Act
            _checker.CheckNonCpStageInEcid(table, row);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckNonCpStageInEcid_EcidBankWithCpStage_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [4] = "CP1" });

            // Act
            _checker.CheckNonCpStageInEcid(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckNonCpStageInEcid_NonEcidBank_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("CFG");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [4] = "FTF" });

            // Act
            _checker.CheckNonCpStageInEcid(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckIfProgramStageIsFtf_StageIsFtf_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [4] = "FTF" });

            // Act
            _checker.CheckIfProgramStageIsFtf(table, row);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckIfProgramStageIsFtf_StageIsNotFtf_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [4] = "CP1" });

            // Act
            _checker.CheckIfProgramStageIsFtf(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckBitFormat_AllBitsNumberAndConsistent_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [1] = "7", [2] = "0", [3] = "8" });

            // Act
            InvokeCheckBitFormat(table, row, out string width);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
            Assert.AreEqual("8", width);
        }

        [TestMethod]
        public void CheckBitFormat_WidthMismatch_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [1] = "7", [2] = "0", [3] = "5" });

            // Act
            InvokeCheckBitFormat(table, row, out _);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckBitFormat_MsbNotNumber_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [1] = "X", [2] = "0", [3] = "8" });

            // Act
            InvokeCheckBitFormat(table, row, out _);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckBitFormat_WidthNotNumber_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [1] = "7", [2] = "0", [3] = "X" });

            // Act
            InvokeCheckBitFormat(table, row, out _);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckBitFormat_WidthOver1024_AddsWarning()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [1] = "2000", [2] = "0", [3] = "2001" });

            // Act
            InvokeCheckBitFormat(table, row, out _);

            // Assert
            Assert.IsTrue(ErrorReportManager.Contains(EFuseErrorType.W_InvalidBit_01.FormatMessage()));
        }

        private void InvokeCheckBitFormat(BitDefTable table, BitDefRow row, out string width)
        {
            width = "";
            _checker.CheckBitFormat(table, row, ref width);
        }

        [TestMethod]
        public void CheckCrcFormat_ValidBitRangeInDescription_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [9] = "crc",
                [10] = "[3:0]",
                [8] = ""
            });

            // Act
            _checker.CheckCrcFormat(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckCrcFormat_InvalidBitRangeInDescription_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [9] = "crc",
                [10] = "[bad]",
                [8] = ""
            });

            // Act
            _checker.CheckCrcFormat(table, row);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckCrcFormat_OneComplementTargetKeyword_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [9] = "crc",
                [10] = "one_complement_target[bad]",
                [8] = ""
            });

            // Act
            _checker.CheckCrcFormat(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckCrcFormat_InvalidBitRangeInDefaultValue_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [9] = "crc",
                [10] = "",
                [8] = "[bad]"
            });

            // Act
            _checker.CheckCrcFormat(table, row);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckCrcFormat_NotCrcAlgorithm_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [9] = "base",
                [10] = "[bad]"
            });

            // Act
            _checker.CheckCrcFormat(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckDefualtValue_HexDefaultValueTooLongForBitWidth_AddsError()
        {
            // Arrange - bit width 4 -> bitSize = ceil(4/4) = 1 hex digit allowed, "FF" has 2
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [3] = "4", [8] = "0xFF" });

            // Act
            _checker.CheckDefualtValue(table, row);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckDefualtValue_HexDefaultValueFitsBitWidth_NoError()
        {
            // Arrange - bit width 8 -> bitSize = ceil(8/4) = 2 hex digits allowed
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [3] = "8", [8] = "0xFF" });

            // Act
            _checker.CheckDefualtValue(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckDefualtValue_BinaryDefaultValueTooLongForBitWidth_AddsError()
        {
            // Arrange - bit width 2 but binary value has 4 digits
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [3] = "2", [8] = "0b1111" });

            // Act
            _checker.CheckDefualtValue(table, row);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckDefualtValue_BinaryDefaultValueFitsBitWidth_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [3] = "4", [8] = "0b1111" });

            // Act
            _checker.CheckDefualtValue(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckDefualtValue_BitWidthNotNumber_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [3] = "notANumber", [8] = "0xFF" });

            // Act
            _checker.CheckDefualtValue(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckCmpBitWidthInUdr_MismatchedBitWidthWithUdr_AddsError()
        {
            // Arrange
            BitDefTable cmpTable = NewTable("CMP_E");
            BitDefRow cmpRow = NewRow(new SortedDictionary<int, string> { [0] = "Item1", [3] = "8" });

            BitDefTable udrTable = NewTable("UDR_E");
            BitDefRow udrRow = NewRow(new SortedDictionary<int, string> { [0] = "Item1", [3] = "4" });
            udrTable.Rows.Add(udrRow);

            var udrItems = new List<BitDefTable> { udrTable };

            // Act
            _checker.CheckCmpBitWidthInUdr(cmpTable, cmpRow, udrItems);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckCmpBitWidthInUdr_MatchingBitWidthWithUdr_NoError()
        {
            // Arrange
            BitDefTable cmpTable = NewTable("CMP_E");
            BitDefRow cmpRow = NewRow(new SortedDictionary<int, string> { [0] = "Item1", [3] = "8" });

            BitDefTable udrTable = NewTable("UDR_E");
            BitDefRow udrRow = NewRow(new SortedDictionary<int, string> { [0] = "Item1", [3] = "8" });
            udrTable.Rows.Add(udrRow);

            var udrItems = new List<BitDefTable> { udrTable };

            // Act
            _checker.CheckCmpBitWidthInUdr(cmpTable, cmpRow, udrItems);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void WorkFlow_EcidTableWithFtfProgrammingStage_AddsErrors()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [0] = "Item1",
                [1] = "7",
                [2] = "0",
                [3] = "8",
                [4] = "FTF",
                [5] = "Default",
                [6] = "",
                [7] = "",
                [8] = "1",
                [9] = "notbase",
                [10] = "",
                [11] = ""
            });
            table.Rows.Add(row);

            // Act
            _checker.WorkFlow([table]);

            // Assert - both CheckNonCpStageInEcid and CheckIfProgramStageIsFtf should fire
            Assert.IsTrue(ErrorReportManager.GetErrorCount() >= 2);
        }

        [TestMethod]
        public void WorkFlow_EmptyTableList_NoErrorAndNoThrow()
        {
            // Act
            _checker.WorkFlow([]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckShadowFieldStage_ShadowAndNonShadowSameStage_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow baseRow = NewRow(new SortedDictionary<int, string> { [0] = "Field1", [4] = "CP1" });
            BitDefRow shadowRow = NewRow(new SortedDictionary<int, string> { [0] = "Field1_shadow", [4] = "CP1" });
            table.Rows.Add(baseRow);
            table.Rows.Add(shadowRow);

            // Act
            _checker.CheckShadowFieldStage([table]);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckShadowFieldStage_ShadowAndNonShadowDifferentStage_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow baseRow = NewRow(new SortedDictionary<int, string> { [0] = "Field1", [4] = "CP1" });
            BitDefRow shadowRow = NewRow(new SortedDictionary<int, string> { [0] = "Field1_shadow", [4] = "CP2" });
            table.Rows.Add(baseRow);
            table.Rows.Add(shadowRow);

            // Act
            _checker.CheckShadowFieldStage([table]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckShadowFieldStage_NoShadowFields_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow baseRow = NewRow(new SortedDictionary<int, string> { [0] = "Field1", [4] = "CP1" });
            table.Rows.Add(baseRow);

            // Act
            _checker.CheckShadowFieldStage([table]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckShadowFieldStage_EmptyTableList_NoError()
        {
            // Act
            _checker.CheckShadowFieldStage([]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckLimitFormat_RealWithNonNumericLowLimit_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [5] = "real",
                [6] = "notANumber",
                [7] = "10",
                [8] = "",
                [9] = "notcrc"
            });

            // Act
            _checker.CheckLimitFormat(table, row, "8");

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckLimitFormat_RealWithNonNumericHighLimit_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [5] = "real",
                [6] = "0",
                [7] = "notANumber",
                [8] = "",
                [9] = "notcrc"
            });

            // Act
            _checker.CheckLimitFormat(table, row, "8");

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckLimitFormat_CfgBank_SkipsRealLimitCheckEvenWhenNonNumeric()
        {
            // Arrange - CFG bank is explicitly excluded from the real-limit check
            BitDefTable table = NewTable("CFG");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [5] = "real",
                [6] = "notANumber",
                [7] = "notANumber",
                [8] = "",
                [9] = "notcrc"
            });

            // Act
            _checker.CheckLimitFormat(table, row, "8");

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckLimitFormat_DefaultWithNonNumericAndNonHexValue_AddsWarning()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [5] = "Default",
                [6] = "",
                [7] = "",
                [8] = "!!!notHexOrNumber",
                [9] = "notcrc"
            });

            // Act
            _checker.CheckLimitFormat(table, row, "8");

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckLimitFormat_DefaultWithNumericValue_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [5] = "Default",
                [6] = "",
                [7] = "",
                [8] = "5",
                [9] = "notcrc"
            });

            // Act
            _checker.CheckLimitFormat(table, row, "8");

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckLimitFormat_RealNonCrcHighLessThanLow_CascadesToLimitConsistencyError()
        {
            // Arrange - all values parse as numbers, so CheckLimitFormat delegates to
            // CheckLimitConsistency, which flags High Limit <= Low Limit.
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [5] = "real",
                [6] = "10",
                [7] = "5",
                [8] = "",
                [9] = "notcrc"
            });

            // Act
            _checker.CheckLimitFormat(table, row, "8");

            // Assert
            Assert.IsTrue(ErrorReportManager.GetErrorCount() >= 1);
        }

        [TestMethod]
        public void CheckLimitConsistency_RealNonCrcHighLessThanOrEqualLow_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [5] = "real", [9] = "notcrc" });

            // Act
            _checker.CheckLimitConsistency(table, row, "8", "10", "5", "");

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckLimitConsistency_RealNonCrcHighGreaterThanLow_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [5] = "real", [9] = "notcrc" });

            // Act
            _checker.CheckLimitConsistency(table, row, "8", "5", "10", "");

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckLimitConsistency_CrcAlgorithm_SkipsHighLowComparison()
        {
            // Arrange - algorithm is crc, so the real/non-crc comparison is skipped
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [5] = "real", [9] = "crc" });

            // Act
            _checker.CheckLimitConsistency(table, row, "8", "10", "5", "");

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckLimitConsistency_IdsAlgorithmResolutionExceedsWidth_AddsError()
        {
            // Arrange - resolution * (2^width - 1) + resolution < highLimit
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [9] = "ids", [11] = "1" });

            // Act - width=2 -> max representable = 1*(2^2-1)+1 = 4, highLimit=100 exceeds it
            _checker.CheckLimitConsistency(table, row, "2", "0", "100", "");

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckLimitConsistency_IdsAlgorithmResolutionWithinWidth_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [9] = "ids", [11] = "1" });

            // Act - width=8 -> max representable = 1*(2^8-1)+1 = 256, highLimit=100 fits
            _checker.CheckLimitConsistency(table, row, "8", "0", "100", "");

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckBaseAlgorithmLimit_NotBaseAlgorithm_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [9] = "notbase" });

            // Act
            _checker.CheckBaseAlgorithmLimit(table, row, "5", "5", "5");

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckBaseAlgorithmLimit_ValuesDiffer_AddsRuleViolationWarning()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [9] = "base" });

            // Act
            _checker.CheckBaseAlgorithmLimit(table, row, "1", "2", "3");

            // Assert
            Assert.IsTrue(ErrorReportManager.Contains(EFuseErrorType.W_RuleViolationLimit_01.FormatMessage()));
        }

        [TestMethod]
        public void CheckBaseAlgorithmLimit_FirstBaseRow_AllEqual_NoError()
        {
            // Arrange - the first base row establishes _baseVoltage; matching values produce no error
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [9] = "base" });

            // Act
            _checker.CheckBaseAlgorithmLimit(table, row, "3", "3", "3");

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckBaseAlgorithmLimit_SubsequentBaseRowDiffersFromBaseVoltage_AddsMismatchWarning()
        {
            // Arrange - first call establishes _baseVoltage = 3; second call with equal-but-different
            // values (all equal to 5, not 3) should flag a mismatch against the established base voltage.
            BitDefTable table = NewTable("ECID");
            BitDefRow firstRow = NewRow(new SortedDictionary<int, string> { [9] = "base" });
            BitDefRow secondRow = NewRow(new SortedDictionary<int, string> { [9] = "base" });
            _checker.CheckBaseAlgorithmLimit(table, firstRow, "3", "3", "3");
            ErrorReportManager.ClearErrors();

            // Act
            _checker.CheckBaseAlgorithmLimit(table, secondRow, "5", "5", "5");

            // Assert
            Assert.IsTrue(ErrorReportManager.Contains(EFuseErrorType.W_MismatchRow_01.FormatMessage()));
        }
    }
}
