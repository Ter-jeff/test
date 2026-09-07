using System.Collections.Generic;

using Automation.GenerateIgxl.EFuse.Business;
using Automation.GenerateIgxl.EFuse.InputChecker;

using CommonLib.ErrorReport;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Efuse.Input;

namespace Automation.Test.UT.Efuse
{
    [TestClass]
    public class EfuseConfigureBdfCheckerTests
    {
        private readonly EfuseConfigureBdfChecker _checker = new();

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
                row.RowData.Add(values.TryGetValue(i, out string? value) ? value : "");
            }

            return row;
        }

        private List<EfuseCrcItem> GetIgnorBitsList()
        {
            return _checker._ignorBitsList;
        }

        #region CheckPattenTypeWithBdf

        [TestMethod]
        public void CheckPattenTypeWithBdf_EmptyCurrentType_ReturnsFalse()
        {
            // Arrange
            var table = new BitDefTable { AccessMode = "JTAG-Access Mode data)" };

            // Act
            bool result = _checker.CheckPattenTypeWithBdf("", table);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void CheckPattenTypeWithBdf_JtgType_AccessModeMatches_ReturnsTrue()
        {
            // Arrange
            var table = new BitDefTable { AccessMode = "JTAG-Access Mode data)" };

            // Act
            bool result = _checker.CheckPattenTypeWithBdf("JTG", table);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void CheckPattenTypeWithBdf_JtgType_AccessModeDoesNotMatch_ReturnsFalse()
        {
            // Arrange
            var table = new BitDefTable { AccessMode = "Direct-Access Mode data)" };

            // Act
            bool result = _checker.CheckPattenTypeWithBdf("jtg", table);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void CheckPattenTypeWithBdf_DaaType_AccessModeMatches_ReturnsTrue()
        {
            // Arrange
            var table = new BitDefTable { AccessMode = "Direct-Access Mode data)" };

            // Act
            bool result = _checker.CheckPattenTypeWithBdf("DAA", table);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void CheckPattenTypeWithBdf_DaaType_AccessModeDoesNotMatch_ReturnsFalse()
        {
            // Arrange
            var table = new BitDefTable { AccessMode = "JTAG-Access Mode data)" };

            // Act
            bool result = _checker.CheckPattenTypeWithBdf("daa", table);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void CheckPattenTypeWithBdf_DaaType_MissingEndToken_ReturnsFalse()
        {
            // Arrange
            var table = new BitDefTable { AccessMode = "Direct-Access Mode" };

            // Act
            bool result = _checker.CheckPattenTypeWithBdf("DAA", table);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void CheckPattenTypeWithBdf_UnrecognizedType_ReturnsTrue()
        {
            // Arrange - neither JTG nor DAA -> falls through both branches untouched
            var table = new BitDefTable { AccessMode = "anything" };

            // Act
            bool result = _checker.CheckPattenTypeWithBdf("OTHER", table);

            // Assert
            Assert.IsTrue(result);
        }

        #endregion

        #region GetJobSequence

        [DataTestMethod]
        [DataRow("CP1", EfuseBitDefTableJobSequence.Cp1)]
        [DataRow("CP2", EfuseBitDefTableJobSequence.Cp2)]
        [DataRow("FT1", EfuseBitDefTableJobSequence.Ft1)]
        [DataRow("FT2", EfuseBitDefTableJobSequence.Ft2)]
        [DataRow("FT3", EfuseBitDefTableJobSequence.Ft3)]
        [DataRow("FTF", EfuseBitDefTableJobSequence.Ft3)]
        [DataRow("cp1_lowercase", EfuseBitDefTableJobSequence.Cp1)]
        [DataRow("SomethingElse", EfuseBitDefTableJobSequence.Unknown)]
        public void GetJobSequence_VariousJobStrings_ReturnsExpectedSequence(string job, EfuseBitDefTableJobSequence expected)
        {
            // Act
            EfuseBitDefTableJobSequence result = _checker.GetJobSequence(job);

            // Assert
            Assert.AreEqual(expected, result);
        }

        #endregion

        #region GetIgnoreBitRange

        [TestMethod]
        public void GetIgnoreBitRange_SingleRange_ReturnsAllBitsInRange()
        {
            // Act
            List<int> result = _checker.GetIgnoreBitRange("[3:0]");

            // Assert
            CollectionAssert.AreEquivalent(new List<int> { 0, 1, 2, 3 }, result);
        }

        [TestMethod]
        public void GetIgnoreBitRange_MultipleRangesCommaSeparated_ReturnsCombinedBits()
        {
            // Act
            List<int> result = _checker.GetIgnoreBitRange("3:0,7:5");

            // Assert
            CollectionAssert.AreEquivalent(new List<int> { 0, 1, 2, 3, 5, 6, 7 }, result);
        }

        [TestMethod]
        public void GetIgnoreBitRange_EmptyString_ReturnsEmptyList()
        {
            // Act
            List<int> result = _checker.GetIgnoreBitRange("");

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetIgnoreBitRange_LettersAndUnderscoresStripped_StillParsesCorrectly()
        {
            // Act - underscores and letters must be stripped before range parsing
            List<int> result = _checker.GetIgnoreBitRange("crc_calcbits=[7:4]");

            // Assert
            CollectionAssert.AreEquivalent(new List<int> { 4, 5, 6, 7 }, result);
        }

        #endregion

        #region CheckPatternMode

        [TestMethod]
        public void CheckPatternMode_NullPatternRows_NoErrorNoThrow()
        {
            // Act
            _checker.CheckPatternMode(null, []);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatternMode_EmptyPatternRows_NoError()
        {
            // Act
            _checker.CheckPatternMode([], [NewTable()]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatternMode_EmptyTables_NoError()
        {
            // Arrange
            var patternRows = new List<EfusePatternRow>
            {
                new() { BankName = "ECID", PayloadList = ["a_b_c_d_e_f_g_JTG"] }
            };

            // Act
            _checker.CheckPatternMode(patternRows, []);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatternMode_BankKeyEmpty_Skipped_NoError()
        {
            // Arrange
            var patternRows = new List<EfusePatternRow>
            {
                new() { BankName = "", PayloadList = ["a_b_c_d_e_f_g_JTG"] }
            };

            // Act
            _checker.CheckPatternMode(patternRows, [NewTable()]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatternMode_NoMatchingBdfTable_NoError()
        {
            // Arrange - bank name does not match any table's block name
            var patternRows = new List<EfusePatternRow>
            {
                new() { BankName = "NOMATCH", PayloadList = ["a_b_c_d_e_f_g_JTG"] }
            };

            // Act
            _checker.CheckPatternMode(patternRows, [NewTable("ECID")]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatternMode_MatchingBdfTableTypeMismatch_AddsError()
        {
            // Arrange - JTG pattern type but the table's access mode is Direct-Access
            var patternRows = new List<EfusePatternRow>
            {
                new() { BankName = "ECID", PayloadList = ["a_b_c_d_e_f_g_JTG"] }
            };
            BitDefTable table = NewTable("ECID");
            table.AccessMode = "Direct-Access Mode data)";

            // Act
            _checker.CheckPatternMode(patternRows, [table]);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckPatternMode_MatchingBdfTableTypeMatches_NoError()
        {
            // Arrange
            var patternRows = new List<EfusePatternRow>
            {
                new() { BankName = "ECID", PayloadList = ["a_b_c_d_e_f_g_JTG"] }
            };
            BitDefTable table = NewTable("ECID");
            table.AccessMode = "JTAG-Access Mode data)";

            // Act
            _checker.CheckPatternMode(patternRows, [table]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        #endregion

        #region CheckDefaultValue

        [TestMethod]
        public void CheckDefaultValue_CondAlgorithm_SkipsCheck_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("CFG");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [9] = "cond", [0] = "bank_config_Foo" });
            table.Rows.Add(row);
            var cfgSheet = new EfuseConfigMainSheet("Config");

            // Act
            _checker.CheckDefaultValue(cfgSheet, table);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckDefaultValue_NonCondAlgorithm_RegisterFound_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("CFG");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [9] = "base", [0] = "bank_config_Foo" });
            table.Rows.Add(row);
            var cfgSheet = new EfuseConfigMainSheet("Config");
            cfgSheet.Rows.Add(new EfuseConfigMainRow { Description = "Foo" });

            // Act
            _checker.CheckDefaultValue(cfgSheet, table);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckDefaultValue_NonCondAlgorithm_RegisterNotFound_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("CFG");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [9] = "base", [0] = "bank_config_Foo" });
            table.Rows.Add(row);
            var cfgSheet = new EfuseConfigMainSheet("Config");

            // Act
            _checker.CheckDefaultValue(cfgSheet, table);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        #endregion

        #region CheckCrcIgnorBit

        [TestMethod]
        public void CheckCrcIgnorBit_NoCrcItems_ResultsInEmptyList()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [9] = "base", [2] = "0" }));

            // Act
            _checker.CheckCrcIgnorBit(table);

            // Assert
            Assert.AreEqual(0, GetIgnorBitsList().Count);
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckCrcIgnorBit_OneComplementTargetKeyword_ExcludedFromProcessing()
        {
            // Arrange - the exclusion filter matches against the raw `Line` field, not RowData;
            // deliberately leave Msb (RowData[1]) unset to prove the row is never processed
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [9] = "crc", [2] = "0", [10] = "one_complement_target[3:0]" });
            row.Line = "some_prefix one_complement_target[3:0] some_suffix";
            table.Rows.Add(row);

            // Act
            _checker.CheckCrcIgnorBit(table);

            // Assert
            Assert.AreEqual(0, GetIgnorBitsList().Count);
        }

        [TestMethod]
        public void CheckCrcIgnorBit_CalcBitsSelfRangeOverlapsOwnIgnoreRange_AddsError()
        {
            // Arrange - for "crc_calcbits" items, an error fires when the item's OWN bit range
            // (after offset) is found WITHIN its own declared ignore-bit set (a self-reference bug).
            // Own range [0:1] here matches the ignore range [1:0] exactly.
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [9] = "base", [2] = "0" }));
            table.Rows.Add(NewRow(new SortedDictionary<int, string>
            {
                [9] = "crc",
                [1] = "1",
                [2] = "0",
                [10] = "crc_calcbits=[1:0]",
                [0] = "Item1"
            }));

            // Act
            _checker.CheckCrcIgnorBit(table);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
            Assert.AreEqual(1, GetIgnorBitsList().Count);
            Assert.IsTrue(GetIgnorBitsList()[0].IsCalcBits);
        }

        [TestMethod]
        public void CheckCrcIgnorBit_CalcBitsSelfRangeDoesNotOverlapOwnIgnoreRange_NoError()
        {
            // Arrange - own range [0:1] does NOT overlap the declared ignore range [3:2]
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [9] = "base", [2] = "0" }));
            table.Rows.Add(NewRow(new SortedDictionary<int, string>
            {
                [9] = "crc",
                [1] = "1",
                [2] = "0",
                [10] = "crc_calcbits=[3:2]",
                [0] = "Item1"
            }));

            // Act
            _checker.CheckCrcIgnorBit(table);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckCrcIgnorBit_NonCalcBitsOutsideIgnoreRange_AddsError()
        {
            // Arrange - non-calcbits description; item's own bit range [1:0] must exist
            // WITHIN the ignore range, otherwise flagged
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [9] = "base", [2] = "0" }));
            table.Rows.Add(NewRow(new SortedDictionary<int, string>
            {
                [9] = "crc",
                [1] = "1",
                [2] = "0",
                [10] = "[3:2]",
                [0] = "Item1"
            }));

            // Act
            _checker.CheckCrcIgnorBit(table);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
            Assert.IsFalse(GetIgnorBitsList()[0].IsCalcBits);
        }

        [TestMethod]
        public void CheckCrcIgnorBit_NonCalcBitsWithinIgnoreRange_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [9] = "base", [2] = "0" }));
            table.Rows.Add(NewRow(new SortedDictionary<int, string>
            {
                [9] = "crc",
                [1] = "1",
                [2] = "0",
                [10] = "[1:0]",
                [0] = "Item1"
            }));

            // Act
            _checker.CheckCrcIgnorBit(table);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckCrcIgnorBit_NoIgnoreBitsInDescription_ContinuesWithoutAddingItem()
        {
            // Arrange - description yields an empty ignore range -> "continue" skips the item entirely
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [9] = "base", [2] = "0" }));
            table.Rows.Add(NewRow(new SortedDictionary<int, string>
            {
                [9] = "crc",
                [1] = "1",
                [2] = "0",
                [10] = "",
                [0] = "Item1"
            }));

            // Act
            _checker.CheckCrcIgnorBit(table);

            // Assert
            Assert.AreEqual(0, GetIgnorBitsList().Count);
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        #endregion

        #region CheckCfgJob / CheckJob

        [TestMethod]
        public void CheckCfgJob_CurrentJobSeqHigherAndCalcBitInIgnoreRange_AddsError()
        {
            // Arrange - populate _ignorBitsList with a CP1-job calc-bits ignore item covering bit 0
            BitDefTable crcTable = NewTable("ECID");
            crcTable.Rows.Add(NewRow(new SortedDictionary<int, string> { [9] = "base", [2] = "0" }));
            crcTable.Rows.Add(NewRow(new SortedDictionary<int, string>
            {
                [9] = "crc",
                [1] = "1",
                [2] = "0",
                [10] = "crc_calcbits=[0:0]",
                [4] = "CP1",
                [0] = "Item1"
            }));
            _checker.CheckCrcIgnorBit(crcTable);
            ErrorReportManager.ClearErrors();

            var cfgTable = new EfuseConfigMainSheet("Config") { FuseBlowLocationColNumber = 3 };
            var row = new EfuseConfigMainRow { Fuse = "[0:0]", FuseBlowLocation = "FT1" };

            // Act - FT1 job sequence is higher than the ignore item's CP1 job
            _checker.CheckCfgJob(cfgTable, row);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckCfgJob_CurrentJobSeqNotHigher_NoError()
        {
            // Arrange
            BitDefTable crcTable = NewTable("ECID");
            crcTable.Rows.Add(NewRow(new SortedDictionary<int, string> { [9] = "base", [2] = "0" }));
            crcTable.Rows.Add(NewRow(new SortedDictionary<int, string>
            {
                [9] = "crc",
                [1] = "1",
                [2] = "0",
                [10] = "crc_calcbits=[0:0]",
                [4] = "FT3",
                [0] = "Item1"
            }));
            _checker.CheckCrcIgnorBit(crcTable);
            ErrorReportManager.ClearErrors();

            var cfgTable = new EfuseConfigMainSheet("Config") { FuseBlowLocationColNumber = 3 };
            var row = new EfuseConfigMainRow { Fuse = "[0:0]", FuseBlowLocation = "CP1" };

            // Act - CP1 job sequence is not higher than the ignore item's FT3 job
            _checker.CheckCfgJob(cfgTable, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckJob_InvalidLsb_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [4] = "CP1", [2] = "notANumber", [1] = "3" });

            // Act
            _checker.CheckJob(table, row);

            // Assert
            Assert.IsTrue(ErrorReportManager.GetErrorCount() >= 1);
        }

        [TestMethod]
        public void CheckJob_InvalidMsb_AddsError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [4] = "CP1", [2] = "0", [1] = "notANumber" });

            // Act
            _checker.CheckJob(table, row);

            // Assert
            Assert.IsTrue(ErrorReportManager.GetErrorCount() >= 1);
        }

        [TestMethod]
        public void CheckJob_ValidLsbMsbNoIgnoreItems_NoError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [4] = "CP1", [2] = "0", [1] = "3" });

            // Act - _ignorBitsList is empty by default, so the inner loop never executes
            _checker.CheckJob(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        #endregion

        #region WorkFlow / CheckCrcItems integration

        [TestMethod]
        public void WorkFlow_MinimalValidData_NoThrow()
        {
            // Arrange
            BitDefTable cfgTable = NewTable("CFG");
            cfgTable.Rows.Add(NewRow(new SortedDictionary<int, string> { [9] = "base", [2] = "0", [0] = "bank_config_Foo" }));
            var cfgSheet = new EfuseConfigMainSheet("Config") { FuseBlowLocationColNumber = 3 };
            cfgSheet.Rows.Add(new EfuseConfigMainRow { Description = "Foo", Lsb = 0, Msb = 3, Fuse = "[3:0]", FuseBlowLocation = "CP1" });

            // Act
            _checker.WorkFlow(cfgSheet, [cfgTable]);

            // Assert - should complete without throwing; no assertions on error count since
            // this exercises the full orchestration path across multiple checks
        }

        [TestMethod]
        public void CheckCrcItems_NonCondAlgorithmRow_DelegatesToCheckJob()
        {
            // Arrange - non-cond algorithm rows should be routed to CheckJob, not CheckCfgJob;
            // an invalid Lsb on the target row should surface as an E_InvalidLsbMsb error.
            // A valid first row is required because CheckCrcIgnorBit unconditionally parses
            // table.Rows.First()'s Lsb value to compute an offset, regardless of algorithm.
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [9] = "base", [2] = "0", [1] = "1", [4] = "CP1" }));
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [9] = "base", [2] = "notANumber", [1] = "3", [4] = "CP1" }));
            var cfgSheet = new EfuseConfigMainSheet("Config");

            // Act
            _checker.CheckCrcItems(cfgSheet, [table]);

            // Assert
            Assert.IsTrue(ErrorReportManager.GetErrorCount() >= 1);
        }

        #endregion
    }
}
