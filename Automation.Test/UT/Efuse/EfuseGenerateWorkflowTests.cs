using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.EFuse.Business;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Efuse.Input;

namespace Automation.Test.UT.Efuse
{
    [TestClass]
    public class EfuseGenerateWorkflowTests
    {
        private EFuseInputData _inputData = null!;
        private EfuseGenerate _target = null!;

        [TestInitialize]
        public void Setup()
        {
            _inputData = new EFuseInputData
            {
                EfuseArraySizeSheet = new EfuseArraySizeSheet("EfuseArraySize")
            };
            _target = new EfuseGenerate(_inputData);
            ErrorReportManager.ClearErrors();
        }

        [TestCleanup]
        public void Cleanup()
        {
            ErrorReportManager.ClearErrors();
            LocalSpecs.Options.Device = EnumDevice.AP;
        }

        private static BitDefTable NewTable(string blockName, int maxColIdx = 16)
        {
            return new BitDefTable
            {
                SheetName = "Sheet1",
                BlockName = blockName,
                MaxColIdx = maxColIdx,
                BankEfuseBitDefIdx = 0,
                MsbBitIdx = 1,
                LsbBitIdx = 2,
                BitWidthIdx = 3,
                Na1Idx = 4,
                Na2Idx = 5,
                Na3Idx = 6,
                Na4Idx = 7,
                ProgrammingStageIdx = 8,
                LowLimitIdx = 9,
                HighLimitIdx = 10,
                ResolutionIdx = 11,
                AlgorithmIdx = 12,
                DescriptionIdx = 13,
                DefaultOrRealIdx = 14,
                DefaultValueIdx = 15
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

        [TestMethod]
        public void FilterEfuseConfigSheets_NullSource_ReturnsEmptyList()
        {
            // Act
            List<EfuseConfigMainSheet> result = EfuseGenerate.FilterEfuseConfigSheets(null, "Cfg", true);

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void FilterEfuseConfigSheets_NoMultipleBlock_ReturnsAllSheets()
        {
            // Arrange
            var sheets = new List<EfuseConfigMainSheet>
            {
                new("Config_A"),
                new("Config_B")
            };

            // Act
            List<EfuseConfigMainSheet> result = EfuseGenerate.FilterEfuseConfigSheets(sheets, "A", false);

            // Assert
            Assert.AreEqual(2, result.Count);
        }

        [TestMethod]
        public void FilterEfuseConfigSheets_EmptyConfigType_ReturnsAllSheets()
        {
            // Arrange
            var sheets = new List<EfuseConfigMainSheet> { new("Config_A") };

            // Act
            List<EfuseConfigMainSheet> result = EfuseGenerate.FilterEfuseConfigSheets(sheets, "", true);

            // Assert
            Assert.AreEqual(1, result.Count);
        }

        [TestMethod]
        public void FilterEfuseConfigSheets_MultipleBlockWithConfigType_FiltersByName()
        {
            // Arrange
            var sheets = new List<EfuseConfigMainSheet>
            {
                new("Config_A"),
                new("Config_B")
            };

            // Act
            List<EfuseConfigMainSheet> result = EfuseGenerate.FilterEfuseConfigSheets(sheets, "A", true);

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("Config_A", result[0].SheetName);
        }

        [TestMethod]
        public void ConvertDecimalToBinary_Zero_ReturnsEmptyString()
        {
            // Act
            string result = EfuseGenerate.ConvertDecimalToBinary(0);

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void ConvertDecimalToBinary_Five_ReturnsBinaryRepresentation()
        {
            // Act
            string result = EfuseGenerate.ConvertDecimalToBinary(5);

            // Assert
            Assert.AreEqual("101", result);
        }

        [TestMethod]
        public void ConvertDecimalToBinary_Two_ReturnsBinaryRepresentation()
        {
            // Act
            string result = EfuseGenerate.ConvertDecimalToBinary(2);

            // Assert
            Assert.AreEqual("10", result);
        }

        [TestMethod]
        public void CheckIfEarlyConfig_CfgBankWithEarlyStage_ReturnsTrue()
        {
            // Arrange
            BitDefTable table = NewTable("CFG");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [8] = "Early" }));
            _target.FinalBitDefTables = [table];

            // Act
            bool result = _target.CheckIfEarlyConfig();

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void CheckIfEarlyConfig_CfgBankWithoutEarlyStage_ReturnsFalse()
        {
            // Arrange
            BitDefTable table = NewTable("CFG");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [8] = "CP1" }));
            _target.FinalBitDefTables = [table];

            // Act
            bool result = _target.CheckIfEarlyConfig();

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void CheckIfEarlyConfig_NonCfgBank_ReturnsFalse()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [8] = "Early" }));
            _target.FinalBitDefTables = [table];

            // Act
            bool result = _target.CheckIfEarlyConfig();

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void CheckIfSvmKeyWord_AnySheetHasSvm_ReturnsTrue()
        {
            // Arrange
            _inputData.EfuseConfigMainSheets = [new EfuseConfigMainSheet("Config") { HasSvm = true }];

            // Act
            bool result = _target.CheckIfSvmKeyWord();

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void CheckIfSvmKeyWord_NoSheetHasSvm_ReturnsFalse()
        {
            // Arrange
            _inputData.EfuseConfigMainSheets = [new EfuseConfigMainSheet("Config") { HasSvm = false }];

            // Act
            bool result = _target.CheckIfSvmKeyWord();

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void CheckIfHasBkmProcess_RowStartsWithBkmProcess_SetsLocalSpecsTrue()
        {
            // Arrange
            BitDefTable table = NewTable("CFG");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [0] = "bkm_process" }));
            _target.FinalBitDefTables = [table];

            try
            {
                // Act
                _target.CheckIfHasBkmProcess();

                // Assert
                Assert.IsTrue(LocalSpecs.HasBkmProcess);
            }
            finally
            {
                LocalSpecs.HasBkmProcess = false;
            }
        }

        [TestMethod]
        public void CheckIfHasBkmProcess_NoMatchingRow_SetsLocalSpecsFalse()
        {
            // Arrange
            BitDefTable table = NewTable("CFG");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [0] = "other_field" }));
            _target.FinalBitDefTables = [table];
            LocalSpecs.HasBkmProcess = true;

            try
            {
                // Act
                _target.CheckIfHasBkmProcess();

                // Assert
                Assert.IsFalse(LocalSpecs.HasBkmProcess);
            }
            finally
            {
                LocalSpecs.HasBkmProcess = false;
            }
        }

        [TestMethod]
        public void GetConfigRowInBitDefTables_CondAlgorithmRow_IncludedInConfigTable()
        {
            // Arrange
            BitDefTable table = NewTable("Config_ABC");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [1] = "7", [2] = "0", [12] = "cond", [14] = "", [15] = "" }));
            _target.FinalBitDefTables = [table];

            // Act
            List<BitDefTable> result = _target.GetConfigRowInBitDefTables();

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(1, result[0].Rows.Count);
            Assert.AreEqual("ABC", result[0].BlockName);
        }

        [TestMethod]
        public void GetConfigRowInBitDefTables_DefaultNaRow_IncludedInConfigTable()
        {
            // Arrange
            BitDefTable table = NewTable("Config_XYZ");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [1] = "7", [2] = "0", [12] = "app", [14] = "default", [15] = "NA" }));
            _target.FinalBitDefTables = [table];

            // Act
            List<BitDefTable> result = _target.GetConfigRowInBitDefTables();

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(1, result[0].Rows.Count);
        }

        [TestMethod]
        public void GetConfigRowInBitDefTables_NoMatchingRows_ReturnsEmptyList()
        {
            // Arrange
            BitDefTable table = NewTable("Config_XYZ");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [1] = "7", [2] = "0", [12] = "app", [14] = "real", [15] = "5" }));
            _target.FinalBitDefTables = [table];

            // Act
            List<BitDefTable> result = _target.GetConfigRowInBitDefTables();

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void IsInBitDefConfigBitRange_RowWithinRange_ReturnsTrue()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [1] = "10", [2] = "0" }));
            var configTables = new List<BitDefTable> { table };
            var configRow = new EfuseConfigMainRow { Lsb = 2, Msb = 5 };

            // Act
            bool result = _target.IsInBitDefConfigBitRange(configRow, configTables);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsInBitDefConfigBitRange_RowOutsideRange_ReturnsFalse()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [1] = "10", [2] = "0" }));
            var configTables = new List<BitDefTable> { table };
            var configRow = new EfuseConfigMainRow { Lsb = 20, Msb = 25 };

            // Act
            bool result = _target.IsInBitDefConfigBitRange(configRow, configTables);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void AddDummyFuseForBankMon_NoMonRowInArraySize_ReturnsUnchangedTables()
        {
            // Arrange - EfuseArraySizeSheet has no row ending with "MON"
            BitDefTable table = NewTable("MON");
            var bitDefTables = new List<BitDefTable> { table };

            // Act
            List<BitDefTable> result = _target.AddDummyFuseForBankMon(bitDefTables);

            // Assert
            Assert.AreSame(bitDefTables, result);
        }

        [TestMethod]
        public void AddDummyFuseForBankMon_MonRowFoundAndGapExists_AddsDummyRow()
        {
            // Arrange
            _inputData.EfuseArraySizeSheet.Rows.Add(new EfuseArraySizeRow { FuseArray = "Bank_MON", OfBits = "32" });
            BitDefTable table = NewTable("MON");
            table.BitMode = "";
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [1] = "5", [2] = "0" }));
            var bitDefTables = new List<BitDefTable> { table };

            // Act
            List<BitDefTable> result = _target.AddDummyFuseForBankMon(bitDefTables);

            // Assert
            BitDefRow dummyRow = result[0].Rows[1];
            Assert.AreEqual(2, result[0].Rows.Count);
            Assert.AreEqual("MON_condition_TE", dummyRow.RowData[0]);
            Assert.AreEqual("31", dummyRow.RowData[1]);
            Assert.AreEqual("6", dummyRow.RowData[2]);
            Assert.AreEqual("26", dummyRow.RowData[3]);
            Assert.AreEqual("N/A", dummyRow.RowData[4]);
            Assert.AreEqual("N/A", dummyRow.RowData[5]);
            Assert.AreEqual("N/A", dummyRow.RowData[6]);
            Assert.AreEqual("N/A", dummyRow.RowData[7]);
            Assert.AreEqual("SLT", dummyRow.RowData[8]);
            Assert.AreEqual("0", dummyRow.RowData[9]);
            Assert.AreEqual("0", dummyRow.RowData[10]);
            Assert.AreEqual("N/A", dummyRow.RowData[11]);
            Assert.AreEqual("app", dummyRow.RowData[12]);
            Assert.AreEqual("0", dummyRow.RowData[13]);
            Assert.AreEqual("Default", dummyRow.RowData[14]);
            Assert.AreEqual("0", dummyRow.RowData[15]);
            Assert.AreEqual(3, dummyRow.RowNum);
        }

        [TestMethod]
        public void AddDummyFuseForBankMon_MonRowFoundNoGap_DoesNotAddDummyRow()
        {
            // Arrange - last bit already reaches the max bit from array size, so no gap remains
            _inputData.EfuseArraySizeSheet.Rows.Add(new EfuseArraySizeRow { FuseArray = "Bank_MON", OfBits = "8" });
            BitDefTable table = NewTable("MON");
            table.BitMode = "";
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [1] = "7", [2] = "0" }));
            var bitDefTables = new List<BitDefTable> { table };

            // Act
            List<BitDefTable> result = _target.AddDummyFuseForBankMon(bitDefTables);

            // Assert
            Assert.AreEqual(1, result[0].Rows.Count);
        }

        [TestMethod]
        public void GetConfigRowInBitDefTables_MissingLsbIdx_SkipsCondBranch()
        {
            // Arrange - LsbBitIdx == -1 disables the cond-algorithm branch entirely, and with
            // DefaultOrRealIdx/DefaultValueIdx also disabled the config table ends up empty.
            var table = new BitDefTable
            {
                BlockName = "Config_ABC",
                LsbBitIdx = -1,
                MsbBitIdx = 1,
                AlgorithmIdx = 12,
                DefaultOrRealIdx = -1,
                DefaultValueIdx = -1
            };
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [1] = "7", [12] = "cond" }));
            _target.FinalBitDefTables = [table];

            // Act
            List<BitDefTable> result = _target.GetConfigRowInBitDefTables();

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GenerateBinCheckFlow_BuildsTestRowsAndFinalBinTableRow()
        {
            // Act
            SubFlowSheet result = _target.GenerateBinCheckFlow(["F_EFUSE_Bin5"]);

            // Assert
            Assert.AreEqual("Flow_efuse_BinCheck", result.Name);
            FlowRow testRow = result.Rows.First(r => r.Opcode == OpCode.Test);
            Assert.AreEqual("F_EFUSE_Bin5", testRow.Parameter);
            Assert.AreEqual("F_EFUSE_BinCheck", testRow.FailAction);
            Assert.AreEqual("Bin5", testRow.DeviceName);
            Assert.IsTrue(result.Rows.Exists(r => r.Opcode == OpCode.BinTable && r.Parameter == "Bin_EFUSE_BinCheck"));
        }

        [TestMethod]
        public void CheckDatabaseRevision_RevisionNotPositive_NoError()
        {
            // Arrange
            _inputData.EfuseDatabaseRevision = -1;
            BitDefTable table = NewTable("CFG");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [0] = "efuse_database_revision" });

            // Act
            _target.CheckDatabaseRevision(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckDatabaseRevision_FieldNameDoesNotMatch_NoError()
        {
            // Arrange
            _inputData.EfuseDatabaseRevision = 5;
            BitDefTable table = NewTable("CFG");
            BitDefRow row = NewRow(new SortedDictionary<int, string> { [0] = "some_other_field" });

            // Act
            _target.CheckDatabaseRevision(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckDatabaseRevision_RealDefaultOrReal_AddsMismatchDefaultTypeError()
        {
            // Arrange - revision 5, binLength 4 -> expected binary "b0101"; all four columns match
            // so only the DefaultOrReal=="real" check should add an error.
            _inputData.EfuseDatabaseRevision = 5;
            BitDefTable table = NewTable("CFG");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [0] = "efuse_database_revision",
                [3] = "4",
                [9] = "b0101",
                [10] = "b0101",
                [13] = "b0101",
                [14] = "real",
                [15] = "b0101"
            });

            // Act
            _target.CheckDatabaseRevision(table, row);

            // Assert
            Assert.IsTrue(ErrorReportManager.Contains(EFuseErrorType.E_MismatchDefaultType_01.FormatMessage()));
        }

        [TestMethod]
        public void CheckDatabaseRevision_BinaryMatchesAndNotReal_NoError()
        {
            // Arrange
            _inputData.EfuseDatabaseRevision = 5;
            BitDefTable table = NewTable("CFG");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [0] = "efuse_database_revision",
                [3] = "4",
                [9] = "b0101",
                [10] = "b0101",
                [13] = "b0101",
                [14] = "default",
                [15] = "b0101"
            });

            // Act
            _target.CheckDatabaseRevision(table, row);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckDatabaseRevision_BinaryMismatchNonRfDevice_AddsMismatchRevisionError()
        {
            // Arrange - device defaults to AP, so a mismatch is reported rather than auto-fixed
            _inputData.EfuseDatabaseRevision = 5;
            BitDefTable table = NewTable("CFG");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [0] = "efuse_database_revision",
                [3] = "4",
                [9] = "wrong",
                [10] = "b0101",
                [13] = "b0101",
                [14] = "default",
                [15] = "b0101"
            });

            // Act
            _target.CheckDatabaseRevision(table, row);

            // Assert
            Assert.IsTrue(ErrorReportManager.Contains(EFuseErrorType.E_MismatchRevision_01.FormatMessage()));
        }

        [TestMethod]
        public void CheckDatabaseRevision_BinaryMismatchRfDevice_OverwritesRowAndAddsToolWarning()
        {
            // Arrange - on RF devices, a mismatch is silently corrected by the tool with a warning
            LocalSpecs.Options.Device = EnumDevice.RF;
            _inputData.EfuseDatabaseRevision = 5;
            BitDefTable table = NewTable("CFG");
            BitDefRow row = NewRow(new SortedDictionary<int, string>
            {
                [0] = "efuse_database_revision",
                [3] = "4",
                [9] = "wrong",
                [10] = "b0101",
                [13] = "b0101",
                [14] = "default",
                [15] = "b0101"
            });

            // Act
            _target.CheckDatabaseRevision(table, row);

            // Assert
            Assert.IsTrue(ErrorReportManager.Contains(EFuseErrorType.W_RuleViolationRevision_01.FormatMessage()));
            Assert.AreEqual("b0101", row.RowData[9]);
        }
    }
}
