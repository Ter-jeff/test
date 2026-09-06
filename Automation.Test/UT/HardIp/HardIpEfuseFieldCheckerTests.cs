using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

using CommonLib.ErrorReport;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using TestPlanLib.Efuse.Input;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardIpEfuseFieldCheckerTests
    {
        private readonly HardIpEfuseFieldChecker _checker = new([]);

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

        private static HardIpPattern NewPattern(string miscInfo)
        {
            return new HardIpPattern { SheetName = "S1", RowNum = 0, MiscInfo = miscInfo, Pattern = new PatternClass("") };
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

        private static EFuseHardIpSheet NewHardIpSheet(params EFuseHardIpRow[] rows)
        {
            var sheet = new EFuseHardIpSheet("eFuse_HardIP_Table");
            sheet.HeaderIndex["Width"] = 5;
            sheet.HeaderIndex["Job"] = 6;
            sheet.HeaderIndex["eFuse Definition"] = 7;
            foreach (EFuseHardIpRow row in rows)
            {
                sheet.Rows.Add(row);
            }
            return sheet;
        }

        #region CollectCateNameFromEfuseHardIp

        [TestMethod]
        public void CollectCateNameFromEfuseHardIp_HipEfuseReadWithMCateName_ExtractsMultipleCateNames()
        {
            // Arrange
            var planDic = new Dictionary<string, HardIpSheet>
            {
                { "S1", new HardIpSheet { Rows = [NewPattern("HIP_EFUSE_READ;m_catename:Cat1+Cat2")] } }
            };

            // Act
            List<string> result = _checker.CollectCateNameFromEfuseHardIp(planDic);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "Cat1", "Cat2" }, result);
        }

        [TestMethod]
        public void CollectCateNameFromEfuseHardIp_HardIpFuseWrite_ExtractsCateName()
        {
            // Arrange
            var planDic = new Dictionary<string, HardIpSheet>
            {
                { "S1", new HardIpSheet { Rows = [NewPattern("HardIPFuseWrite;m_catename:CatX")] } }
            };

            // Act
            List<string> result = _checker.CollectCateNameFromEfuseHardIp(planDic);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "CatX" }, result);
        }

        [TestMethod]
        public void CollectCateNameFromEfuseHardIp_MiscInfoWithoutRelevantKeyword_ReturnsEmpty()
        {
            // Arrange
            var planDic = new Dictionary<string, HardIpSheet>
            {
                { "S1", new HardIpSheet { Rows = [NewPattern("Something;m_catename:Cat1")] } }
            };

            // Act
            List<string> result = _checker.CollectCateNameFromEfuseHardIp(planDic);

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void CollectCateNameFromEfuseHardIp_ArgumentWithoutColon_SkippedSafely()
        {
            // Arrange
            var planDic = new Dictionary<string, HardIpSheet>
            {
                { "S1", new HardIpSheet { Rows = [NewPattern("HIP_EFUSE_READ;NoColonHere;m_catename:CatY")] } }
            };

            // Act
            List<string> result = _checker.CollectCateNameFromEfuseHardIp(planDic);

            // Assert
            CollectionAssert.AreEqual(new List<string> { "CatY" }, result);
        }

        #endregion

        #region CheckEfuseRealValue

        [TestMethod]
        public void CheckEfuseRealValue_RealValueMissingFromHardIpSheet_AddsWarning()
        {
            // Arrange
            BitDefTable table = NewTable();
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [5] = "real", [0] = "FieldA" }));
            EFuseHardIpSheet hardIpSheet = NewHardIpSheet();

            // Act
            _checker.CheckEfuseRealValue([table], hardIpSheet);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckEfuseRealValue_RealValueFoundInHardIpSheet_NoWarning()
        {
            // Arrange
            BitDefTable table = NewTable();
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [5] = "real", [0] = "FieldA" }));
            EFuseHardIpSheet hardIpSheet = NewHardIpSheet(new EFuseHardIpRow { Efusedefinition = "FieldA" });

            // Act
            _checker.CheckEfuseRealValue([table], hardIpSheet);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckEfuseRealValue_NonRealAlgorithm_SkippedEvenWhenMissing()
        {
            // Arrange
            BitDefTable table = NewTable();
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [5] = "base", [0] = "FieldA" }));
            EFuseHardIpSheet hardIpSheet = NewHardIpSheet();

            // Act
            _checker.CheckEfuseRealValue([table], hardIpSheet);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        #endregion

        #region CheckEfuseHardIpAndEfuseBitDefTable

        [TestMethod]
        public void CheckEfuseHardIpAndEfuseBitDefTable_PreWriteNotInCateNameList_AddsMissingInPreWriteError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            var hardIpRow = new EFuseHardIpRow { Efusedefinition = "FieldA", Bank = "ECID" };
            EFuseHardIpSheet hardIpSheet = NewHardIpSheet(hardIpRow);

            // Act
            _checker.CheckEfuseHardIpAndEfuseBitDefTable([table], hardIpSheet, []);

            // Assert
            Assert.IsTrue(ErrorReportManager.GetErrorCount() >= 1);
        }

        [TestMethod]
        public void CheckEfuseHardIpAndEfuseBitDefTable_PreWriteFoundInCateNameList_NoMissingInPreWriteError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [0] = "FieldA", [3] = "8", [4] = "CP1", [5] = "real" }));
            var hardIpRow = new EFuseHardIpRow { Efusedefinition = "FieldA", Bank = "ECID", Width = "8", Job = "CP1" };
            EFuseHardIpSheet hardIpSheet = NewHardIpSheet(hardIpRow);

            // Act
            _checker.CheckEfuseHardIpAndEfuseBitDefTable([table], hardIpSheet, ["FieldA"]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckEfuseHardIpAndEfuseBitDefTable_BankNotFound_AddsMissingFieldNameError()
        {
            // Arrange - no BitDefTable matches the hardip row's Bank at all
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [0] = "FieldA", [5] = "real" }));
            var hardIpRow = new EFuseHardIpRow { Efusedefinition = "FieldA", Bank = "NONEXISTENT" };
            EFuseHardIpSheet hardIpSheet = NewHardIpSheet(hardIpRow);

            // Act
            _checker.CheckEfuseHardIpAndEfuseBitDefTable([table], hardIpSheet, ["FieldA"]);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckEfuseHardIpAndEfuseBitDefTable_TargetNotFoundInBank_AddsMissingFieldNameError()
        {
            // Arrange - bank matches, but no row for this field name
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [0] = "OtherField", [5] = "real" }));
            var hardIpRow = new EFuseHardIpRow { Efusedefinition = "FieldA", Bank = "ECID" };
            EFuseHardIpSheet hardIpSheet = NewHardIpSheet(hardIpRow);

            // Act
            _checker.CheckEfuseHardIpAndEfuseBitDefTable([table], hardIpSheet, ["FieldA"]);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckEfuseHardIpAndEfuseBitDefTable_WidthMismatch_AddsMismatchFieldWidthError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [0] = "FieldA", [3] = "8", [4] = "CP1", [5] = "real" }));
            var hardIpRow = new EFuseHardIpRow { Efusedefinition = "FieldA", Bank = "ECID", Width = "4", Job = "CP1" };
            EFuseHardIpSheet hardIpSheet = NewHardIpSheet(hardIpRow);

            // Act
            _checker.CheckEfuseHardIpAndEfuseBitDefTable([table], hardIpSheet, ["FieldA"]);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckEfuseHardIpAndEfuseBitDefTable_JobMismatch_AddsMismatchFieldJobError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [0] = "FieldA", [3] = "8", [4] = "CP1", [5] = "real" }));
            var hardIpRow = new EFuseHardIpRow { Efusedefinition = "FieldA", Bank = "ECID", Width = "8", Job = "FT1" };
            EFuseHardIpSheet hardIpSheet = NewHardIpSheet(hardIpRow);

            // Act
            _checker.CheckEfuseHardIpAndEfuseBitDefTable([table], hardIpSheet, ["FieldA"]);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckEfuseHardIpAndEfuseBitDefTable_NotRealAlgorithm_AddsMismatchRealValueError()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [0] = "FieldA", [3] = "8", [4] = "CP1", [5] = "base" }));
            var hardIpRow = new EFuseHardIpRow { Efusedefinition = "FieldA", Bank = "ECID", Width = "8", Job = "CP1" };
            EFuseHardIpSheet hardIpSheet = NewHardIpSheet(hardIpRow);

            // Act
            _checker.CheckEfuseHardIpAndEfuseBitDefTable([table], hardIpSheet, ["FieldA"]);

            // Assert
            Assert.AreEqual(1, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckEfuseHardIpAndEfuseBitDefTable_AllMatch_NoErrors()
        {
            // Arrange
            BitDefTable table = NewTable("ECID");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [0] = "FieldA", [3] = "8", [4] = "CP1", [5] = "real" }));
            var hardIpRow = new EFuseHardIpRow { Efusedefinition = "FieldA", Bank = "ECID", Width = "8", Job = "CP1" };
            EFuseHardIpSheet hardIpSheet = NewHardIpSheet(hardIpRow);

            // Act
            _checker.CheckEfuseHardIpAndEfuseBitDefTable([table], hardIpSheet, ["FieldA"]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void CheckEfuseHardIpAndEfuseBitDefTable_BankNameWithUnderscoresIgnored_StillMatches()
        {
            // Arrange - "_" characters are stripped before comparing bank names
            BitDefTable table = NewTable("E_C_I_D");
            table.Rows.Add(NewRow(new SortedDictionary<int, string> { [0] = "FieldA", [3] = "8", [4] = "CP1", [5] = "real" }));
            var hardIpRow = new EFuseHardIpRow { Efusedefinition = "FieldA", Bank = "ECID", Width = "8", Job = "CP1" };
            EFuseHardIpSheet hardIpSheet = NewHardIpSheet(hardIpRow);

            // Act
            _checker.CheckEfuseHardIpAndEfuseBitDefTable([table], hardIpSheet, ["FieldA"]);

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        #endregion

        #region GetEfuseBitDefSheet

        [TestMethod]
        public void GetEfuseBitDefSheet_MatchingSheetName_ReturnsSheet()
        {
            // Arrange
            using var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("Other");
            package.Workbook.Worksheets.Add("eFuse_HardIP_Table");

            // Act
            ExcelWorksheet result = HardIpEfuseFieldChecker.GetEfuseBitDefSheet(package.Workbook.Worksheets, "efuse_hardip_table");

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual("eFuse_HardIP_Table", result.Name);
        }

        [TestMethod]
        public void GetEfuseBitDefSheet_NoMatchingSheetName_ReturnsNull()
        {
            // Arrange
            using var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("Other");

            // Act
            ExcelWorksheet? result = HardIpEfuseFieldChecker.GetEfuseBitDefSheet(package.Workbook.Worksheets, "eFuse_HardIP_Table");

            // Assert
            Assert.IsNull(result);
        }

        #endregion
    }
}
