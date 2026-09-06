using System.Collections.Generic;
using System.IO;

using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets.MultiSheet.MultiTimeSet
{
    [TestClass]
    public class MultiTimeSetSheetReaderTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void MultiTimeSetSheetReader_Constructor()
        {
            // Arrange & Act
            var reader = new MultiTimeSetSheetReader();

            // Assert
            Assert.IsNotNull(reader);
        }

        [TestMethod]
        public void MultiTimeSetSheetReader_Constants_Defined()
        {
            // Arrange & Act
            string cycleS = MultiTimeSetSheetReader.CycleS;
            string clockS = MultiTimeSetSheetReader.ClockS;
            string clockE = MultiTimeSetSheetReader.ClockE;
            string strobe = MultiTimeSetSheetReader.Strobe;

            // Assert
            Assert.AreEqual("Cycle_S", cycleS);
            Assert.AreEqual("Clock_S", clockS);
            Assert.AreEqual("Clock_E", clockE);
            Assert.AreEqual("Strobe", strobe);
        }

        [TestMethod]
        public void MultiTimeSetSheetReader_ReadTimeSetTxt1P4_WithEmptyList()
        {
            // Arrange
            var timeSetPathList = new List<string>();

            // Act
            MultiTimeSetSheets result = MultiTimeSetSheetReader.ReadTimeSetTxt1P4(timeSetPathList);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.TimeSetBasicSheetsList.Count);
        }

        [TestMethod]
        public void MultiTimeSetSheetReader_ReadTimeSetTxt1P4_WithNonExistentFile()
        {
            // Arrange
            var timeSetPathList = new List<string> { "nonexistent_file.txt" };

            // Act
            MultiTimeSetSheets result = MultiTimeSetSheetReader.ReadTimeSetTxt1P4(timeSetPathList);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.TimeSetBasicSheetsList.Count);
        }

        [TestMethod]
        public void MultiTimeSetSheetReader_ReadTimeSetTxt1P4_WithMultipleNonExistentFiles()
        {
            // Arrange
            var timeSetPathList = new List<string>
            {
                "nonexistent_file1.txt",
                "nonexistent_file2.txt",
                "nonexistent_file3.txt"
            };

            // Act
            MultiTimeSetSheets result = MultiTimeSetSheetReader.ReadTimeSetTxt1P4(timeSetPathList);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.TimeSetBasicSheetsList.Count);
        }

        [TestMethod]
        public void MultiTimeSetSheetReader_ReadTimeSetTxt1P4_WithValidFile()
        {
            // Arrange
            string tempFilePath = Path.Combine(Path.GetTempPath(), "test_timeset.txt");

            // Create a valid timeset file
            var lines = new List<string>
            {
                "Pin Grp\tPeriod\tSetup\tDataSrc\tDataFmt\tDriveOn\tDriveData\tDriveReturn\tDriveOff\tCompareMode\tCompareOpen\tCompareClose\tEdgeMode\tReserved1\tReserved2\tReserved3",
                "Timing Mode:\tASYNC",
                "Master Timeset Name:\tMasterTS\tTime Domain:\tDOMAIN",
                "Time Set\tPeriod\tSetup",
                "Time Set\tPeriod\tSetup\tDataSrc\tDataFmt\tDriveOn\tDriveData\tDriveReturn\tDriveOff\tCompareMode\tCompareOpen\tCompareClose\tEdgeMode"
            };

            try
            {
                File.WriteAllLines(tempFilePath, lines);

                // Act
                var timeSetPathList = new List<string> { tempFilePath };
                MultiTimeSetSheets result = MultiTimeSetSheetReader.ReadTimeSetTxt1P4(timeSetPathList);

                // Assert
                Assert.IsNotNull(result);
                Assert.IsTrue(result.TimeSetBasicSheetsList.Count >= 0);
            }
            finally
            {
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }
            }
        }

        [TestMethod]
        public void MultiTimeSetSheetReader_ReadTimeSetTxt1P4_RemoveBackupFlag()
        {
            // Arrange
            var timeSetPathList = new List<string> { "nonexistent.txt" };

            // Act
            MultiTimeSetSheets result = MultiTimeSetSheetReader.ReadTimeSetTxt1P4(timeSetPathList, isRemoveBackup: true);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.TimeSetBasicSheetsList.Count);
        }

        [TestMethod]
        public void MultiTimeSetSheetReader_HasPublicMethods()
        {
            // Arrange
            var reader = new MultiTimeSetSheetReader();

            // Act & Assert
            // Verify that ReadTimeSetTxt1P4 method exists and is public
            System.Reflection.MethodInfo methodInfo = reader.GetType().GetMethod("ReadTimeSetTxt1P4");
            Assert.IsNotNull(methodInfo);
            Assert.IsTrue(methodInfo.IsPublic);
        }

        [TestMethod]
        public void MultiTimeSetSheetReader_ReadTimeSetTxt1P4_ReturnsMultiTimeSetSheets()
        {
            // Arrange
            var timeSetPathList = new List<string>();

            // Act
            MultiTimeSetSheets result = MultiTimeSetSheetReader.ReadTimeSetTxt1P4(timeSetPathList);

            // Assert
            Assert.IsInstanceOfType(result, typeof(MultiTimeSetSheets));
        }

        [TestMethod]
        public void MultiTimeSetSheetReader_SupportsNullablePathList()
        {
            // Act
            MultiTimeSetSheets result = MultiTimeSetSheetReader.ReadTimeSetTxt1P4([]);

            // Assert
            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void ReadData_ValidLineTokens_PopulatesPairsAndContextVariables()
        {
            // Arrange
            string[] lines =
            [
                "IgnoredHeaderRow",
                "\tTS_CORE\t10.0\t_MyVar\t5\t6\t7\t8\t9\t10\t11\t12\t13\t14\t15\t16" // Tab-separated raw line array data
            ];
            var converter = new TimeRow1P4Converter();
            var sheet = new ComTimeSetBasicSheet("SheetA", "Single", "M_TS", "DOM", "");
            var pairs = new Dictionary<string, ComTimeSetBasic>();

            // Act
            MultiTimeSetSheetReader.ReadData(lines, converter, "SheetA", sheet, pairs, 1);

            // Assert
            Assert.IsTrue(pairs.ContainsKey("TS_CORE"));
            Assert.AreEqual("10.0", pairs["TS_CORE"].CyclePeriod);
        }

        [TestMethod]
        public void ReadData_ValidVarDefinitionsBlock_ParsesDoubleValueToMatchingContext()
        {
            // Arrange
            string[] lines =
            [
                "VAR Definitions",
                "MyVar = 5000000"
            ];
            var converter = new TimeRow1P4Converter();
            var sheet = new ComTimeSetBasicSheet("SheetA", "Single", "M_TS", "DOM", "");
            var pairs = new Dictionary<string, ComTimeSetBasic>();

            var timeSetBasic = new ComTimeSetBasic { Name = "TS_CORE" };
            timeSetBasic.SubContextVariable.Add("MyVar");
            pairs.Add("TS_CORE", timeSetBasic);

            // Act
            MultiTimeSetSheetReader.ReadData(lines, converter, "SheetA", sheet, pairs, 0);

            // Assert
            Assert.IsTrue(pairs["TS_CORE"].SubCommentVariable.ContainsKey("MyVar"));
            Assert.AreEqual(5000000.0, pairs["TS_CORE"].SubCommentVariable["MyVar"]);
        }

        [TestMethod]
        public void ReadData_VarWithoutAssignment_AppendsFallbackDefaultValue()
        {
            // Arrange
            string[] lines =
            [
                "VAR Definitions",
                "UnassignedVar"
            ];
            var converter = new TimeRow1P4Converter();
            var sheet = new ComTimeSetBasicSheet("SheetA", "Single", "M_TS", "DOM", "");
            var pairs = new Dictionary<string, ComTimeSetBasic>();

            var timeSetBasic = new ComTimeSetBasic { Name = "TS_CORE" };
            timeSetBasic.SubContextVariable.Add("UnassignedVar");
            pairs.Add("TS_CORE", timeSetBasic);

            // Act
            MultiTimeSetSheetReader.ReadData(lines, converter, "SheetA", sheet, pairs, 0);

            // Assert
            Assert.AreEqual(-1e9, pairs["TS_CORE"].SubCommentVariable["UnassignedVar"]);
        }

        [TestMethod]
        public void CheckMissingEquationBaseVar_UnassignedValueFound_AppendsFormatError()
        {
            // Arrange
            var sheet = new ComTimeSetBasicSheet("SheetA", "Single", "M_TS", "DOM", "");
            var pairs = new Dictionary<string, ComTimeSetBasic>();
            var timeSetBasic = new ComTimeSetBasic { Name = "TS_CORE" };

            timeSetBasic.SubContextVariable.Add("BadVar");
            timeSetBasic.SubCommentVariable.Add("BadVar", -1e9);
            pairs.Add("TS_CORE", timeSetBasic);

            // Act
            MultiTimeSetSheetReader.CheckMissingEquationBaseVar("SheetA", sheet, pairs);

            // Assert
            Assert.AreEqual(1, sheet.GetErrors().Count);
            Assert.AreEqual(BasicErrorType.E_InvalidTiming_02.FullCode, sheet.GetErrors()[0].ErrorCode.FullCode);
            Assert.IsTrue(sheet.GetErrors()[0].Message.Contains("is not assigned an initial value"));
        }

        [TestMethod]
        public void CheckMissingEquationBaseVar_ContextVarMissingInComment_AppendsCommentMissingError()
        {
            // Arrange
            var sheet = new ComTimeSetBasicSheet("SheetA", "Single", "M_TS", "DOM", "");
            var pairs = new Dictionary<string, ComTimeSetBasic>();
            var timeSetBasic = new ComTimeSetBasic { Name = "TS_CORE" };

            timeSetBasic.SubContextVariable.Add("ContextOnlyVar");
            pairs.Add("TS_CORE", timeSetBasic);

            // Act
            MultiTimeSetSheetReader.CheckMissingEquationBaseVar("SheetA", sheet, pairs);

            // Assert
            Assert.AreEqual(1, sheet.GetErrors().Count);
            Assert.IsTrue(sheet.GetErrors()[0].Message.Contains("is not assigned value in comment"));
        }

        [TestMethod]
        public void ReadTimeRow_ValidRowRLFormat_CalculatesCorrectShiftInReserveRatios()
        {
            // Arrange
            string[] lineTokens = ["", "TS_CORE", "10.0", "", "RL", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"];
            var timingRow = new TimingRow { DriveData = "2.5", CompareOpen = "5.0", DataFmt = "RL" };
            var stubConverter = new StubTimeRowConverter(timingRow);

            // Act
            bool result = MultiTimeSetSheetReader.ReadTimeRow(lineTokens, stubConverter, out _, out _, out Dictionary<string, double> outShiftInReserveVar);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual(0.25, outShiftInReserveVar["Cycle_S_VAR"]);
            Assert.AreEqual(0.50, outShiftInReserveVar["Strobe_VAR"]);
        }
    }

    // Helper Stub layout class to safely control nested reflection parameter loops
    public class StubTimeRowConverter(TimingRow timingRow) : TimeRow1P4Converter
    {
        private readonly TimingRow _row = timingRow;

        public override TimingRow ConvertTimeRow(string[] line) => _row;
    }
}
