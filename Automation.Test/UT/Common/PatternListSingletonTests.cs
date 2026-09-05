using System;
using System.Collections.Generic;
using System.IO;

using Automation.Singleton;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.PatternListCsvFile;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class PatternListSingletonTests
    {
        private readonly List<string> _tempFiles = [];

        [TestInitialize]
        public void Setup()
        {
            PatternListSingleton.Initialize();
        }

        [TestCleanup]
        public void Cleanup()
        {
            PatternListSingleton.Initialize();
            foreach (string file in _tempFiles)
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
            _tempFiles.Clear();
        }

        private string NewTempCsvFile(string[] lines)
        {
            string path = Path.Combine(Path.GetTempPath(), $"PatternListSingletonTests_{Guid.NewGuid():N}.csv");
            File.WriteAllLines(path, lines);
            _tempFiles.Add(path);
            return path;
        }

        [TestMethod]
        public void GetInstance_NonexistentFile_IsTsmcPatternListReturnsTrueWithoutTouchingDisk()
        {
            // Arrange - a path that does not exist skips all file I/O in the constructor
            PatternListSingleton instance = PatternListSingleton.GetInstance(@"Z:\definitely_missing_12345.csv", "");

            // Act
            bool isTsmc = instance.IsCompiledPatternDashboard();

            // Assert
            Assert.IsTrue(isTsmc);
        }

        [TestMethod]
        public void GetPatternData_NonexistentFile_ReturnsEmptyPatternDictionary()
        {
            // Arrange
            PatternListSingleton instance = PatternListSingleton.GetInstance(@"Z:\definitely_missing_12345.csv", "");

            // Act
            Dictionary<string, OriPatListItem> result = instance.GetPatternData();

            // Assert
            Assert.AreEqual(0, result.Count);
            Assert.AreSame(instance.PatternDictionary, result);
        }

        [TestMethod]
        public void MergeWithPatListCsv_NonexistentFile_NoOp()
        {
            // Arrange
            PatternListSingleton instance = PatternListSingleton.GetInstance(@"Z:\definitely_missing_12345.csv", "");

            // Act
            instance.UpdatePatternDashboardWithCompiledPatCsv([]);

            // Assert
            Assert.AreEqual(string.Empty, instance.CompiledPatternDashboardFile ?? string.Empty);
        }

        [TestMethod]
        public void UpdateCsvByLatestPattern_NonexistentFile_ReturnsNull()
        {
            // Arrange
            PatternListSingleton instance = PatternListSingleton.GetInstance(@"Z:\definitely_missing_12345.csv", "");

            // Act
            List<string>? result = instance.UpdateCsvByLatestPattern([]);

            // Assert
            Assert.IsNull(result);
        }

        [TestMethod]
        public void NormalizeTwDataArray_ShorterThanHeaderCount_PadsWithEmptyStrings()
        {
            // Arrange
            string[] result = ["A", "B"];

            // Act
            PatternListSingleton.NormalizeTwDataArray(ref result, 4);

            // Assert
            Assert.AreEqual(4, result.Length);
            Assert.AreEqual("A", result[0]);
            Assert.AreEqual("B", result[1]);
            Assert.AreEqual("", result[2]);
            Assert.AreEqual("", result[3]);
        }

        [TestMethod]
        public void NormalizeTwDataArray_AlreadyLongEnough_LeavesArrayUnchanged()
        {
            // Arrange
            string[] result = ["A", "B", "C"];

            // Act
            PatternListSingleton.NormalizeTwDataArray(ref result, 2);

            // Assert
            CollectionAssert.AreEqual(new[] { "A", "B", "C" }, result);
        }

        [TestMethod]
        public void GetTwField_KnownHeader_ReturnsMappedValue()
        {
            // Arrange
            string[] dataArray = ["1", "PAT_01"];
            var headerIndexDic = new Dictionary<string, int> { ["#"] = 0, ["PATTERN"] = 1 };

            // Act
            string result = PatternListSingleton.GetTwField(dataArray, headerIndexDic, "Pattern");

            // Assert
            Assert.AreEqual("PAT_01", result);
        }

        [TestMethod]
        public void GetTwField_UnknownHeader_ReturnsEmptyString()
        {
            // Arrange
            string[] dataArray = ["1", "PAT_01"];
            var headerIndexDic = new Dictionary<string, int> { ["#"] = 0, ["PATTERN"] = 1 };

            // Act
            string result = PatternListSingleton.GetTwField(dataArray, headerIndexDic, "Org");

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void BuildTwOriPatItem_MapsAllFieldsFromDataArray()
        {
            // Arrange
            var headerIndexDic = new Dictionary<string, int>
            {
                ["#"] = 0,
                ["PATTERN"] = 1,
                ["LATEST VERSION"] = 2,
                ["T/P CATEGORY"] = 3
            };
            string[] dataArray = ["7", "PAT_02", "V2", "CAT2"];

            // Act
            OriPatListItem result = PatternListSingleton.BuildTwOriPatItem(dataArray, headerIndexDic, 7);

            // Assert
            Assert.AreEqual(7, result.RowNum);
            Assert.AreEqual("7", result.Idx);
            Assert.AreEqual("PAT_02", result.Pattern);
            Assert.AreEqual("V2", result.LatestVersion);
            Assert.AreEqual("CAT2", result.TpCategory);
        }

        [TestMethod]
        public void GetInstance_TwFormatCsvFile_PopulatesPatternDictionaryFromTwReader()
        {
            // Arrange - header matches the TW-format detection (Pattern...Release Notes...T/P Category)
            string path = NewTempCsvFile(
            [
                "#,Pattern,Latest Version,Release Date,USE/No Use,DRI,Release Notes,Radar #,Org,Type Spec,Timeset Version,File Versions,OpCode,ScanMode,Halt,Compilation,HLV,T/P Category,ScanSetupTSet,Original Timing Mode,Check,CheckComment",
                "1,PAT_TW_01,V1,2024-01-01,USE,DRI1,Notes1,RADAR1,ORG1,TYPE1,TSV1,FV1,OP1,SCAN1,HALT1,COMP1,HLV1,TPCAT1,STSET1,OTM1,CHK1,CC1"
            ]);

            // Act
            PatternListSingleton instance = PatternListSingleton.GetInstance(path, "");

            // Assert
            Assert.IsTrue(instance.IsCompiledPatternDashboard());
            Assert.IsTrue(instance.PatternDictionary.ContainsKey("PAT_TW_01"));
            OriPatListItem item = instance.PatternDictionary["PAT_TW_01"];
            Assert.AreEqual("OP1", item.OpCode);
            Assert.AreEqual("TPCAT1", item.TpCategory);
        }

        [TestMethod]
        public void GetInstance_NonTwFormatCsvFile_PopulatesPatternDictionaryFromNonTwReader()
        {
            // Arrange - header matches the non-TW detection (Pattern...Timeset Latest...File Versions)
            // but has no "Release Notes" or "T/P Category" column, so JudgeTw reports false.
            string path = NewTempCsvFile(
            [
                "#,Pattern,Latest Version,Release Date,USE/No Use,DRI,Radar #,Org,Type Spec,Timeset Latest,File Versions,OpCode,ScanMode,Halt,Compilation,HLV,ScanSetupTSet",
                "1,PAT_NONTW_01,LV1,2024-02-02,USE1,DRI1,RADAR1,ORG1,TYPE1,TSL1,FV1,OP1,SCAN1,HALT1,COMP1,HLV1,STSET1"
            ]);

            // Act
            PatternListSingleton instance = PatternListSingleton.GetInstance(path, "");

            // Assert
            Assert.IsFalse(instance.IsCompiledPatternDashboard());
            Dictionary<string, OriPatListItem> data = instance.GetPatternData();
            Assert.IsTrue(data.ContainsKey("PAT_NONTW_01#FV1"));
            OriPatListItem item = data["PAT_NONTW_01#FV1"];
            Assert.AreEqual("PAT_NONTW_01", item.Pattern);
            Assert.AreEqual("OP1", item.OpCode);
            Assert.AreEqual("HLV1", item.HLv);
        }

        [TestMethod]
        public void UpdateCsvByLatestPattern_NonTwFileWithNewerVersion_UpdatesFileVersionsAndReturnsMessage()
        {
            // Arrange
            string path = NewTempCsvFile(
            [
                "#,Pattern,Latest Version,Release Date,USE/No Use,DRI,Radar #,Org,Type Spec,Timeset Latest,File Versions,OpCode,ScanMode,Halt,Compilation,HLV,ScanSetupTSet",
                "1,PAT_NONTW_02,LV1,2024-02-02,USE1,DRI1,RADAR1,ORG1,TYPE1,TSL1,PAT_NONTW_02_1,OP1,SCAN1,HALT1,COMP1,HLV1,STSET1"
            ]);
            PatternListSingleton instance = PatternListSingleton.GetInstance(path, "");
            var latestPatternDict = new Dictionary<string, string> { ["PAT_NONTW_02"] = "2" };

            // Act
            List<string>? result = instance.UpdateCsvByLatestPattern(latestPatternDict);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
            StringAssert.Contains(result[0], "PAT_NONTW_02");
        }

        [TestMethod]
        public void UpdateCsvByLatestPattern_NonTwFileAlreadyAtLatestVersion_ReturnsNoMessages()
        {
            // Arrange
            string path = NewTempCsvFile(
            [
                "#,Pattern,Latest Version,Release Date,USE/No Use,DRI,Radar #,Org,Type Spec,Timeset Latest,File Versions,OpCode,ScanMode,Halt,Compilation,HLV,ScanSetupTSet",
                "1,PAT_NONTW_03,LV1,2024-02-02,USE1,DRI1,RADAR1,ORG1,TYPE1,TSL1,PAT_NONTW_03_2,OP1,SCAN1,HALT1,COMP1,HLV1,STSET1"
            ]);
            PatternListSingleton instance = PatternListSingleton.GetInstance(path, "");
            var latestPatternDict = new Dictionary<string, string> { ["PAT_NONTW_03"] = "2" };

            // Act
            List<string>? result = instance.UpdateCsvByLatestPattern(latestPatternDict);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Count);
        }
    }
}
