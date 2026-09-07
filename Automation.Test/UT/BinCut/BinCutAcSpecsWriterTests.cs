using System.Collections.Generic;

using Automation.GenerateIgxl.BinCut.Business;
using Automation.Singleton;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Basic;

namespace Automation.Test.UT.BinCut
{
    [TestClass]
    public class BinCutAcSpecsWriterTests
    {
        private BinCutAcSpecsWriter _writer = null!;

        [TestInitialize]
        public void Setup()
        {
            _writer = new BinCutAcSpecsWriter();
            AcTSetCategoryMapSingleton.Instance().MultiTimeSetSheets = [];
        }

        #region ModifyValueFromTimeSettings

        [TestMethod]
        public void ModifyValueFromTimeSettings_ValueProvidedUsingTSetSymbol_ReturnsCategoryWithOnlyTyp()
        {
            // Arrange
            var acSpec = new AcSpec("Using TSet", null);

            // Act
            CategoryInSpec result = _writer.ModifyValueFromTimeSettings("Orig", "NewCat", "1.5", acSpec);

            // Assert
            Assert.AreEqual("NewCat", result.Name);
            Assert.AreEqual("1.5", result.Typ);
            Assert.AreEqual("", result.Min);
            Assert.AreEqual("", result.Max);
        }

        [TestMethod]
        public void ModifyValueFromTimeSettings_ValueProvidedNonTSetSymbol_ReturnsCategoryWithAllFieldsEqual()
        {
            // Arrange
            var acSpec = new AcSpec("VccFreq", null);

            // Act
            CategoryInSpec result = _writer.ModifyValueFromTimeSettings("Orig", "NewCat", "2.0", acSpec);

            // Assert
            Assert.AreEqual("2.0", result.Typ);
            Assert.AreEqual("2.0", result.Min);
            Assert.AreEqual("2.0", result.Max);
        }

        [TestMethod]
        public void ModifyValueFromTimeSettings_NoValueOriginalFound_CopiesOriginalCategoryValues()
        {
            // Arrange
            var acSpec = new AcSpec("VccFreq", null);
            acSpec.AddCategory(new CategoryInSpec("Orig", "1.0", "0.9", "1.1"));

            // Act
            CategoryInSpec result = _writer.ModifyValueFromTimeSettings("Orig", "NewCat", "", acSpec);

            // Assert
            Assert.AreEqual("NewCat", result.Name);
            Assert.AreEqual("1.0", result.Typ);
            Assert.AreEqual("0.9", result.Min);
            Assert.AreEqual("1.1", result.Max);
        }

        [TestMethod]
        public void ModifyValueFromTimeSettings_NoValueOriginalNotFound_FallsBackToFirstCategory()
        {
            // Arrange
            var acSpec = new AcSpec("VccFreq", null);
            acSpec.AddCategory(new CategoryInSpec("SomeOtherCat", "3.0", "2.9", "3.1"));

            // Act
            CategoryInSpec result = _writer.ModifyValueFromTimeSettings("NotThere", "NewCat", "", acSpec);

            // Assert
            Assert.AreEqual("NewCat", result.Name);
            Assert.AreEqual("3.0", result.Typ);
        }

        #endregion

        #region UpdateSymbolByTimeSet

        [TestMethod]
        public void UpdateSymbolByTimeSet_NoMatchingTimeSetVersion_ReturnsEmpty()
        {
            // Arrange
            AcTSetCategoryMapSingleton.Instance().MultiTimeSetSheets = [];
            var patternDatas = new Dictionary<string, PatternData>();

            // Act
            string result = _writer.UpdateSymbolByTimeSet(patternDatas, "TS1", []);

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void UpdateSymbolByTimeSet_PatternNotInDictionary_SkippedReturnsEmpty()
        {
            // Arrange
            AcTSetCategoryMapSingleton.Instance().MultiTimeSetSheets = [new ComTimeSetBasicSheet("TS1")];
            var patternDatas = new Dictionary<string, PatternData>();

            // Act
            string result = _writer.UpdateSymbolByTimeSet(patternDatas, "TS1", ["PatA"]);

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void UpdateSymbolByTimeSet_MatchingTSetRowFound_ReturnsTrimmedLastSegment()
        {
            // Arrange
            var tsetSheet = new ComTimeSetBasicSheet("TS1");
            tsetSheet.Rows.Add(new TSet { Name = "ScanTs1", CyclePeriod = "=1/(_Freq_VAR)" });
            AcTSetCategoryMapSingleton.Instance().MultiTimeSetSheets = [tsetSheet];
            var patternDatas = new Dictionary<string, PatternData> { { "pata", new PatternData { ScanTset = "ScanTs1" } } };

            // Act
            string result = _writer.UpdateSymbolByTimeSet(patternDatas, "TS1", ["PatA"]);

            // Assert
            Assert.AreEqual("Freq_VAR", result);
        }

        [TestMethod]
        public void UpdateSymbolByTimeSet_ScanTsetEmpty_SkippedReturnsEmpty()
        {
            // Arrange
            AcTSetCategoryMapSingleton.Instance().MultiTimeSetSheets = [new ComTimeSetBasicSheet("TS1")];
            var patternDatas = new Dictionary<string, PatternData> { { "pata", new PatternData { ScanTset = "" } } };

            // Act
            string result = _writer.UpdateSymbolByTimeSet(patternDatas, "TS1", ["PatA"]);

            // Assert
            Assert.AreEqual("", result);
        }

        #endregion

        #region UpdateAcSpecByShift

        [TestMethod]
        public void UpdateAcSpecByShift_NoUserDefinedAcNoShiftFreq_ReturnsOriginalAcUnchanged()
        {
            // Act
            string result = _writer.UpdateAcSpecByShift("OrigAc", "TS1", "", "", "");

            // Assert
            Assert.AreEqual("OrigAc", result);
        }

        [TestMethod]
        public void UpdateAcSpecByShift_UserDefinedAcButNoTimeSettingSheet_ReturnsEmpty()
        {
            // Arrange
            _writer._timeSettingSheet = null;

            // Act
            string result = _writer.UpdateAcSpecByShift("OrigAc", "TS1", "", "", "UserAc");

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void UpdateAcSpecByShift_UserDefinedAcNotFoundInTimeSettingSheet_ReturnsMissingMessage()
        {
            // Arrange
            var timeSettingSheet = new TimeSettingSheet("TimeSetting");
            _writer._timeSettingSheet = timeSettingSheet;

            // Act
            string result = _writer.UpdateAcSpecByShift("OrigAc", "TS1", "", "", "UserAc");

            // Assert
            Assert.AreEqual("Missing user define AC", result);
        }

        [TestMethod]
        public void UpdateAcSpecByShift_UserDefinedAcAlreadySetup_SkipsSetupReturnsUserDefinedAc()
        {
            // Arrange - _alreadySetupAc already contains "UserAc", so the whole setup block is skipped
            var timeSettingSheet = new TimeSettingSheet("TimeSetting");
            _writer._timeSettingSheet = timeSettingSheet;
            HashSet<string> alreadySetup = _writer._alreadySetupAc;
            alreadySetup.Add("UserAc");

            // Act
            string result = _writer.UpdateAcSpecByShift("OrigAc", "TS1", "", "", "UserAc");

            // Assert
            Assert.AreEqual("UserAc", result);
        }

        [TestMethod]
        public void UpdateAcSpecByShift_UserDefinedAcNewCategorySetup_AddsCategoryAndMarksSetup()
        {
            // Arrange
            var acSpecSheet = new AcSpecSheet("AC_Spec");
            var acSpec = new AcSpec("VccFreq", null);
            acSpec.AddCategory(new CategoryInSpec("OrigAc", "1.0", "0.9", "1.1"));
            acSpecSheet.Rows.Add(acSpec);
            _writer._acSpecSheet = acSpecSheet;

            var timeSettingSheet = new TimeSettingSheet("TimeSetting");
            var tsRow = new TimeSettingRow("UserAc");
            tsRow.SymbolValues["VccFreq"] = "1.2";
            timeSettingSheet.Rows.Add(tsRow);
            _writer._timeSettingSheet = timeSettingSheet;

            // Act
            string result = _writer.UpdateAcSpecByShift("OrigAc", "TS1", "", "", "UserAc");

            // Assert
            Assert.AreEqual("UserAc", result);
            Assert.IsTrue(acSpecSheet.CategoryList.Contains("UserAc"));
            HashSet<string> alreadySetup = _writer._alreadySetupAc;
            Assert.IsTrue(alreadySetup.Contains("UserAc"));
        }

        [TestMethod]
        public void UpdateAcSpecByShift_ShiftFreqWithSymbol_BuildsShiftAcNameAndMarksSetup()
        {
            // Arrange
            var acSpecSheet = new AcSpecSheet("AC_Spec");
            var acSpec = new AcSpec("FreqSym", null);
            acSpec.AddCategory(new CategoryInSpec("OrigAc", "1.0", "0.9", "1.1"));
            acSpecSheet.Rows.Add(acSpec);
            _writer._acSpecSheet = acSpecSheet;

            // Act - no userDefinedAc, so newAcName stays "OrigAc"; shiftFreq builds "OrigAc_100"
            string result = _writer.UpdateAcSpecByShift("OrigAc", "TS1", "100", "FreqSym", "");

            // Assert
            Assert.AreEqual("OrigAc_100", result);
            Assert.IsTrue(acSpecSheet.CategoryList.Contains("OrigAc_100"));
        }

        #endregion
    }
}
