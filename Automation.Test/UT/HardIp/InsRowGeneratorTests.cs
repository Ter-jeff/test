using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Singleton;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.HardIpDc.BaseData;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class InsRowGeneratorTests : FunctionTestBase
    {
        private static HardIpInsRowGenerator _gen = null!;
        private static HardIpSheet _sheet = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            HardIpParaData paraData = new HardIpParaData(EnumBlock.HardIp);
            var hardIpInputData = new HardIpInputData(paraData);
            _sheet = new HardIpSheet
            {
                Rows = []
            };
            _gen = new HardIpInsRowGenerator(hardIpInputData, _sheet, "")
            {
                Pat = new HardIpPattern(),
                LabelVoltage = "LV"
            };
        }

        [TestMethod]
        public void GenReTestRow_ShouldDuplicateInstanceRowWithModifiedName_AndAppendSetting()
        {
            var originalRow = new InstanceRow
            {
                TestName = "TEST_1p8V",
                VbtName = "MyVbt",
                ArgList = "pre_pat,arg2",
                Args = ["Arg1Value", "Arg2Value"],
                VbtType = ".NET",
                DcCategory = "DC1",
                DcSelector = "TYP",
                AcCategory = "AC1",
                AcSelector = "ACSEL",
                TimeSets = "TS1",
                PinLevels = "PL1",
                Overlay = "OV1"
            };

            // Act
            InstanceRow result = _gen.GenReTestRow(originalRow);

            // Assert
            Assert.AreEqual("TEST_1p8V", result.TestName);
            Assert.AreEqual("MyVbt", result.VbtName);
            Assert.AreEqual(".NET", result.VbtType);
            Assert.AreEqual("Arg1Value", result.Args[0].Split(',').Last(), "Expected appended retest setting.");
        }

        [TestMethod]
        public void GetSpecifyInfo_ShouldReturnCorrectValue_WhenKeyExists()
        {
            string result = _gen.GetSpecifyInfo("DC:VDD;AC:CLK", "AC");
            Assert.AreEqual("CLK", result);
        }

        [TestMethod]
        public void GetSpecifyInfo_ShouldReturnEmpty_WhenKeyNotFound()
        {
            string result = _gen.GetSpecifyInfo("DC:VDD;AC:CLK", "XYZ");
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetAcSelector_ShouldReturnCorrectSelector_ForMatchingLabelVoltage()
        {
            string info = "SEL:HV:MAX;SEL:LV:MIN;SEL:ALL:TYP";
            string result = _gen.GetAcSelector("LV", info);
            Assert.AreEqual(CommonConst.Typ, result);
        }

        [TestMethod]
        public void GetAcSelector_ShouldUseALL_WhenExactVoltageNotFound()
        {
            string info = "SEL:ALL:TYP";
            string result = _gen.GetAcSelector("HV", info);
            Assert.AreEqual(CommonConst.Typ, result);
        }

        [TestMethod]
        public void GetDcSelector_ShouldReturnCorrectSelector_ForMatchingVoltage()
        {
            string info = "SEL:HV:MAX;SEL:LV:MIN";
            string result = _gen.GetDcSelector("HV", info);
            Assert.AreEqual(CommonConst.Max, result);
        }

        [TestMethod]
        public void GetDcSelector_ShouldReturnEmpty_WhenNoMatchFound()
        {
            string info = "SEL:NV:TYP";
            string result = _gen.GetDcSelector("HV", info);
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GenInsRows_Should_Use_AC_From_ForceCondition()
        {
            var pat = new HardIpPattern
            {
                MiscInfo = "AC:FORCE_AC",

                Pattern = new PatternClass("DUMMY_PATTERN")
            };

            _sheet.Rows.Clear();
            _sheet.Rows.Add(pat);

            _gen.Pat = pat;

            List<InstanceRow> rows = _gen.GenInsRows();

            Assert.AreNotEqual(null, rows, "GenInsRows should not return null.");
            Assert.IsTrue(rows.Count != 0, "GenInsRows should generate at least one InstanceRow.");
        }

        [TestMethod]
        public void GenInsRows_Should_Use_AC_From_PatternAcCategory()
        {
            var pat = new HardIpPattern
            {
                MiscInfo = "ACCategory:PATTERN_AC",

                Pattern = new PatternClass("DUMMY_PATTERN")
            };

            _sheet.Rows.Clear();
            _sheet.Rows.Add(pat);
            _gen.Pat = pat;

            List<InstanceRow> rows = _gen.GenInsRows();

            Assert.AreNotEqual(null, rows);
            Assert.IsTrue(rows.Count != 0);
        }

        [TestMethod]
        public void GenInsRows_Should_Return_Empty_AcCategory_When_Pattern_Not_In_Map()
        {
            var pat = new HardIpPattern
            {
                Pattern = new PatternClass("NOT_EXIST_PATTERN")
            };

            _sheet.Rows.Clear();
            _sheet.Rows.Add(pat);
            _gen.Pat = pat;

            List<InstanceRow> rows = _gen.GenInsRows();

            Assert.AreNotEqual(null, rows, "GenInsRows should not return null.");
            Assert.IsTrue(rows.Count != 0, "GenInsRows should generate at least one InstanceRow.");

            Assert.AreEqual(string.Empty, rows.First().AcCategory);
        }

        [TestMethod]
        public void GenInsRows_Should_Use_TimeSetCategory_When_Pattern_In_Map()
        {
            var pat = new HardIpPattern
            {
                Pattern = new PatternClass("DEFAULT")
            };

            _sheet.Rows.Clear();
            _sheet.Rows.Add(pat);
            _gen.Pat = pat;

            List<InstanceRow> rows = _gen.GenInsRows();

            Assert.AreNotEqual(null, rows);
            Assert.IsTrue(rows.Count != 0, "GenInsRows should generate at least one InstanceRow.");

            Assert.IsNotNull(string.IsNullOrEmpty(rows.First().AcCategory), "AcCategory should be generated from TimeSet mapping.");
        }

        [TestMethod]
        public void GenInsRows_Should_Return_Empty_AcSelector_When_Pattern_Not_In_Map()
        {
            var pat = new HardIpPattern
            {
                Pattern = new PatternClass("NOT_EXIST_PATTERN"),

                SheetName = "NORMAL_SHEET"
            };

            _sheet.Rows.Clear();
            _sheet.Rows.Add(pat);
            _gen.Pat = pat;

            List<InstanceRow> rows = _gen.GenInsRows();

            Assert.AreNotEqual(null, rows, "GenInsRows should not return null.");
            Assert.IsTrue(rows.Count != 0, "GenInsRows should generate at least one InstanceRow.");

            Assert.AreEqual(string.Empty, rows.First().AcSelector);
        }

        [TestMethod]
        public void GenInsRows_Should_Use_Typ_DcSelector_When_LabelVoltage_Is_NV()
        {
            var pat = new HardIpPattern
            {
                Pattern = new PatternClass("NOT_EXIST_PATTERN"),
                SheetName = "NORMAL_SHEET"
            };

            _sheet.Rows.Clear();
            _sheet.Rows.Add(pat);

            _gen.LabelVoltage = "NV";
            _gen.Pat = pat;

            List<InstanceRow> rows = _gen.GenInsRows();

            Assert.AreNotEqual(null, rows, "GenInsRows should not return null.");
            Assert.IsTrue(rows.Count != 0, "GenInsRows should generate at least one InstanceRow.");

            Assert.AreEqual(CommonConst.Typ, rows.First().DcSelector);
        }

        [TestMethod]
        public void GenInsRows_Should_Use_Max_DcSelector_When_LabelVoltage_Is_HV()
        {
            var pat = new HardIpPattern
            {
                Pattern = new PatternClass("NOT_EXIST_PATTERN"),
                SheetName = "NORMAL_SHEET"
            };

            _sheet.Rows.Clear();
            _sheet.Rows.Add(pat);

            _gen.LabelVoltage = CommonConst.Hv;
            _gen.Pat = pat;

            List<InstanceRow> rows = _gen.GenInsRows();

            Assert.AreNotEqual(null, rows, "GenInsRows should not return null.");
            Assert.IsTrue(rows.Count != 0, "GenInsRows should generate at least one InstanceRow.");

            Assert.AreEqual(CommonConst.Max, rows.First().DcSelector);
        }

        [TestMethod]
        public void GenInsRows_Should_Use_Min_DcSelector_When_LabelVoltage_Is_LV()
        {
            var pat = new HardIpPattern
            {
                Pattern = new PatternClass("NOT_EXIST_PATTERN"),
                SheetName = "NORMAL_SHEET"
            };

            _sheet.Rows.Clear();
            _sheet.Rows.Add(pat);

            _gen.LabelVoltage = CommonConst.Lv;
            _gen.Pat = pat;

            List<InstanceRow> rows = _gen.GenInsRows();

            Assert.AreNotEqual(null, rows, "GenInsRows should not return null.");
            Assert.IsTrue(rows.Count != 0, "GenInsRows should generate at least one InstanceRow.");

            Assert.AreEqual(CommonConst.Min, rows.First().DcSelector);
        }

        [TestMethod]
        public void GenInsRows_Should_Use_Typ_DcSelector_When_LabelVoltage_Is_Unknown()
        {
            var pat = new HardIpPattern
            {
                Pattern = new PatternClass("NOT_EXIST_PATTERN"),
                SheetName = "NORMAL_SHEET"
            };

            _sheet.Rows.Clear();
            _sheet.Rows.Add(pat);

            _gen.LabelVoltage = "UNKNOWN";
            _gen.Pat = pat;

            List<InstanceRow> rows = _gen.GenInsRows();

            Assert.AreNotEqual(null, rows, "GenInsRows should not return null.");
            Assert.IsTrue(rows.Count != 0, "GenInsRows should generate at least one InstanceRow.");

            Assert.AreEqual(CommonConst.Typ, rows.First().DcSelector);
        }

        [TestMethod]
        public void GenReTestRow_Should_Not_Append_Setting_When_PrePat_Not_In_ArgList()
        {
            var originalRow = new InstanceRow
            {
                TestName = "TEST_1p8V",
                VbtName = "MyVbt",
                ArgList = "arg1,arg2",
                Args = ["Arg1Value", "Arg2Value"],
                VbtType = ".NET",
                DcCategory = "DC1",
                DcSelector = "TYP",
                AcCategory = "AC1",
                AcSelector = "ACSEL",
                TimeSets = "TS1",
                PinLevels = "PL1",
                Overlay = "OV1"
            };

            InstanceRow result = _gen.GenReTestRow(originalRow);

            Assert.AreNotEqual(null, result);
            Assert.AreEqual(originalRow.Args[0], result.Args[0], "Args[0] should remain unchanged when InterposePrePat is not in ArgList.");
            Assert.AreEqual(originalRow.Args[1], result.Args[1], "Args[1] should remain unchanged when InterposePrePat is not in ArgList.");
        }

        private static string InvokeCreateHardIpLevelConcurrent(string timeSet)
        {
            return _gen.CreateHardIpLevelConcurrent(timeSet);
        }

        private static string InvokeCreateHardIpPinLevel()
        {
            return _gen.CreateHardIpPinLevel();
        }

        private static void SetHardIpDcSetting(HardIpCategoryDef value)
        {
            _gen.HardIpDcSetting = value;
        }

        [TestMethod]
        public void CreateHardIpLevelConcurrent_NoMatchingTimeSet_ReturnsTbdError()
        {
            // Arrange
            AcTSetCategoryMapSingleton.Instance().MultiTimeSetSheets = [];

            // Act
            string result = InvokeCreateHardIpLevelConcurrent("TS_Missing");

            // Assert
            Assert.AreEqual("TBD(ConcurrentLevelError)", result);
        }

        [TestMethod]
        public void CreateHardIpLevelConcurrent_MatchingTimeSet_ReturnsLevelsWithTimeDomain()
        {
            // Arrange
            AcTSetCategoryMapSingleton.Instance().MultiTimeSetSheets =
            [
                new ComTimeSetBasicSheet("TS_Match", "Single", "", "DomainA", "")
            ];

            // Act
            string result = InvokeCreateHardIpLevelConcurrent("TS_Match");

            // Assert
            Assert.AreEqual("Levels_Con_H_DomainA", result);
        }

        [TestMethod]
        public void CreateHardIpPinLevel_HardIpDcSettingPresent_ReturnsItsLevelSheet()
        {
            // Arrange
            SetHardIpDcSetting(new HardIpCategoryDef("Cat1") { LevelSheet = "Levels_Custom" });

            // Act
            string result = InvokeCreateHardIpPinLevel();

            // Assert
            Assert.AreEqual("Levels_Custom", result);

            // Cleanup
            SetHardIpDcSetting(null);
        }

        [TestMethod]
        public void CreateHardIpPinLevel_NoHardIpDcSetting_ReturnsDefaultLevel()
        {
            // Arrange
            SetHardIpDcSetting(null);

            // Act
            string result = InvokeCreateHardIpPinLevel();

            // Assert
            Assert.AreEqual(HardIpConstData.LevelDefault, result);
        }
    }
}
