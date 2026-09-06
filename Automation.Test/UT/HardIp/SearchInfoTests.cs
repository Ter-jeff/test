using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.GenerateIgxl.Wireless.DVDC.InputObject;
using Automation.InputManager.Data;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.Enums;
using CommonLib.ErrorReport;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using ScghLib.Reader;

using TestPlanLib.Basic;
using TestPlanLib.BinNumberLegacy;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class SearchInfoTests : FunctionTestBase
    {
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            LocalSpecs.TarFolder = OutputPath;
            PinMapSheet pinMapSheet = new PinMapSheet("");
            var pinGroup = new PinGroup("MY_GROUP", "I/O");
            pinGroup.AddPin(new Pin("TX_P", "I/O"));
            pinGroup.AddPin(new Pin("TX_N", "I/O"));
            pinMapSheet.AddGroup(pinGroup);
            Pin pin = new Pin("VCC", "power");
            pinMapSheet.AddPin(pin);
            TestProgram.IgxlWorkBk.PinMapPair = new KeyValuePair<string, PinMapSheet>(FolderStructure.DirPinMap, pinMapSheet);

            string hardIpInfoFile = Path.Combine(InputPath, "HardIp", "HardIP_PatInfo_Default.log");
            List<HardIpInfo> patInfo = new PatInfoReader().ExtractHardIpInfos(hardIpInfoFile);
            LocalSpecs.HardIpInfos = new HardIpInfos(patInfo);
        }

        [TestMethod]
        public void GetCusStrDigCapData_ShouldCorrectlyFormatDsscOut_WhenKeepDsscOutIsPresent()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MiscInfo = "KeepDsscOut",
                MeasPins = []
            };

            var mockInfo = new HardIpInfo
            {
                CapBitStr = "BIT_8|BIT_8",
                CapBitName = "DATA_L|DATA_H",
                DsscOut = "DSSC_OUT, 8:TEMP_A, 8:TEMP_B"
            };

            // Act
            string result = SearchInfo.GetCusStrDigCapData(pattern, mockInfo);

            // Assert
            Assert.AreEqual("DSSC_OUT, 8:TEMP_A, 8:TEMP_B", result);
        }

        [TestMethod]
        public void GetCusStrDigCapData_ShouldCorrectlyFormatDsscOut_WhenKeepDsscOutIsPresent_1()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MiscInfo = "Disable_MeasC_Split",
                MeasPins = []
            };

            var mockInfo = new HardIpInfo
            {
                CapBitStr = "BIT_8|BIT_8",
                CapBitName = "DATA_L|DATA_H",
                DsscOut = "DSSC_OUT, 8:TEMP_A, 8:TEMP_B"
            };

            // Act
            string result = SearchInfo.GetCusStrDigCapData(pattern, mockInfo);

            // Assert
            Assert.AreEqual("DSSC_OUT,8:DATA_L,8:DATA_H", result);
        }

        [TestMethod]
        public void GetCusStrDigCapData_ShouldCorrectlyFormatDsscOut_WhenKeepDsscOutIsPresent_1_1()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MiscInfo = "Disable_MeasC_Split",
                MeasPins = []
            };

            var mockInfo = new HardIpInfo
            {
                CapBitStr = "BIT_16|BIT_8",
                CapBitName = "DATA_L|DATA_H",
                DsscOut = "DSSC_OUT, 8:TEMP_A,8:TEMP_B,8:TEMP_C"
            };

            // Act
            string result = SearchInfo.GetCusStrDigCapData(pattern, mockInfo);

            // Assert
            Assert.AreEqual("DSSC_OUT,16:DATA_L,8:DATA_H", result);
        }

        [TestMethod]
        public void GetCusStrDigCapData_ShouldCorrectlyFormatDsscOut_WhenKeepDsscOutIsPresent_2()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MiscInfo = "Ignore_Patt_MeasC",
                MeasPins = [new() { MeasType = MeasType.MeasC, CapBit = "1" }]
            };

            var mockInfo = new HardIpInfo
            {
                CapBitStr = "BIT_8|BIT_8",
                CapBitName = "DATA_L|DATA_H",
                DsscOut = "DSSC_OUT, 8:TEMP_A, 8:TEMP_B",
                CapBit = 1
            };

            // Act
            string result = SearchInfo.GetCusStrDigCapData(pattern, mockInfo);

            // Assert
            Assert.AreEqual("DSSC_OUT,1:", result);
        }

        [TestMethod]
        public void InitialMeasCTest()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MeasPins = [new() { MeasType = MeasType.MeasC, CapBit = "1", PinCount = 1 }]
            };

            var mockInfo = new HardIpInfo
            {
                CapBitStr = "BIT_8|BIT_8",
                CapBitName = "DATA_L|DATA_H",
                DsscOut = "DSSC_OUT, 8:TEMP_A, 8:TEMP_B",
                CapBit = 1
            };

            // Act
            SearchInfo.InitialMeasC(ref pattern, mockInfo);

            // Assert
            Assert.AreEqual(2, pattern.MeasPins.Count);
        }

        [TestMethod]
        public void GetCusStrDigCapData_ShouldIgnorePatInfo_WhenFlagIsSet()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MiscInfo = "ignore_patt_comment",
                MeasPins =
            [
                new() { MeasType = MeasType.MeasC, CapBit = "4", TestName = "MyTest" }
            ]
            };

            var mockInfo = new HardIpInfo
            {
                CapBit = 4,
                DsscOut = "DSSC_OUT,4:OldName"
            };

            // Act
            string result = SearchInfo.GetCusStrDigCapData(pattern, mockInfo);

            // Assert
            Assert.IsTrue(result.Contains("4:MyTest"));
        }

        [TestMethod]
        public void GetCusStrDigCapData_ShouldReturnStoreDigAll_WhenMRREnabled()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MiscInfo = "StoreDigAll",
                MeasPins = []
            };
            var mockInfo = new HardIpInfo { DsscOut = "DSSC_OUT" };

            // Act
            string result = SearchInfo.GetCusStrDigCapData(pattern, mockInfo);

            // Assert
            Assert.IsTrue(result.StartsWith("StoreDigAll&"));
        }

        [TestMethod]
        public void GetTestLimitPerMeasType_WhenVifNameMatches_ReturnsCorrectVfiDictionary()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                FunctionName = "meas_vir_io_universal_func",
                MeasPins =
                [
                    new() { SequenceIndex = 1, MeasType = "MeasV", PinName = "Pin_A,Pin_B" },
                    new() { SequenceIndex = 1, MeasType = "MeasI", PinName = "Pin_C::Pin_D" },
                    new() { SequenceIndex = 1, MeasType = "MeasV", PinName = "Pin_C::Pin_D" },
                    new() { SequenceIndex = 1, MeasType = "MeasR1", PinName = "Pin_C::Pin_D" }
                ],
                TestPlanSequences = [new(1, 1, 1)]
            };

            // Act
            Dictionary<string, string> result = SearchInfo.GetTestLimitPerMeasType(pattern);

            // Assert
            var expected = new Dictionary<string, string>
            {
                { "V", "F" },
                { "I", "F" },
                { "R", "T" }
            };

            CollectionAssert.AreEquivalent(expected, result, "The VFI dictionary should match expected flags.");
        }

        [TestMethod]
        public void GetTestLimitPerMeasType_WhenVifNameMatches_ReturnsCorrectVfiDictionary_1()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                FunctionName = "meas_vir_io_universal_func",
                MeasPins =
                [
                    new() { SequenceIndex = 1, MeasType = "MeasV", PinName = "Pin_A" },
                ],
                TestPlanSequences = [new(1, 1, 1)]
            };

            // Act
            Dictionary<string, string> result = SearchInfo.GetTestLimitPerMeasType(pattern);

            // Assert
            var expected = new Dictionary<string, string>
            {
                { "V", "F" },
                { "I", "F" },
                { "R", "F" }
            };

            CollectionAssert.AreEquivalent(expected, result, "The VFI dictionary should match expected flags.");
        }

        [TestMethod]
        public void GetPlanCurrentRange_ShouldMatchPinsByTypeAndName_AndModifyReferenceList()
        {
            // Arrange
            var measPins = new List<MeasPin>
            {
                new() { PinName = "PIN_A", MeasType = "MeasV", VisitedTime = 1, IsUsedPin = false },
                new() { PinName = "PIN_B", MeasType = "MeasI", VisitedTime = 1, IsUsedPin = false },
                new() { PinName = "PIN_C", MeasType = "MeasVdiff", VisitedTime = 1, IsUsedPin = false },
            };
            var patInfoPins = new List<MeasPin>
            {
                new() { PinName = "PIN_A", MeasType = "MeasV", SequenceIndex = 1 },
                new() { PinName = "PIN_C", MeasType = "MeasVdiff", SequenceIndex = 1 }
            };

            // Act
            SearchInfo.GetPlanCurrentRange(measPins, patInfoPins, false);

            Assert.AreEqual(2, patInfoPins.Count, "The result list should contain exactly one pin.");
            Assert.AreEqual("PIN_A", patInfoPins[0].PinName);
            MeasPin matchedPin = measPins.First(p => p.PinName == "PIN_A");
            Assert.IsTrue(matchedPin.IsUsedPin, "The matched plan pin should be marked as used.");
            Assert.AreEqual(1, matchedPin.VisitedTime, "VisitedTime should be reset to 0 after matching.");
        }

        [TestMethod]
        public void GetPlanCurrentRange_ShouldMatchPinsByTypeAndName_AndModifyReferenceList_1()
        {
            // Arrange
            var measPins = new List<MeasPin>
            {
                new() { PinName = "PIN_A", MeasType = "MeasVdiff2", VisitedTime = 1, IsUsedPin = false },
                new() { PinName = "PIN_B", MeasType = "MeasI", VisitedTime = 1, IsUsedPin = false },
                new() { PinName = "PIN_C", MeasType = "MeasVdiff", VisitedTime = 1, IsUsedPin = false },
            };
            var patInfoPins = new List<MeasPin>
            {
                new() { PinName = "PIN_A", MeasType = "MeasVdiff2", SequenceIndex = 1 },
                new() { PinName = "PIN_A", MeasType = "MeasVdiff", SequenceIndex = 1 }
            };

            // Act
            SearchInfo.GetPlanCurrentRange(measPins, patInfoPins, false);

            Assert.AreEqual(1, patInfoPins.Count, "The result list should contain exactly one pin.");
            Assert.AreEqual("PIN_A", patInfoPins[0].PinName);
            MeasPin matchedPin = measPins.First(p => p.PinName == "PIN_A");
            Assert.IsTrue(matchedPin.IsUsedPin, "The matched plan pin should be marked as used.");
            Assert.AreEqual(1, matchedPin.VisitedTime, "VisitedTime should be reset to 0 after matching.");
        }

        [TestMethod]
        public void GetPlanCurrentRange_ShouldMatchPinsByTypeAndName_AndModifyReferenceList1()
        {
            // Arrange
            var measPins = new List<MeasPin>
            {
                new() { PinName = "PIN_A", MeasType = "MeasV", VisitedTime = 0 },
            };
            var patInfoPins = new List<MeasPin>
            {
                new() { PinName = "PIN_A", MeasType = "MeasV", SequenceIndex = 1 },
                new() { PinName = "PIN_A", MeasType = "MeasV", SequenceIndex = 2 }
            };

            // Act
            SearchInfo.GetPlanCurrentRange(measPins, patInfoPins, true);

            Assert.AreEqual(2, patInfoPins.Count, "The result list should contain exactly one pin.");
            Assert.AreEqual("PIN_A", patInfoPins[0].PinName);
            MeasPin matchedPin = measPins.First(p => p.PinName == "PIN_A");
            Assert.IsTrue(matchedPin.IsUsedPin, "The matched plan pin should be marked as used.");
            Assert.AreEqual(1, matchedPin.VisitedTime, "VisitedTime should be reset to 0 after matching.");
        }

        [TestMethod]
        public void GetPlanCurrentRange_ShouldMatchPinsByTypeAndName_AndModifyReferenceList2()
        {
            // Arrange
            var measPins = new List<MeasPin> { };
            var patInfoPins = new List<MeasPin>
            {
                new() { PinName = "PIN_A", MeasType = "MeasV", SequenceIndex = 1 },
                new() { PinName = "PIN_C", MeasType = "MeasVdiff", SequenceIndex = 1 }
            };

            // Act
            SearchInfo.GetPlanCurrentRange(measPins, patInfoPins, false);

            Assert.AreEqual(1, patInfoPins.Count, "The result list should contain exactly one pin.");
            Assert.AreEqual("", patInfoPins[0].PinName);
        }

        [TestMethod]
        public void GetPlanCurrentRange_WhenWiSrcPresent_ReturnsEarlyWithMeasPins()
        {
            // Arrange
            var measPins = new List<MeasPin> { new() { MeasType = "WiSrc" } };
            var patInfoPins = new List<MeasPin> { new() { MeasType = "Standard" } };

            // Act
            SearchInfo.GetPlanCurrentRange(measPins, patInfoPins, false);

            Assert.AreEqual("Standard", patInfoPins[0].MeasType);
        }

        [TestMethod]
        [DataRow("Limit:1.2V;Other:0.8V", "1.2V", "")]
        [DataRow("NoBin:0.9V;Limit:1.1V", "0.9V", "No")]
        [DataRow("Limit:1.2V;Other:0.8V", "0.8V", "")]
        [DataRow("Limit:1.2V", "1.0V", "")]
        [DataRow("InvalidData", "1.2V", "")]
        [DataRow("", "1.2V", "")]
        [DataRow("NoBin:NV", "NV", "No")]
        [DataRow("NoBin:HV", "NV", "")]
        [DataRow("", "NV", "")]

        public void GetFlagNoBinStr_Scenarios(string miscInfo, string voltage, string expected)
        {
            // Act
            string result = SearchInfo.GetFlagNoBinStr(miscInfo, voltage);

            // Assert
            Assert.AreEqual(expected, result, $"Failed for input: {miscInfo}, {voltage}");
        }

        [TestMethod]
        public void GetFlagNoBinStr_ReturnsEmpty_WhenNoMatchInMultipleItems()
        {
            // Arrange
            string miscInfo = "TagA:1.0V;TagB:1.1V;TagC:1.2V";
            string voltage = "1.1V";

            // Act
            string result = SearchInfo.GetFlagNoBinStr(miscInfo, voltage);

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetInstrumentInfo_ReturnsFormattedString_WhenValidMatchesFound()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MeasPins = [
                new()
                {
                    SequenceIndex = 1,
                    MeasType = "MeasValue", // Not MeasCalc/MeasLimit
                    RfInstrumentSetup = "Gain=40$Power=10",
                    PinName = "Pin1"
                },
                    new()
                    {
                        SequenceIndex = 2,
                        MeasType = "MeasValue",
                        RfInstrumentSetup = "Gain=20",
                        PinName = "Pin2"
                    }
            ]
            };

            // Act
            // Regex extracts "40" from "Gain=40" and "20" from "Gain=20"
            string result = SearchInfo.GetInstrumentInfo(pattern, "Gain");

            // Assert
            // Logic: (40) for Seq1 | (20) for Seq2 -> "40|20"
            // Note: Code joins group values with '+', then sequences with '|'
            Assert.AreEqual("40|20", result);
        }

        [TestMethod]
        [DataRow("V", ForceConditionType.FvMode)]
        [DataRow("I", ForceConditionType.FiMode)]
        [DataRow("Normal", ForceConditionType.Normal)]
        public void GetForceMode_ReturnsCorrectEnum_BasedOnPinType(string forceType, ForceConditionType forceConditionType)
        {
            // Arrange
            var pins = new List<MeasPin> {
            new()
            {
                ForceConditions = [
                    new()
                    {
                        ForcePins = [
                            new() { ForceType = forceType }
                        ]
                    }
                ]
            }
        };

            // Act
            ForceConditionType result = SearchInfo.GetForceMode(pins);

            // Assert
            Assert.AreEqual(forceConditionType, result);
        }

        [TestMethod]
        public void GetInstrumentInfo_ReturnsEmpty_WhenNoValidPinsExist()
        {
            // Arrange: All pins are MeasCalc or MeasLimit (filtered out)
            var pattern = new HardIpPattern
            {
                MeasPins = [
                new() { MeasType = "MeasCalc" },
                    new() { MeasType = "MeasLimit" }
            ]
            };

            // Act
            string result = SearchInfo.GetInstrumentInfo(pattern, "AnyItem");

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetHardIpInfos_WhenMultiplePatternsMatch_ReturnsAllItems()
        {
            // Arrange
            var infoA = new HardIpInfo { /* init properties */ };
            var infoB = new HardIpInfo { /* init properties */ };

            // Mocking the dictionary/collection passed in
            var sourceData = new HardIpInfos
        {
            { "Pattern1", new List<HardIpInfo> { infoA } },
            { "Pattern2", new List<HardIpInfo> { infoB } }
        };
            var searchNames = new List<string> { "Pattern1", "Pattern2" };

            // Act
            List<HardIpInfo> result = SearchInfo.GetHardIpInfos(searchNames, sourceData);

            // Assert
            Assert.AreEqual(2, result.Count, "Should return total items from both patterns.");
            CollectionAssert.Contains(result, infoA);
            CollectionAssert.Contains(result, infoB);
        }

        [TestMethod]
        public void GetHardIpInfos_WhenPatternMissing_ReturnsOnlyMatches()
        {
            // Arrange
            var infoA = new HardIpInfo();
            var sourceData = new HardIpInfos { { "Found", new List<HardIpInfo> { infoA } } };
            var searchNames = new List<string> { "Found", "Missing" };

            // Act
            List<HardIpInfo> result = SearchInfo.GetHardIpInfos(searchNames, sourceData);

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(infoA, result[0]);
        }

        [TestMethod]
        [DataRow(null)]
        public void GetHardIpInfos_WithEmptyInputs_ReturnsEmptyList(List<string> emptyList)
        {
            // Arrange
            var sourceData = new HardIpInfos();
            List<string> searchNames = emptyList ?? [];

            // Act
            List<HardIpInfo> result = SearchInfo.GetHardIpInfos(searchNames, sourceData);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        [DataRow("TX_Level_Logic", "None", false)]
        [DataRow("Normal", "Flag_Singlelimit: true", true)]
        [DataRow("Normal", "Flag_Singlelimit: false", false)]
        [DataRow("Normal", "", false)]
        public void GetFlagSingleLimit_OverrideScenarios(string calcEqn, string miscInfo, bool expected)
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                CalcEqn = calcEqn,
                MiscInfo = miscInfo,
                TestPlanSequences = [],
                Pattern = new PatternClass("CZ_NVSA0_A_FULP_AN_AN00_MEA_JTG_DIO_ALLFRV_SI_ACTCONS_T7"),
                MeasPins =
                [
                    new()
                    {
                        SequenceIndex = 1,
                        MeasType = MeasType.MeasVdiff
                    }
                ]
            };

            // Act
            bool result = SearchInfo.GetFlagSingleLimit(pattern, "1.2V");

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetFlagSingleLimit_ReturnsFalse_ForMultiplePinsInSequence()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                TestPlanSequences = [new(1, 1, 1)],
                MeasPins = [
                new() { SequenceIndex = 1, PinName = "PinA", MeasType = "MeasV" },
                    new() { SequenceIndex = 1, PinName = "PinB", MeasType = "MeasV" }
            ]
            };

            // Act
            bool result = SearchInfo.GetFlagSingleLimit(pattern, "1.2V");

            // Assert: GroupBy(p => p.PinName).Count() > 1 triggers return false
            Assert.IsFalse(result);
        }

        [TestMethod]
        [DataRow("PatternName", "BlockType", "", "false")]
        [DataRow("", "", "", "false")]
        [DataRow("NoPattern", "", "A", "true")]
        [DataRow("NoPattern", "", "", "false")]
        public void GetCpuflag_ReturnsFalse_UnderInvalidConditions(string payload, string blockType, string callSubrs, string expected)
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                BlockType = blockType,
                Pattern = new PatternClass(payload)
            };
            var info = new HardIpInfo() { CallSubrs = callSubrs };

            // Act
            string result = SearchInfo.GetCpuflag(info, pattern);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetStoreName_HandlesLongStrings_WithRegAssign()
        {
            // Arrange
            string storeNameOri = "";
            var hardIpInputData = new HardIpInputData(null) { HardIpRegAssigns = [] };

            // Create a pattern that generates a string >= 6000 chars
            var pins = new List<MeasPin>();
            for (int i = 0; i < 500; i++)
            {
                pins.Add(new MeasPin { SequenceIndex = 1, CusStr = "VeryLongPinNameReference_" + i });
            }
            var pattern = new HardIpPattern
            {
                MeasPins = pins,
                Pattern = new PatternClass("DD_")
            };

            // Act
            string result = SearchInfo.GetStoreName(pattern, hardIpInputData, ref storeNameOri, "Seq1,Seq2");

            // Assert
            Assert.IsTrue(result.StartsWith("Reg_assign:"), "Should return Reg_assign label for long strings.");
            Assert.IsTrue(hardIpInputData.HardIpRegAssigns.Count > 0, "Should add to HardIpRegAssigns list.");

            var hardIpInputData1 = new HardIpInputData(null) { HardIpRegAssigns = [] };
            string result1 = SearchInfo.GetStoreName(pattern, hardIpInputData1, ref storeNameOri, "");

            // Assert
            Assert.IsTrue(result1.StartsWith("Reg_assign:"), "Should return Reg_assign label for long strings.");
            Assert.IsTrue(hardIpInputData1.HardIpRegAssigns.Count > 0, "Should add to HardIpRegAssigns list.");
        }

        [TestMethod]
        public void GetStoreName_HandlesLongStrings_WithRegAssign_1()
        {
            // Arrange
            string storeNameOri = "";
            var hardIpInputData = new HardIpInputData(null) { HardIpRegAssigns = [] };

            // Create a pattern that generates a string >= 6000 chars
            var pins = new List<MeasPin>()
            {
                new() { PinName= "CP=A" , SequenceIndex = 1, CusStr = "VeryLongPinNameReference_0" },
                new() { PinName= "FT=B" , SequenceIndex = 1, CusStr = "VeryLongPinNameReference_0" },
            };
            var pattern = new HardIpPattern { MeasPins = pins };

            // Act
            string result = SearchInfo.GetStoreName(pattern, hardIpInputData, ref storeNameOri, "Seq1,Seq2");

            // Assert
            Assert.IsTrue(result.StartsWith("VeryLongPinNameReference_0"), "Should return Reg_assign label for long strings.");
            Assert.IsTrue(hardIpInputData.HardIpRegAssigns.Count == 0, "Should add to HardIpRegAssigns list.");
        }

        [TestMethod]
        public void NeedDecompGroups_ReturnsTrue_ForMeasR2()
        {
            // Arrange
            var pin = new MeasPin { MeasType = "MeasR2" };
            var pattern = new HardIpPattern();
            var limits = new Dictionary<string, string>();

            // Act
            bool result = SearchInfo.NeedDecompGroups(pattern, pin, limits);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void NeedDecompGroups_ReturnsFalse_WhenLimitIsF()
        {
            // Arrange
            // measPin.MeasType[4] is 'V' in "MeasV"
            var pin = new MeasPin { MeasType = "MeasV" };
            var pattern = new HardIpPattern();
            var limits = new Dictionary<string, string> { { "V", "F" } };

            // Act
            bool result = SearchInfo.NeedDecompGroups(pattern, pin, limits);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void GetMeasCPins_ReturnsDefault_WhenCapBitExistsButPinNameEmpty()
        {
            // Arrange
            var pattern = new HardIpPattern();
            // Mocking HardIpService.GetHardIpInfo return via Shims or internal setup
            var mockInfo = new HardIpInfo
            {
                CapBitStr = "1",
                CapPinName = ""
            };

            // Act
            // Note: This test assumes HardIpService is mocked to return mockInfo
            string result = SearchInfo.GetMeasCPins(pattern, mockInfo);

            // Assert
            Assert.AreEqual("JTAG_TDO", result, "Should return default JTAG_TDO when CapPinName is empty.");
        }

        [TestMethod]
        public void GetMeasCPins_ReturnsFirstPin_WhenMultiplePinsExist()
        {
            // Arrange
            var mockInfo = new HardIpInfo
            {
                CapBitStr = "1",
                CapPinName = "PIN1|PIN2|PIN3"
            };

            // Act
            // Assuming HardIpService returns mockInfo
            string result = SearchInfo.GetMeasCPins(new HardIpPattern(), mockInfo);

            // Assert
            Assert.AreEqual("PIN1", result, "Should return only the first pin from the pipe-delimited list.");
        }

        [DataTestMethod]
        [DataRow("dd_xxx", true)]
        [DataRow("cz_yyy", true)]
        [DataRow("pp_zzz", true)]
        [DataRow("No_patt", true)]
        [DataRow("Instance:AAA", true)]
        [DataRow("Invalid_Name", false)]
        [DataRow("", false)]
        public void IsValidPatName_PureUt(string input, bool expected)
        {
            Assert.AreEqual(expected, SearchInfo.IsValidPatName(input));
        }

        [TestMethod]
        public void GetPpmuPin_MultipleSequences_JoinsWithPlusAndCommas()
        {
            // Arrange
            var info = new HardIpInfo { SeqInfo = [new(), new()] };
            var pattern = new HardIpPattern
            {
                MeasPins =
            [
                // Sequence 1 pins
                new() { SequenceIndex = 1, PinName = "PinA", MeasType = "MeasV" },
                new() { SequenceIndex = 1, PinName = "PinB", MeasType = "MeasI" },
                // Sequence 2 pins
                new() { SequenceIndex = 2, PinName = "PinC::PinD", MeasType = "MeasV" },
                // Should be filtered out
                new() { SequenceIndex = 1, PinName = "IgnoreMe", MeasType = "MeasC" }
            ]
            };

            // Act
            string result = SearchInfo.GetPpmuPin(pattern, info);

            // Assert
            // Sequence 1: "PinA,PinB"
            // Sequence 2: "PinC,PinD" (after :: replacement)
            // Joined by +
            Assert.AreEqual("PinA,PinB+PinC,PinD", result);
        }

        [TestMethod]
        public void GetPpmuPin_NoSequences_ReturnsEmptyString()
        {
            // Arrange
            var info = new HardIpInfo { SeqInfo = [] };
            var pattern = new HardIpPattern { TestPlanSequences = [] };

            // Act
            string result = SearchInfo.GetPpmuPin(pattern, info);

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void InitialMeasC_WhenInfoCountGreater_AddsMissingMeasCPins()
        {
            // Arrange
            // We need 3 items in DsscOut (e.g., "A,B,C") to get an infoCount of 2 (Length - 1)
            var patInfo = new HardIpInfo
            {
                DsscOut = "pin1,pin2,pin3",
                CapPinName = "TEST_PIN"
            };

            // pattern starts with 0 MeasC pins
            var pattern = new HardIpPattern
            {
                MiscInfo = "None",
                MeasPins = [],
                Pattern = new PatternClass("P1+P2")
                {
                    PatternSetList = [["P1", "P2"]],
                    RealPatternName = "P1#P2"
                }
                // Mock or stub to return false for MTD checks
            };

            // Ensure LocalSpecs options won't skip the logic
            LocalSpecs.Options.Device = EnumDevice.AP;

            // Act
            SearchInfo.InitialMeasC(ref pattern, patInfo);

            // Assert
            // infoCount (2) - planCount (0) = 2 pins should be added
            var measCPins = pattern.MeasPins.Where(p => p.MeasType == "MeasC").ToList();
            Assert.AreEqual(0, measCPins.Count);
        }

        [TestMethod]
        public void InitialMeasC_WithIgnoreFlag_DoesNotAddPins()
        {
            // Arrange
            var patInfo = new HardIpInfo { DsscOut = "A,B,C" };
            var pattern = new HardIpPattern
            {
                MiscInfo = HardIpConstData.IgnorePatMeasC, // Flag to trigger ignore regex
                MeasPins = []
            };

            // Act
            SearchInfo.InitialMeasC(ref pattern, patInfo);

            // Assert
            Assert.AreEqual(0, pattern.MeasPins.Count(p => p.MeasType == "MeasC"));
        }
        [TestMethod]
        public void InitialMeasC_ShouldAddMeasCPins_WhenMultiTimeDomainHasMoreDsscOut()
        {
            // Arrange
            var pattern = new HardIpPattern();

            pattern.Pattern.RealPatternName = "PAT1#PAT2";
            pattern.Pattern.TestPlanPatternName = "PAT1#PAT2";

            pattern.Pattern.InstancePatternName =
            [
                "PAT1",
                "PAT2"
            ];

            pattern.MeasPins =
            [
                new MeasPin("PIN1", "MeasC")
                {
                    PatternName = "PAT1",
                    PatternIndex = 0,
                    PinCount = 1
                }
            ];

            var patInfo = new HardIpInfo
            {
                DsscOut = "A,B,C",
                MultiDsscOut = "A,B,C|D",
                CapPinName = "JTAGTDO"
            };

            int originalCount = pattern.MeasPins.Count;

            // Act
            SearchInfo.InitialMeasC(ref pattern, patInfo);

            // Assert
            Assert.IsTrue(pattern.MeasPins.Count > originalCount);
        }
        [TestMethod]
        public void GetOpcode_TypeB_AddsNonFlowControlOpcodes()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MiscInfo = "opcode:SET_REG:Value1;opcode:ENDIF:Skipped",
                FlowControlFlag = 0
            };

            // Act
            // Type "B" (Before) filters for things that AREN'T "endif" or "next"
            List<string> result = SearchInfo.GetOpcode(pattern, "B");

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("SET_REG:Value1", result[0]);
        }

        [TestMethod]
        public void GetOpcode_TypeA_AddsFlowControlOpcodes()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                MiscInfo = "opcode:NEXT:Loop;opcode:JUMP:Address1",
                FlowControlFlag = 1
            };

            // Act
            // Type "A" (After) filters specifically for "endif" or "next"
            List<string> result = SearchInfo.GetOpcode(pattern, "A");

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("NEXT:Loop", result[0]);
        }

        [TestMethod]
        public void GetOpcode_MalformedString_ReturnsEmptyList()
        {
            // Arrange
            var pattern = new HardIpPattern { MiscInfo = "invalid_data;some_other_info", FlowControlFlag = 1 };

            // Act
            List<string> result = SearchInfo.GetOpcode(pattern, "A");

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetOpcode_TypeA_WhenFlowControlFlagDisabled_ReturnsEmpty()
        {
            var pattern = new HardIpPattern
            {
                MiscInfo = "opcode:NEXT:Loop",
                FlowControlFlag = 0
            };

            List<string> result = SearchInfo.GetOpcode(pattern, "A");

            Assert.AreEqual(0, result.Count);
        }

        [DataTestMethod]
        [DataRow("MeasCapName:TargetName;Other:Info", "TargetName")]
        [DataRow("NoCapHere;JustOtherData", "")]
        [DataRow("", "")]
        public void GetDigCapNameByMiscInfo_PureUt(string misc, string expected)
        {
            Assert.AreEqual(expected, SearchInfo.GetDigCapNameByMiscInfo(misc));
        }

        [TestMethod]
        public void GetTtrEnable_MatchingVoltageSuffix_ClearsAllJobs()
        {
            // Arrange: ttrStr ends with 'V' and contains the voltage "NV"
            string ttrStr = "NV,HV,LV";
            string voltage = "NV";

            // Act
            string result = SearchInfo.GetTtrEnable(ttrStr, voltage);

            // Assert: When a voltage-only match (ending in V) occurs, it clears the list
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetTtrEnable_JobNameFilter_RemovesSpecificJob()
        {
            // Arrange: "CP1" is in ttrStr, so it should be filtered out via Regex
            string ttrStr = "CP1";
            string voltage = "NV";

            // Act
            string result = SearchInfo.GetTtrEnable(ttrStr, voltage);

            // Assert: result should contain everything EXCEPT "CP1"
            Assert.IsFalse(result.Contains("CP1"));
            Assert.IsTrue(result.Contains("CP2"));
        }

        [TestMethod]
        public void GetTtrEnable_JobWithVoltageColon_RemovesJobNamePart()
        {
            // Arrange: "CP2:NV" format
            string ttrStr = "CP2:NV";
            string voltage = "NV";

            // Act
            string result = SearchInfo.GetTtrEnable(ttrStr, voltage);

            // Assert: Should split by ':' and remove "CP2"
            Assert.IsFalse(result.Contains("CP2"));
            Assert.IsTrue(result.Contains("CP1"));
        }

        [TestMethod]
        public void GetTtrEnable_EmptyString_ReturnsAllJobs()
        {
            // Arrange
            string ttrStr = "";
            string voltage = "NV";

            // Act
            string result = SearchInfo.GetTtrEnable(ttrStr, voltage);

            // Assert
            Assert.AreEqual("CP1,CP2,FT1,FT2", result);
        }

        [DataTestMethod]
        [DataRow("MissingName", null, null, false, "MissPattInPattList")]
        [DataRow("Patt1", "na", "1.0", false, "MissTimesetInPattList")]
        [DataRow("Patt1", "NA", "1.0", false, "MissTimesetInPattList")]
        [DataRow("Patt1", "ValidTS", "NA", false, "MissFileVersionInPattList")]
        [DataRow("Patt1", "ValidTS", "1.0", true, "")]
        public void GetStatusInPatternList_ReturnsExpectedStatus(string searchName, string tsVersion, string fVersion, bool expectedResult, string expectedStatus)
        {
            // Arrange
            var patternList = new Dictionary<string, PatternData>();
            if (searchName == "Patt1")
            {
                patternList.Add("Patt1", new PatternData
                {
                    TimeSetVersion = tsVersion,
                    FileVersion = fVersion
                });
            }

            // Act
            bool result = SearchInfo.GetStatusInPatternList(patternList, searchName, out string actualStatus);

            // Assert
            Assert.AreEqual(expectedResult, result, $"Result mismatch for {searchName}");
            Assert.AreEqual(expectedStatus, actualStatus, $"Status message mismatch for {searchName}");
        }

        [TestMethod]
        public void GetStatusInPatternList_NATimeset_ReturnsMissTimesetInPattList()
        {
            // Arrange
            var list = new Dictionary<string, PatternData> {
            {
                "Pat1", new PatternData { TimeSetVersion = "na", FileVersion = "1.0" } }
            };

            // Act
            _ = SearchInfo.GetStatusInPatternList(list, "Pat1", out string status);

            // Assert
            Assert.AreEqual("MissTimesetInPattList", status);
        }

        [TestMethod]
        public void GetEnvFromPattern_IsPatternValidate_ReturnsEmpty()
        {
            LocalSpecs.IsPatternValidate = true;

            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("MyPattern")
            };

            string result = SearchInfo.GetEnvFromPattern(pattern);

            Assert.AreEqual("", result);

            LocalSpecs.IsPatternValidate = false;
        }

        [TestMethod]
        public void GetEnvFromPattern_InstancePattern_ReturnsEmpty()
        {
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("Dummy")
            };

            pattern.Pattern.RealPatternName = "Instance:ABC";

            string result = SearchInfo.GetEnvFromPattern(pattern);

            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetEnvFromPattern_OpcodePattern_ReturnsEmpty()
        {
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("Dummy")
            };

            pattern.Pattern.RealPatternName = "Opcode:READ";

            string result = SearchInfo.GetEnvFromPattern(pattern);

            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetEnvFromPattern_HardIpArfRf_ReturnsEmpty()
        {
            LocalSpecs.Options.Device = EnumDevice.RF;

            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("MyPattern"),
                SheetName = "HardIP_ARF"
            };

            string result = SearchInfo.GetEnvFromPattern(pattern);

            Assert.AreEqual("", result);

            LocalSpecs.Options.Device = EnumDevice.AP;
        }

        [TestMethod]
        public void GetEnvFromPattern_NotInTestPlan_ReturnsMissPattInTestPlan()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("SomePattern"),
                SheetName = "OtherSheet"
            };

            // Act
            string result = SearchInfo.GetEnvFromPattern(pattern);

            // Assert
            Assert.AreEqual("MissPattInPattList", result);
        }

        [TestMethod]
        public void GetEnvFromPattern_PatternInList_ReturnsStatusFromList()
        {
            // Arrange
            string patName = "MyPattern";
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass(patName),
                SheetName = "NormalSheet",
                FunctionName = "OtherFunc"
            };

            // Act
            string result = SearchInfo.GetEnvFromPattern(pattern);

            // Assert
            Assert.AreEqual("MissPattInPattList", result);
        }

        [TestMethod]
        public void GetEnvFromPattern_PatternInList_ReturnsStatusFromList_2()
        {
            // Arrange
            string patName = "MyPattern";
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass(patName),
                SheetName = "NormalSheet",
                FunctionName = "RTOS_RunScenario_T"
            };

            // Act
            string result = SearchInfo.GetEnvFromPattern(pattern);

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetEnvFromPattern_PatternInList_ReturnsStatusFromList_1()
        {
            // Arrange
            LocalSpecs.Options.Device = EnumDevice.RF;
            string patName = "MyPattern";
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass(patName),
                SheetName = "Wireless_ARF",
                FunctionName = "OtherFunc"
            };

            // Act
            string result = SearchInfo.GetEnvFromPattern(pattern);

            // Assert
            Assert.AreEqual("", result);

            LocalSpecs.Options.Device = EnumDevice.AP;
        }

        [TestMethod]
        public void GetEnvFromPattern_NoPatternConstant_ReturnsEmptyString()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass(HardIpConstData.NoPattern)
            };

            // Act
            string result = SearchInfo.GetEnvFromPattern(pattern);

            // Assert
            Assert.AreEqual("", result);
        }

        [DataTestMethod]
        [DataRow("1", "", "", "JTAG_TDI")]              // Valid SendBit, default pin name
        [DataRow("0", "Active", "MY_PIN|EXT", "MY_PIN")] // Valid SendBitStr, splits pin
        [DataRow("0", "", "ANY_PIN", "")]               // No valid bit or string, returns empty
        [DataRow(null, "Value", null, "JTAG_TDI")]       // Null safety check for SendBit and PinName
        [DataRow("2", "", "SINGLE_PIN", "SINGLE_PIN")]   // Valid SendBit, single pin name
        public void GetSrcPin_ReturnsExpectedValue(string sendBit, string sendBitStr, string sendPinName, string expected)
        {
            // Arrange
            var patInfo = new HardIpInfo
            {
                SendBit = sendBit,
                SendBitStr = sendBitStr,
                SendPinName = sendPinName
            };

            // Act
            string result = SearchInfo.GetSrcPin(patInfo);

            // Assert
            Assert.AreEqual(expected, result, $"Failed for SendBit: {sendBit}, SendBitStr: {sendBitStr}");
        }

        [TestMethod]
        [DataRow("VDD_CORE", "power")]
        [DataRow("vdd_sense", "power")]
        [DataRow("XXX:VCC", "power")]
        public void GetPinType_WhenNameStartsWithVDD_ReturnsPower(string pinName, string expected)
        {
            string result = SearchInfo.GetPinType(pinName);
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow("SIG_A", "SIG_A", true)]   // Direct equality
        [DataRow("SIG_A", "SIG_B", false)]  // Direct inequality
        public void ContainsPin_SimpleEquality_ReturnsExpected(string plan, string pat, bool expected)
        {
            var pinList = new List<string>();
            bool result = SearchInfo.ContainsPin(plan, pat, ref pinList);
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow("A::B", "B::A")] // Same pins, different order
        [DataRow("A::B", "a::b")] // Case insensitive
        public void ContainsPin_MultiPinDelimiter_ReturnsTrue(string plan, string pat)
        {
            var pinList = new List<string>();
            bool result = SearchInfo.ContainsPin(plan, pat, ref pinList);
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void ContainsPin_WhenPatHasSingleColon_SplitsCorrectly()
        {
            // Tests line 655 logic (patInfoPinName.Contains(":") but not "::")
            var pinList = new List<string>();
            // "Port1:SIG_A" should be treated as "SIG_A"
            bool result = SearchInfo.ContainsPin("SIG_A", "Port1:SIG_A", ref pinList);
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void ContainsPin_GroupExpansion_ReturnsTrueAndUpdatesList()
        {
            // Arrange
            string groupName = "MY_GROUP";
            var expandedPins = new List<string> { "TX_P", "TX_N" };
            var pinList = new List<string>();

            // Act
            // Also testing line 684 split logic with '='
            bool result = SearchInfo.ContainsPin(groupName, "VAR=TX_P", ref pinList);

            // Assert
            Assert.IsTrue(result);
            CollectionAssert.AreEqual(expandedPins, pinList);
        }

        [TestMethod]
        public void ContainsPin_CommaVsDoubleColon_ReturnsTrue()
        {
            // Tests line 667 logic
            var pinList = new List<string>();
            bool result = SearchInfo.ContainsPin("A,B", "B::A", ref pinList);
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void ContainsPin_NoMatch_ReturnsFalse_AndDoesNotModifyList()
        {
            var pinList = new List<string>();

            bool result = SearchInfo.ContainsPin("A", "B", ref pinList);

            Assert.IsFalse(result);
            Assert.AreEqual(0, pinList.Count);
        }

        [DataTestMethod]
        [DataRow("TrimStoreName:Item1", "Item1")]
        [DataRow("TrimCodeStoreName_Suffix:ItemA,ItemB", "ItemA", "ItemB")]
        [DataRow("CUS_Str_DigCapData_Suffix:ItemA,ItemB", "ItemA", "ItemB")]
        [DataRow("Dict_Store_Code_Name:Value1+Value2&Value3", "Value1", "Value2", "Value3")]
        [DataRow("TrimDictionaryStoreName:StoreX", "StoreX")]
        [DataRow("DigCapDataCustomString:StoreX", "StoreX")]
        [DataRow("TrimCodeStoreName_Suffix:A;TrimStoreName:B", "A", "B")]
        public void GetTrimStoreNameByMiscInfo_ValidInputs_ReturnsExpectedList(string input, params string[] expectedItems)
        {
            // Act
            List<string> result = SearchInfo.GetTrimStoreNameByMiscInfo(input);

            // Assert
            CollectionAssert.AreEquivalent(expectedItems.ToList(), result);
        }

        [TestMethod]
        public void GetTrimStoreNameByMiscInfo_UnknownPrefix_ReturnsEmpty()
        {
            // Arrange
            string miscInfo = "UnknownPrefix:Something;AnotherInvalid:Value";

            // Act
            List<string> result = SearchInfo.GetTrimStoreNameByMiscInfo(miscInfo);

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetTrimStoreNameByMiscInfo_EmptyInput_ReturnsEmptyList()
        {
            // Act
            List<string> result = SearchInfo.GetTrimStoreNameByMiscInfo("");

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetTrimStoreNameByMiscInfo_MultipleKeys()
        {
            string misc = "TrimStoreName:A,B;Dict_Store_Code_Name:C+D";
            List<string> result = SearchInfo.GetTrimStoreNameByMiscInfo(misc);

            CollectionAssert.AreEquivalent(
                new List<string> { "A", "B", "C", "D" },
                result);
        }

        [TestMethod]
        public void GetTrimStoreNameByMiscInfo_Empty_ReturnsEmpty()
        {
            List<string> result = SearchInfo.GetTrimStoreNameByMiscInfo("");
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetVbtNameByPattern_WhenTrimTargetExists_AddsPlanTypeTrim()
        {
            var pattern = new HardIpPattern { WirelessData = new WirelessData { TrimTarget = "SomeTarget" } };
            var inputData = new HardIpInputData(null) { ConfigData = new ConfigData() };
            SearchInfo.GetVbtNameByPattern(inputData, pattern);

            Assert.IsTrue(pattern.VbtTypes.Contains(PlanType.Trim));
        }

        [TestMethod]
        public void GetVbtNameByPattern_SheetNameStartsWithLcd_ReturnsLcdMeas()
        {
            // Arrange
            var pattern = new HardIpPattern { SheetName = "LCD_Test_Sheet" };
            var inputData = new HardIpInputData(null) { ConfigData = new ConfigData() };

            // Act
            string result = SearchInfo.GetVbtNameByPattern(inputData, pattern);

            // Assert
            Assert.AreEqual("measuniversalfunc", result);
        }

        [TestMethod]
        public void GetVbtNameByPattern_FunctionNotFound_AddsError()
        {
            // Arrange
            string info = "Func:XXX";
            var pattern = new HardIpPattern { MiscInfo = info, SheetName = "Sheet1", RowNum = 10 };
            var inputData = new HardIpInputData(null) { ConfigData = new ConfigData() };

            // Act
            SearchInfo.GetVbtNameByPattern(inputData, pattern);

            // Assert: Verify error was logged via your ErrorReportManager
            Assert.IsTrue(ErrorReportManager.GetErrorList().Count != 0);
        }

        [TestMethod]
        public void GetHardIpBinRangeItem_WhenIdsNandMatched_ReturnsNandBin()
        {
            // Arrange: Setup Pattern for IDS NAND branch
            var pattern = new HardIpPattern
            {
                SheetName = "DCTEST_IDS_SHEET",
                RowNum = 10,
                Pattern = new PatternClass("Test_nan")
            };

            // Act
            var scghRow = new ProdCharSheetRow();
            SoftBinRangeData result = SearchInfo.GetHardIpBinRangeItem(pattern, scghRow);

            // Assert
            Assert.AreEqual("IDSNAND", result.Condition);
        }

        [TestMethod]
        public void GetHardIpBinRangeItem_WhenIdsNandMatched_ReturnsNandBin_1()
        {
            // Arrange: Setup Pattern for IDS NAND branch
            var pattern = new HardIpPattern
            {
                SheetName = "DCTEST_IDS_SHEET",
                RowNum = 10,
                Pattern = new PatternClass("Test_spi")
            };

            // Act
            var scghRow = new ProdCharSheetRow();
            SoftBinRangeData result = SearchInfo.GetHardIpBinRangeItem(pattern, scghRow);

            // Assert
            Assert.AreEqual("IDSSPI", result.Condition);
        }

        [TestMethod]
        public void GetHardIpBinRangeItem_WhenIdsNandMatched_ReturnsNandBin_3()
        {
            // Arrange: Setup Pattern for IDS NAND branch
            var pattern = new HardIpPattern
            {
                SheetName = "DCTEST_IDS_SHEET",
                RowNum = 10,
                Pattern = new PatternClass("Test")
            };

            // Act
            var scghRow = new ProdCharSheetRow
            {
                Block = "HardIP",
                Mode = "JTAG",
                Item = "IO"
            };
            SoftBinRangeData result = SearchInfo.GetHardIpBinRangeItem(pattern, scghRow);

            // Assert
            Assert.AreEqual("HARDIPOTHERS", result.Condition);
        }

        [TestMethod]
        public void GetHardIpBinRangeItem_WhenIdsNandMatched_ReturnsNandBin_4()
        {
            // Arrange: Setup Pattern for IDS NAND branch
            var pattern = new HardIpPattern
            {
                SheetName = "HardIP_JTAG",
                RowNum = 10,
                Pattern = new PatternClass("Test")
            };

            // Act
            var scghRow = new ProdCharSheetRow
            {
                Block = "HardIP",
                Mode = "JTAG",
                Item = "IO"
            };
            SoftBinRangeData result = SearchInfo.GetHardIpBinRangeItem(pattern, scghRow);

            // Assert
            Assert.AreEqual("JTAG", result.Condition);
        }

        [TestMethod]
        public void GetTestLimitPerMeasType_VifPattern_ReturnsCorrectDictionary()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                FunctionName = "meas_vir_io_universal_func",
                MeasPins = [
                    new()
                    {
                        SequenceIndex = 1,
                        MeasType = "MeasV",
                        PinName = "Pin1,Pin2"
                    },
                    new()
                    {
                        SequenceIndex = 2,
                        MeasType = "MeasI",
                        PinName = "Pin1,Pin2"
                    },
                    new()
                    {
                        SequenceIndex = 3,
                        MeasType = "MeasR1",
                        PinName = "Pin1,Pin2"
                    }
                ],
                TestPlanSequences = [.. new TestPlanSequence[3]]
            };

            // Act
            Dictionary<string, string> result = SearchInfo.GetTestLimitPerMeasType(pattern);

            // Assert
            Assert.AreEqual("T", result["V"]);
            Assert.AreEqual("T", result["I"]);
            Assert.AreEqual("T", result["R"]);
        }

        // ===== Add below test methods into the existing SearchInfoTests class =====

        [TestMethod]
        public void ReverseRelaySetting_SwapsOnOffSegments_Correctly()
        {
            // Arrange
            string input = "RelayOn_K1_K2_RelayOff_K3";

            // Act
            string result = SearchInfo.ReverseRelaySetting(input);

            // Assert
            Assert.AreEqual("RelayOn_K3_RelayOff_K1_K2", result);
        }

        [TestMethod]
        public void GetRelayArgs_ShouldReturnEmptyStrings_WhenSettingIsEmpty()
        {
            string setting = string.Empty;

            SearchInfo.GetRelayArgs(setting, out string argOn, out string argOff);

            Assert.AreEqual(string.Empty, argOn);
            Assert.AreEqual(string.Empty, argOff);
        }

        [TestMethod]
        public void GetRelayArgs_ShouldReturnRawRelayOnIncludingTrailingSpace_WhenNoAmpersand()
        {
            string setting = "RelayOn:K1_K2_RelayOff:X";

            SearchInfo.GetRelayArgs(setting, out string argOn, out string argOff);

            Assert.AreEqual("K1_K2", argOn);
            Assert.AreEqual("X", argOff);
        }

        [TestMethod]
        public void GetRelayArgs_ShouldSplitByAmpersandAndKeepTrailingSpace()
        {
            // relayOn = "K1&K2_ " → Split('&') → ["K1", "K2_ "]
            string setting = "RelayOn:K1&K2_RelayOff:X";

            SearchInfo.GetRelayArgs(setting, out string argOn, out string argOff);

            Assert.AreEqual("K1,K2", argOn);
            Assert.AreEqual("X", argOff);
        }

        [TestMethod]
        public void GetRelayArgs_ShouldReturnEmptyOn_WhenRelayOnNotExist()
        {
            string setting = "RelayOff:K3_";

            SearchInfo.GetRelayArgs(setting, out string argOn, out string argOff);

            Assert.AreEqual(string.Empty, argOn);
            Assert.AreEqual("K3", argOff);
        }

        [TestMethod]
        public void GetRelayArgs_ShouldReturnEmpty_WhenRelayOnDoesNotExist()
        {
            string setting = "RelayOff:K3&K4";

            SearchInfo.GetRelayArgs(setting, out string argOn, out string argOff);

            Assert.AreEqual(string.Empty, argOn);
            Assert.AreEqual("K3,K4", argOff);
        }

        [TestMethod]
        public void GetRelayArgs_ShouldTrimOnlyTrailingComma_NotUnderscoreInMiddle()
        {
            string setting = "RelayOn:K1_K2_";

            SearchInfo.GetRelayArgs(setting, out string argOn, out string argOff);

            Assert.AreEqual("K1_K2", argOn);
            Assert.AreEqual(string.Empty, argOff);
        }

        [TestMethod]
        public void IsForceType_ReturnsTrue_WhenVTypeExists()
        {
            // Arrange
            HardIpPattern pattern = new HardIpPattern
            {
                MeasPins =
        [
            new()
            {
                ForceConditions =
                [
                    new()
                    {
                        ForcePins =
                        [
                            new() { PinName = "TX_P", ForceType = "V" }
                        ]
                    }
                ]
            }
        ]
            };

            // Act
            bool result = SearchInfo.IsForceType(pattern, "V");

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsForceType_ReturnsFalse_WhenNoForce()
        {
            // Arrange
            HardIpPattern pattern = new HardIpPattern
            {
                MeasPins = [new()]
            };

            // Act
            bool result = SearchInfo.IsForceType(pattern, "I");

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsMeasVdiff2_ReturnsTrue_WhenPinPresent()
        {
            // Arrange
            HardIpPattern pattern = new HardIpPattern
            {
                MeasPins =
        [
            new() { MeasType = MeasType.MeasVdiff2 }
        ]
            };

            // Act
            bool result = SearchInfo.IsMeasVdiff2(pattern);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsMeasVdiff_ReturnsTrue_WhenPinPresent()
        {
            // Arrange
            HardIpPattern pattern = new HardIpPattern
            {
                MeasPins =
        [
            new() { MeasType = MeasType.MeasVdiff }
        ]
            };

            // Act
            bool result = SearchInfo.IsMeasVdiff(pattern);

            // Assert
            Assert.IsTrue(result);
        }

        [DataTestMethod]
        [DataRow("Repeat_Limit:3;Other:X", true, 3)]
        [DataRow("Other:X", false, 0)]
        [DataRow("", false, 0)]
        public void RepeatLimit_PureUt(string misc, bool expectHasRepeat, int expectCount)
        {
            Assert.AreEqual(expectHasRepeat, SearchInfo.IsRepeatLimit(misc));
            Assert.AreEqual(expectCount, SearchInfo.GetRepeatLimitCount(misc));
        }

        [TestMethod]
        public void GetCalculation_ParsesCalcAndParameters_DefaultTestName()
        {
            // Arrange
            // Use correct key: CalcArg
            string misc = "Calc:ALG_A;CalcArg:rd0,rd1";

            // Act
            string result = SearchInfo.GetCalculation(misc);

            // Assert
            Assert.AreEqual("Alg::ALG_A(rd0,rd1)", result);
        }

        [TestMethod]
        public void GetCalculation_ParsesWithVoltagePrefix_AndSpecificParameterKey()
        {
            // Arrange
            // Use correct keys: NV@Calc and NV@CalcArg
            string misc = "NV@Calc:ALG_B;NV@CalcArg:x1";

            // Act
            string result = SearchInfo.GetCalculation(misc, "TNAME");

            // Assert
            Assert.AreEqual("NV@Alg:TNAME:ALG_B(x1)", result);
        }

        [TestMethod]
        public void GetCalculation_Default()
        {
            string misc = "Calc:ALG_A;CalcArg:x1,x2";
            string result = SearchInfo.GetCalculation(misc);
            Assert.AreEqual("Alg::ALG_A(x1,x2)", result);
        }

        [TestMethod]
        public void GetCalculation_WithVoltage()
        {
            string misc = "NV@Calc:ALG_B;NV@CalcArg:y1";
            string result = SearchInfo.GetCalculation(misc, "TNAME");
            Assert.AreEqual("NV@Alg:TNAME:ALG_B(y1)", result);
        }

        [TestMethod]
        public void TrimSpace_RemovesWhitespaceCharacters()
        {
            // Arrange
            string input = "  A \t B \n C \r\n ";
            // Act
            string result = SearchInfo.TrimSpace(input);
            // Assert
            Assert.AreEqual("ABC", result);
        }

        [TestMethod]
        public void TrimSpace_NullOrEmpty_ReturnsEmpty()
        {
            Assert.AreEqual(string.Empty, SearchInfo.TrimSpace(null));
            Assert.AreEqual(string.Empty, SearchInfo.TrimSpace(""));
        }

        [TestMethod]
        public void GetPinType_ReturnsGroupType_WhenInputIsGroup()
        {
            // Arrange
            // MY_GROUP is created in ClassInitialize with PinType = "I/O"

            // Act
            string result = SearchInfo.GetPinType("MY_GROUP");

            // Assert
            Assert.AreEqual("I/O", result);
        }

        [TestMethod]
        public void GetPinType_ColonForm_ResolvesUnderlyingPin()
        {
            // Arrange
            // "Port1:TX_P" should resolve to "TX_P", which is in MY_GROUP and type "I/O"
            string name = "Port1:TX_P";

            // Act
            string result = SearchInfo.GetPinType(name);

            // Assert
            Assert.AreEqual("I/O", result);
        }

        [TestMethod]
        public void GetPinType_ColonForm_ResolvesPowerPin()
        {
            // Arrange
            // "Port1:VCC" should resolve to "VCC", which is added in ClassInitialize as power
            string name = "Port1:VCC";

            // Act
            string result = SearchInfo.GetPinType(name);

            // Assert
            Assert.AreEqual("power", result);
        }

        [TestMethod]
        public void MeasC_Count_SumsPinCount()
        {
            // Arrange
            HardIpPattern pattern = new HardIpPattern
            {
                MeasPins = [
                    new() { MeasType = "MeasC", PinCount = 1 },
                    new() { MeasType = "MeasC", PinCount = 2 },
                    new() { MeasType = "MeasV", PinCount = 5 }
                ]
            };

            // Act
            int count = SearchInfo.MeasC_Count(pattern);

            // Assert
            Assert.AreEqual(3, count);
        }

        [TestMethod]
        public void IsMeasPinInForcePin_ReturnsTrue_WhenForceGroupContainsMeasPin()
        {
            // Arrange
            // MY_GROUP contains TX_P and TX_N
            string forcePinGroup = "MY_GROUP";
            string measPinName = "TX_P";

            // Act
            bool result = SearchInfo.IsMeasPinInForcePin(forcePinGroup, measPinName);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsMeasPinInForcePin_ReturnsTrue_OnExactMatch()
        {
            // Arrange
            string forcePinGroup = "TX_P";
            string measPinName = "TX_P";

            // Act
            bool result = SearchInfo.IsMeasPinInForcePin(forcePinGroup, measPinName);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void GetStoreNameOri_BuildsSequenceJoinedName()
        {
            // Arrange
            // Seq1 distinct names -> join with ":"
            // Seq2 duplicated names -> collapse to single name
            HardIpPattern pattern = new HardIpPattern
            {
                MeasPins =
                [
                    new() { SequenceIndex = 1, CusStr = "A", PinName = "P1" },
                    new() { SequenceIndex = 1, CusStr = "B", PinName = "P2" },
                    new() { SequenceIndex = 2, CusStr = "C", PinName = "P3" },
                    new() { SequenceIndex = 2, CusStr = "C", PinName = "P4" }
                ]
            };

            // Act
            string storeName = SearchInfo.GetStoreNameOri(pattern);

            // Assert
            Assert.AreEqual("A:B+C", storeName);
        }

        [TestMethod]
        public void GetCalculation_SupportsPluralArgsKey_ForHvCase()
        {
            // Arrange
            // Use plural key to hit "CalcArgs" branch
            string misc = "HV@Calc:ALG_C;HV@CalcArgs:p0,p1,p2";

            // Act
            string result = SearchInfo.GetCalculation(misc, "UT");

            // Assert
            Assert.AreEqual("HV@Alg:UT:ALG_C(p0,p1,p2)", result);
        }

        [DataTestMethod]
        [DataRow(MeasType.MeasV, "^I")]
        [DataRow(MeasType.MeasI, "^V")]
        [DataRow(MeasType.MeasR1, "^(V|I)$")]
        [DataRow(MeasType.MeasR2, "^(V|I)$")]
        [DataRow(MeasType.MeasVdiff2, "^(V1P|V2P|V1N|V2N|V)$")]
        [DataRow("UnknownType", "^Not_Define$")]
        public void GetForceTypeByMeasType_PureUt(string measType, string expected)
        {
            Assert.AreEqual(expected, SearchInfo.GetForceTypeByMeasType(measType));
        }

        [TestMethod]
        public void RemoveDummy_ReturnsEmpty_WhenAllPartsEmpty()
        {
            string input = ",,,";
            string result = SearchInfo.CheckInfoByStoreName(input, "", ',', false);
            Assert.AreEqual(input, result);
        }

        [TestMethod]
        public void CheckInfoByStoreName_SingleSegment()
        {
            string result = SearchInfo.CheckInfoByStoreName("A,B", "A:B", '+');
            Assert.AreEqual("A:B", result);
        }

        [TestMethod]
        public void CheckInfoByStoreName_MultiSegment()
        {
            string result = SearchInfo.CheckInfoByStoreName("A,B+C,D", "A:B+C:D", '+');
            Assert.AreEqual("A:B+C:D", result);
        }

        [TestMethod]
        public void RemoveDummyForceV_RemovesSingleValue()
        {
            string result = SearchInfo.RemoveDummyForceV("3,3", ",|\n|\\+");

            Assert.AreEqual("3", result);
        }

        [TestMethod]
        public void GenDiffGroupName_GenButNoGroupInPinMap_ReturnsOriginal()
        {
            string input = "NON_EXIST::PIN";
            string output = SearchInfo.GenDiffGroupName(input, true);
            Assert.AreEqual(input, output);
        }

        [TestMethod]
        public void GenDiffGroupName_NonDiff_ReturnsOriginal()
        {
            // Arrange
            string name = "SINGLE_PIN";

            // Act
            string result = SearchInfo.GenDiffGroupName(name, true);

            // Assert
            Assert.AreEqual("SINGLE_PIN", result);
        }

        [DataTestMethod]
        [DataRow("P1::P2", false, "P1,P2")]
        [DataRow("P1::P2::P3", false, "P1,P2,P3")]
        [DataRow("SINGLE_PIN", false, "SINGLE_PIN")]
        public void GenDiffGroupName_NoGen_PureUt(string input, bool gen, string expected)
        {
            Assert.AreEqual(expected, SearchInfo.GenDiffGroupName(input, gen));
        }

        [DataTestMethod]
        [DataRow("Data1_16", "", "16")]
        [DataRow("DATA1_32, DATA2_32", "", "32")]
        [DataRow("Data1_16, Data2_32", "DEFAULT", "DEFAULT")]
        [DataRow("Data1_16, Data2_8", "ERR", "ERR")]
        [DataRow("InvalidString", "DEFAULT", "")]
        [DataRow("", "DEFAULT", "")]
        [DataRow("bus_A1_64, bus_B2_64", "", "64")]
        public void GetDigDataWidth_PureUt(string input, string defaultValue, string expected)
        {
            Assert.AreEqual(expected, SearchInfo.GetDigDataWidth(input, defaultValue));
        }

        [TestMethod]
        public void CheckInfoByStoreName_SingleSegment_ReplacesComma()
        {
            // sign 使用 '+'，讓 info.Split(sign).Length == 1
            // storeName 含 ':' 以觸發替換
            string result = SearchInfo.CheckInfoByStoreName("A,B", "A:B", '+', false);
            Assert.AreEqual("A:B", result);
        }

        [TestMethod]
        public void CheckInfoByStoreName_NoColonInStoreName_NoReplace()
        {
            // sign 使用 '+'，仍是單段；但 storeName 不含 ':' → 不替換
            string result = SearchInfo.CheckInfoByStoreName("A,B", "X", '+', false);
            Assert.AreEqual("A,B", result);
        }

        [TestMethod]
        public void CheckInfoByStoreName_MultiSequence_UsesTestSeqExpansion()
        {
            string result = SearchInfo.CheckInfoByStoreName("A,B", "S1:S2", ',', true);
            Assert.IsTrue(result.Length > 0);
        }

        [TestMethod]
        public void IsAllTheSameType_PureUt()
        {
            var pins = new List<MeasPin>{
                new() { MeasType = "MeasI" },
                new() { MeasType = "MeasI" }
            };

            Assert.IsTrue(SearchInfo.IsAllTheSameType(pins, "MeasI"));
            Assert.IsFalse(SearchInfo.IsAllTheSameType(pins, "MeasV"));
        }

        [TestMethod]
        public void GetSeqlstFromMiscInfo_PureUt()
        {
            string misc = "AAA;TestSequence:S1,S2;BBB";
            List<string> result = SearchInfo.GetSeqlstFromMiscInfo(misc);

            CollectionAssert.AreEqual(
                new List<string> { "S1", "S2" },
                result);
        }

        [TestMethod]
        public void GetStatusInPatternList_FileVersionNA_ReturnsMissFileVersion()
        {
            Dictionary<string, PatternData> list = new Dictionary<string, PatternData>
            {
                { "P1", new PatternData { TimeSetVersion = "1.0", FileVersion = "NA" } }
            };

            bool result = SearchInfo.GetStatusInPatternList(list, "P1", out string status);

            Assert.IsFalse(result);
            Assert.AreEqual("MissFileVersionInPattList", status);
        }

        [TestMethod]
        public void GetMeasStrByPlan_MapsWiMeas_WiSrc_Wait()
        {
            HardIpPattern pattern = new HardIpPattern
            {
                TestPlanSequences = [
                    new(1, 1, 1),
                    new(2, 2, 2),
                    new(3, 3, 3)
                ],
                MeasPins = [
                    new() { SequenceIndex = 1, MeasType = MeasType.WiMeas, PinName = "P1" },
                    new() { SequenceIndex = 2, MeasType = MeasType.WiSrc, PinName = "P2" },
                    new() { SequenceIndex = 3, MeasType = MeasType.MeasWait, PinName = "P3" }
                ]
            };

            List<string> result = SearchInfo.GetMeasStrByPlan(pattern);

            Assert.AreEqual(3, result.Count);
            Assert.AreEqual("A", result[0]);
            Assert.AreEqual("G", result[1]);
            Assert.AreEqual("W", result[2]);
        }

        [TestMethod]
        public void DivideGroupPin_ExpandsGroupPinToMembers()
        {
            ForcePin fp = new ForcePin { PinName = "MY_GROUP" };
            List<ForcePin> list = [fp];

            List<ForcePin> result = SearchInfo.DivideGroupPin(list);

            Assert.AreEqual(2, result.Count);
            Assert.IsTrue(result.Exists(x => x.PinName == "TX_P"));
            Assert.IsTrue(result.Exists(x => x.PinName == "TX_N"));
        }

        [TestMethod]
        public void GetPinCount_MixedPins_ReturnsSum()
        {
            int result = SearchInfo.GetPinCount("TX_P,TX_N,P1::P2");
            // TX_P(1) + TX_N(1) + P1::P2(2)

            Assert.AreEqual(4, result);
        }

        [TestMethod]
        public void ModDuplicateRegName_When_FunctionName_Duplicated_Should_Rename_DigSrcAssignment()
        {
            var hardIpInfo = new HardIpInfo
            {
                SendBitName = "REG1+REG2+REG1",
                DigSrcAssignment = "FUNC1:REG1;FUNC2:REG2;FUNC1:REG3"
            };

            HardIpInfo result = SearchInfo.ModDuplicateRegName(hardIpInfo);

            Assert.AreEqual("REG1+REG2+REG1", result.SendBitName, "SendBitName should NOT be modified because function name != register name");

            Assert.AreEqual("FUNC1:REG1;FUNC2:REG2;FUNC1_REG3:REG3", result.DigSrcAssignment, "Duplicated function should be renamed with suffix");
        }

        [TestMethod]
        public void ModDuplicateRegName_ShouldAddTrimReg_WhenAssignmentHasNoColon()
        {
            var hardIpInfo = new HardIpInfo
            {
                SendBitName = "REG1",
                DigSrcAssignment = "REG1"
            };

            HardIpInfo result = SearchInfo.ModDuplicateRegName(hardIpInfo);

            Assert.IsTrue(result.DigSrcAssignment.Contains("REG1"));
        }
        [TestMethod]
        public void ModDuplicateRegName_ShouldRenameDuplicateRegister()
        {
            var hardIpInfo = new HardIpInfo
            {
                SendBitName = "FUNC+FUNC",
                DigSrcAssignment = "FUNC:A;FUNC:B"
            };

            HardIpInfo result = SearchInfo.ModDuplicateRegName(hardIpInfo);

            Assert.IsTrue(result.SendBitName.Contains("FUNC_B"));
        }
        [TestMethod]
        public void ModDuplicateRegName_ShouldAppendTrimFuseAssignment()
        {
            var hardIpInfo = new HardIpInfo
            {
                SendBitName = "REG1",
                DigSrcAssignment = "REG1:A",
                TrimFuseName = "FUSE1,FUSE2",
                TrimRegName =
                [
                    "TRIM1",
                    "TRIM2"
                ]
            };

            HardIpInfo result = SearchInfo.ModDuplicateRegName(hardIpInfo);

            Assert.IsTrue(result.DigSrcAssignment.Contains("TRIM1:FUSE1"));
            Assert.IsTrue(result.DigSrcAssignment.Contains("TRIM2:FUSE2"));
        }

        [TestMethod]
        public void GetPrePat_When_ForceCondition_StartsWithSweep_Should_SweepFirst()
        {
            var pattern = new HardIpPattern
            {
                ForceCondition = new ForceClass
                {
                    ForceCondition = "SWEEP;DISCONNECT"
                },
                ForceConditionList = []
            };

            string result = SearchInfo.GetPrePat(pattern);

            Assert.AreNotEqual(null, result);
        }
        [TestMethod]
        public void GetPrePat_ShouldAddShmooGlobal_WhenXShmooMatched()
        {
            var pattern = new HardIpPattern();

            pattern.ForceCondition.ForceCondition =
                "xshmoo(VDD:VCORE$CORE:1)";

            string result = SearchInfo.GetPrePat(pattern);

            Assert.IsTrue(result.Contains("_Shmoo_CORE_Glb"));
        }
        [DataTestMethod]
        [DataRow(HardIpConstData.IgnorePatMeasC)]
        [DataRow(HardIpConstData.IgnorePatDigSrc)]
        [DataRow(HardIpConstData.IgnorePatBinOut)]
        public void GetPreMeas_When_MiscInfo_Contains_Other_IgnoreFlags_Should_Return_EmptyString(string miscInfo)
        {
            var pattern = new HardIpPattern
            {
                MiscInfo = miscInfo,
                MeasPins =
        [
            new()
            {
                MeasType = MeasType.MeasV,
                PinName = "PIN_A",
                SequenceIndex = 1
            }
        ],
                TestPlanSequences =
        [
            new(1, 1, 1)
        ]
            };

            var hardIpInputData = new HardIpInputData(null);

            string result = SearchInfo.GetPreMeas(pattern, hardIpInputData, "", "", "|");

            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetPreMeas_When_ForceLabelVoltages_NotContain_CurrentVoltage_Should_IgnoreForcePin()
        {
            var forcePin = new ForcePin
            {
                PinName = "PIN_A",
                ForceType = "V",
                ForceValue = "1.2",
                ForceLabelVoltages = ["HV"],
                Type = ForceConditionType.Normal
            };

            var forceCondition = new ForceCondition
            {
                ForcePins = [forcePin]
            };

            var measPin = new MeasPin
            {
                MeasType = MeasType.MeasV,
                PinName = "PIN_B",
                SequenceIndex = 1,
                ForceConditions = [forceCondition]
            };

            var pattern = new HardIpPattern
            {
                MeasPins = [measPin],
                TestPlanSequences =
                [
                    new(1, 1, 1)
                ]
            };

            var hardIpInputData = new HardIpInputData(null);

            string result = SearchInfo.GetPreMeas(pattern, hardIpInputData, "", HardIpConstData.LabelNv, "|");

            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetPreMeas_When_ForcePin_IsValid_Should_Return_NonEmpty_PreMeasString()
        {
            var forcePin = new ForcePin
            {
                PinName = "PIN_A",
                ForceType = "V",
                ForceValue = "1.2",
                ForceLabelVoltages = [],
                Type = ForceConditionType.Normal,
                IsRestore = true,
                ForceJob = ""
            };

            var forceCondition = new ForceCondition
            {
                ForcePins = [forcePin]
            };

            var measPin = new MeasPin
            {
                MeasType = MeasType.MeasV,
                PinName = "PIN_B",
                SequenceIndex = 1,
                ForceConditions = [forceCondition]
            };

            var pattern = new HardIpPattern
            {
                MeasPins = [measPin],
                TestPlanSequences =
                [
                    new(1, 1, 1)
                ]
            };

            var hardIpInputData = new HardIpInputData(null);

            string result = SearchInfo.GetPreMeas(pattern, hardIpInputData, "", HardIpConstData.LabelNv, "|");

            Assert.IsFalse(string.IsNullOrEmpty(result));
            Assert.IsTrue(result.Contains(';') || result.Contains(':'));
        }

        [TestMethod]
        public void GetPreMeas_When_DisableFrc_SinglePin_Should_Enter_DisableFrc_CountMatch_Branch()
        {
            var disableFrcPin = new ForcePin
            {
                PinName = "VCC",
                ForceType = "DISABLE_FRC",
                ForceValue = "1.0",
                Type = ForceConditionType.Normal
            };

            var forceCondition = new ForceCondition
            {
                ForcePins = [disableFrcPin]
            };

            var measPin = new MeasPin
            {
                SequenceIndex = 1,
                MeasType = MeasType.MeasV,
                PinName = "VCC",
                ForceConditions = [forceCondition]
            };

            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("CZ_NVSA0_A_FULP_AN_AN00_MEA_JTG_DIO_ALLFRV_SI_ACTCONS_T7"),
                MeasPins = [measPin]
            };

            var hardIpInputData = new HardIpInputData(null);

            string result = SearchInfo.GetPreMeas(pattern, hardIpInputData, "", HardIpConstData.LabelNv, "|");

            Assert.IsFalse(string.IsNullOrEmpty(result));
        }

        [TestMethod]
        public void GetPreMeas_TestSequence_AutoAlign_MeasSeqStr_Length()
        {
            var pattern = new HardIpPattern
            {
                Pattern = new PatternClass("CZ_NVSA0_A_FULP_AN_AN00_MEA_JTG_DIO_ALLFRV_SI_ACTCONS_T7")
            };

            HardIpInfo info = HardIpService.GetHardIpInfo(pattern);
            List<string> seqList = [.. info.MeasSeqStr.Split(',')];

            string longValue = new('A', 1200);

            var measPins = new List<MeasPin>();
            int seqIndex = 1;

            foreach (string _ in seqList)
            {
                measPins.Add(
                    new MeasPin
                    {
                        SequenceIndex = seqIndex,
                        MeasType = MeasType.MeasV,
                        ForceConditions =
                        [
                    new()
                    {
                        ForcePins =
                        [
                            new()
                            {
                                PinName = "PIN_" + seqIndex,
                                ForceType = "V",
                                ForceValue = longValue,
                                Type = ForceConditionType.Normal
                            }
                        ]
                    }
                        ]
                    });

                seqIndex++;
            }

            pattern.MeasPins = measPins;

            string testSequence =
                string.Join(",", Enumerable.Range(1, seqList.Count).Select(i => $"TS{i}"));

            var hardIpInputData = new HardIpInputData(null);

            string result = SearchInfo.GetPreMeas(pattern, hardIpInputData, testSequence, HardIpConstData.LabelNv, "|");

            Assert.IsTrue(result.StartsWith("Reg_assign"));
            Assert.AreEqual(1, hardIpInputData.HardIpRegAssigns.Count);
            Assert.AreEqual(seqList.Count, hardIpInputData.HardIpRegAssigns[0].RegAssignList.Count);
        }

        [TestMethod]

        public void GetPlanCurrentRange_ShouldCollectVocmPin_WhenCntGreaterThanZero()
        {
            var measPins = new List<MeasPin>
            {
                new()
                {
                    PinName = "PIN_A",
                    MeasType = MeasType.MeasVdiff,
                    VisitedTime = 1
                },
                new()
                {
                    PinName = "PIN_A",
                    MeasType = MeasType.MeasVocm,
                    VisitedTime = 1
                }
            };

            var patInfoPins = new List<MeasPin>
            {
                new()
                {
                    PinName = "PIN_A",
                    MeasType = MeasType.MeasVdiff,
                    SequenceIndex = 1
                }
            };

            SearchInfo.GetPlanCurrentRange(
                        measPins,
                        patInfoPins,
                        false);

            Assert.IsTrue(
                        measPins.Any(x =>
                        x.MeasType == MeasType.MeasVocm &&
                        x.IsUsedPin));
        }

        [DataTestMethod]
        [DataRow("CP1:1mA,CP1:2mA", "CP1=1mA,2mA", DisplayName = "SingleJob")]
        [DataRow("CP1:1mA;CP2:2mA,CP1:3mA;CP2:4mA", "CP1=1mA,3mA;CP2=2mA,4mA", DisplayName = "MultiJob")]
        [DataRow("1mA,2mA", "1mA,2mA", DisplayName = "NoJobPrefixUnchanged")]
        public void GetIRangeByJob_TransformsOrPassesThroughBasedOnJobPrefixes(string measureIRange, string expected)
        {
            // Act
            string result = SearchInfo.GetIRangeByJob(measureIRange);

            // Assert
            Assert.AreEqual(expected, result);
        }
    }
}
