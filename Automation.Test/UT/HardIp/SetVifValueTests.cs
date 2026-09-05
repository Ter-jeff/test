using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;

using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class SetVifValueTests
    {
        private SetVifValue _setVifValue = null!;
        private HardIpInputData _inputData = null!;

        [TestInitialize]
        public void Setup()
        {
            HardIpParaData paraData = new HardIpParaData(EnumBlock.HardIp);
            var dummySheet = new HardIpSheet();
            _inputData = new HardIpInputData(paraData) { HardIpRegAssigns = [] };
            _setVifValue = new SetVifValue(_inputData, dummySheet);
        }

        private static HardIpPattern CreatePattern(string lastPayload, string sheetName = "Sheet1", string miscInfo = "Misc")
        {
            var pattern = new HardIpPattern
            {
                SheetName = sheetName,
                MiscInfo = miscInfo,
                Pattern = new PatternClass("")
                {
                    PatternSetList = [[lastPayload]]
                }
            };
            return pattern;
        }

        [TestMethod]
        [DataRow("V,FDIFF,R1,IDIFF,R2,VDIFF2", "V,F,R,I,Z,VDIFF2", DisplayName = "01_GetMeasSeq_ReplacementsWork")]
        [DataRow("FDIFF", "F", DisplayName = "02_GetMeasSeq_FDIFF_Replaced")]
        [DataRow("VDIFF,IDIFF,R2", "V,I,Z", DisplayName = "03_GetMeasSeq_MultipleReplacements")]
        public void GetMeasSeq_ShouldReplaceSequencesCorrectly(string input, string expected)
        {
            // Act
            string result = _setVifValue.GetMeasSeq(input);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow("1:2+3", "@1:2+3", DisplayName = "01_CheckAddSymbol_AddsAtSymbolForRange")]
        [DataRow("2+4", "@2+4", DisplayName = "02_CheckAddSymbol_AddsAtSymbolForAddition")]
        [DataRow("2:4", "@2:4", DisplayName = "03_CheckAddSymbol_AddsAtSymbolForColon")]
        [DataRow("A+B", "A+B", DisplayName = "04_CheckAddSymbol_NoAtSymbolForAlphaExpression")]
        [DataRow("5", "5", DisplayName = "05_CheckAddSymbol_NoAtSymbolForSingleNumber")]
        public void CheckAddSymbol_ShouldAddAtSymbolWhenNumericExpression(string input, string expected)
        {
            // Act
            string result = _setVifValue.CheckAddSymbol(input);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow("F,I,V", "F", "0.01++", DisplayName = "01_GetWaitTime_Found_F")]
        [DataRow("F,V,I", "I", "++0.01", DisplayName = "02_GetWaitTime_Found_I")]
        [DataRow("A,B,C", "V", "", DisplayName = "03_GetWaitTime_SeqNotFound")]
        public void Get_WaitTime_ShouldReturnCorrectPattern(string allSeq, string speSeq, string expected)
        {
            // Act
            string result = _setVifValue.GetWaitTime(allSeq, speSeq);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetInstSpecialSetup_ShouldReturnNumericValue()
        {
            string result = _setVifValue.GetInstSpecialSetup("INST_SPECIAL_SETUP:5");
            Assert.AreEqual("5", result);
        }

        [TestMethod]
        public void GetInstSpecialSetup_ShouldResolveFromDictionary()
        {
            HardIpParaData paraData = new HardIpParaData(EnumBlock.HardIp);
            var inputData = new HardIpInputData(paraData);
            var setVifValue = new SetVifValue(inputData, new HardIpSheet());

            string result = setVifValue.GetInstSpecialSetup("INST_SPECIAL_SETUP:KEY_X");

            Assert.AreEqual("KeyNotFoundInSetting", result);
        }

        [TestMethod]
        public void GetInstSpecialSetup_ShouldReturnKeyNotFound()
        {
            string result = _setVifValue.GetInstSpecialSetup("INST_SPECIAL_SETUP:UNKNOWN_KEY");
            Assert.AreEqual("KeyNotFoundInSetting", result);
        }

        [TestMethod]
        [DataRow("dd_payload", "Seq1,Seq2", "PIN1,PIN2", "MeasV", DisplayName = "01_WriteMeasPinToRegAssign_Appends_DD_Suffix")]
        [DataRow("normal_payload", "T1,T2,T3", "A,B,C", "MeasI", DisplayName = "02_WriteMeasPinToRegAssign_MeasI_Type_Assigned")]
        [DataRow("payloadX", "T1,T2", "X,Y", "NoneType", DisplayName = "03_WriteMeasPinToRegAssign_MeasF_SingleEnd_Type_Assigned")]
        public void WriteMeasPinToRegAssign_ShouldCreateCorrectRegAssign(string lastPayload, string testSeq, string pins, string type)
        {
            // Arrange
            HardIpPattern pattern = CreatePattern(lastPayload);
            List<string> pinList = [.. pins.Split(',')];

            // Act
            string result = _setVifValue.WriteMeasPinToRegAssign(pattern, testSeq, pinList, type);

            // Assert
            StringAssert.StartsWith(result, "Reg_assign:");

            HardIpRegAssign? regAssign = _inputData.HardIpRegAssigns.FirstOrDefault();
            Assert.AreNotEqual(null, regAssign, "HardIpRegAssign should have been added");
            Assert.IsTrue(regAssign!.SubBlockName.Contains("SHEET1"), "SubBlockName should contain SheetName part");

            if (lastPayload.StartsWithIgnoreCase("dd_"))
            {
                StringAssert.Contains(regAssign.SubBlockName, "_DD", "Should append '_DD' when payload starts with dd_");
            }

            switch (type)
            {
                case "MeasV":
                    Assert.AreEqual(RegisterAssignType.MeasV_PinS, regAssign.Type);
                    break;
                case "MeasI":
                    Assert.AreEqual(RegisterAssignType.MeasI_PinS, regAssign.Type);
                    break;
                case "MeasF_PinS_SingleEnd":
                    Assert.AreEqual(RegisterAssignType.MeasF_PinS_SingleEnd, regAssign.Type);
                    break;
            }

            Assert.AreEqual(pinList.Count, regAssign.RegAssignList.Count, "Each pin should have matching test sequence data");
        }

        [TestMethod]
        [DataRow("payloadZ", "Seq1,Seq2,Seq3", "PIN1,PIN2", "MeasV", DisplayName = "04_WriteMeasPinToRegAssign_MismatchedCounts_NoAssignAdded")]
        public void WriteMeasPinToRegAssign_ShouldNotAdd_WhenCountsMismatch(string lastPayload, string testSeq, string pins, string type)
        {
            // Arrange
            HardIpPattern pattern = CreatePattern(lastPayload);
            List<string> pinList = [.. pins.Split(',')];

            // Act
            string result = _setVifValue.WriteMeasPinToRegAssign(pattern, testSeq, pinList, type);

            // Assert
            Assert.AreEqual(1, _inputData.HardIpRegAssigns.Count, "Should not add when sequence/pin counts differ");
            StringAssert.StartsWith(result, "Reg_assign:");
        }

        [TestMethod]
        [DataRow("payload", "S1", "P1", "MeasI", DisplayName = "05_WriteMeasPinToRegAssign_DuplicateNotAdded")]
        public void WriteMeasPinToRegAssign_ShouldNotAddDuplicate(string lastPayload, string testSeq, string pins, string type)
        {
            // Arrange
            HardIpPattern pattern = CreatePattern(lastPayload);
            List<string> pinList = [.. pins.Split(',')];

            _setVifValue.WriteMeasPinToRegAssign(pattern, testSeq, pinList, type);
            int countAfterFirst = _inputData.HardIpRegAssigns.Count;

            _setVifValue.WriteMeasPinToRegAssign(pattern, testSeq, pinList, type);
            int countAfterSecond = _inputData.HardIpRegAssigns.Count;

            // Assert
            Assert.AreEqual(countAfterFirst, countAfterSecond, "Duplicate SubBlockName+Type should not be added again");
        }

        [TestMethod]
        public void RearrangePin_DecomposedGroupsMatch_ReturnsForcePins()
        {
            // Arrange
            var pins = new List<MeasPin>
            {
                new()
                {
                    PinName = "MEAS_GRP",
                    ForceConditions = [
                        new()
                        {
                            ForcePins = [new() { PinName = "FORCE_GRP" }]
                        }
                    ]
                }
            };

            var expectedReturn = new List<string> { "MEAS_GRP" };
            // Act
            List<string> result = _setVifValue.RearrangePin(pins);

            // Assert
            CollectionAssert.AreEqual(expectedReturn, result);
        }
    }
}
