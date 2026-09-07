using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;
using Automation.Test.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.VbtLib;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class SetValueBaseTests : FunctionTestBase
    {
        public class SetValueBaseStub(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : SetValueBase(hardIpInputData, hardIpSheet)
        {

            // Implement the abstract method (can leave empty if not testing it)
            public override void SetArgsListValue(HardIpPattern hardIpPattern, ref Function function, string voltage) { }
        }

        private class TestSetValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : SetValueBase(hardIpInputData, hardIpSheet)
        {
            public override void SetArgsListValue(HardIpPattern hardIpPattern, ref Function function, string voltage)
            {
                throw new NotImplementedException();
            }
        }

        private HardIpInputData _hardIpInputData = null!;
        private HardIpSheet _hardIpSheet = null!;
        private TestSetValue _service = null!;
        private static List<Function> _functions = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            _functions = TestSuiteInitialize.Functions;
        }

        [TestInitialize]
        public void Setup()
        {
            _hardIpInputData = new HardIpInputData(new HardIpParaData(EnumBlock.HardIp));
            _hardIpSheet = new HardIpSheet
            {
                SheetName = "HardIp_1",
                PlanHeaderIdx =
                {
                    ["registerIndex"] = 5,
                    ["miscInfoIndex"] = 6,
                }
            };

            _service = new TestSetValue(_hardIpInputData, _hardIpSheet);
        }

        [DataTestMethod]
        [DataRow("TESTSEQUENCE", "VAL1,VAL2", RegisterAssignType.TestSequence, DisplayName = "01_TestSequence")]
        [DataRow("MEASSTORENAME", "STORE1|STORE2", RegisterAssignType.Meas_StoreName, DisplayName = "02_MeasStoreName")]
        [DataRow("CALCEQUNAME", "EQ1|EQ2", RegisterAssignType.CalcEquName, DisplayName = "03_CalcEquName")]
        [DataRow("CALCSTORENAME", "CALC1|CALC2", RegisterAssignType.CalcStoreName, DisplayName = "04_CalcStoreName")]
        [DataRow("MEASNAME", "M1|M2", RegisterAssignType.MeasName, DisplayName = "05_MeasName")]
        [DataRow("MEASPINS", "P1+P2", RegisterAssignType.MeasPins, DisplayName = "06_MeasPins")]
        [DataRow("CUS_STR_DIGCAPDATA", "C1:V1,C2:V2", RegisterAssignType.CUS_Str_DigCapData, DisplayName = "07_CUS_Str_DigCapData")]
        [DataRow("DIGSRC_EQUATION", "D1+D2", RegisterAssignType.DigSrc_Equation, DisplayName = "08_DigSrc_Equation")]
        [DataRow("DIGSRC_ASSIGNMENT", "X=1;Y=2", RegisterAssignType.DigSrc_Assignment, DisplayName = "09_DigSrc_Assignment")]
        [DataRow("CALC_EQN", "A:1;B:2", RegisterAssignType.Calc_Eqn, DisplayName = "10_Calc_Eqn")]
        public void CheckArgsExceedLimitationLcd_ShouldCreateRegAssign_ForAllTypes(string type, string infoParameter, RegisterAssignType registerAssignType)
        {
            // Arrange
            var function = new Function
            {
                FunctionName = "hip_X_X",
                Parameters = type + "," + type,
                ArgList = [infoParameter, new('x', 6000)]
            };

            var pattern = new HardIpPattern
            {
                SheetName = "Sheet1",
                RowNum = 1,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("P1#1") { RealPatternName = "Pat1" },
                BurstPatterns = [
                    new()
                    {
                        DigSrcEquation = "DigSrcEquation"
                    }
                ]
            };

            _hardIpInputData.HardIpRegAssigns.Clear();

            // Act
            _service.CheckArgsExceedLimitationLcd(function, pattern);

            // Assert
            Assert.AreEqual(1, _hardIpInputData.HardIpRegAssigns.Count);

            Assert.AreEqual(infoParameter, function.GetParamValue(type));
        }

        [TestMethod]
        public void CheckArgsExceedLimitationLcd_ShouldCreateRegAssign_WhenInfoParameterTooLong()
        {
            // Arrange
            TestProgram.Clear();
            TestProgram.VbtFunctionLib.AddVbtFunctionRange([.. _functions.Where(x => x.Type.EqualsIgnoreCase("VBT"))]);

            var function = new Function
            {
                Parameters = "DigSrc_Assignment",
                ArgList = [new('A', 6001), "A,B"]
            };

            var pattern = new HardIpPattern
            {
                SheetName = "Sheet1",
                RowNum = 10,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("") { RealPatternName = "Pat1" }
            };

            // Act
            _service.CheckArgsExceedLimitationLcd(function, pattern);

            // Assert
            Assert.AreEqual("Reg_assign:SHEET1_", function.ArgList[0]);
        }

        [DataTestMethod]
        [DataRow("NullParaData", true, "Sheet1", true, DisplayName = "01_Return_When_HardIpParaData_IsNull")]
        [DataRow("ClockCheck", false, "ClockCheck", true, DisplayName = "02_Return_When_SheetName_Is_ClockCheck")]
        [DataRow("MissingRegisterIdx", false, "Sheet1", false, DisplayName = "03_Return_When_RegisterIndex_Missing")]
        public void CheckArgsExceedLimitation_Should_Return_On_Outer_Early_Exit(string caseId, bool useNullInput, string sheetName, bool hasRegisterIndex)
        {
            // Arrange
            if (useNullInput)
            {
                _hardIpInputData = new HardIpInputData(null);
                _service = new TestSetValue(_hardIpInputData, _hardIpSheet);
            }

            if (!hasRegisterIndex)
            {
                _hardIpSheet.PlanHeaderIdx.Clear();
            }

            Function function = new Function
            {
                FunctionName = "hip_X_X",
                Parameters = "ANY",
                ArgList = ["A"]
            };
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = sheetName,
                RowNum = 1,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("P1")
            };

            // Act
            _service.CheckArgsExceedLimitation(function, pattern);

            // Assert
            Assert.AreEqual(0, _hardIpInputData.HardIpRegAssigns.Count, caseId);
        }

        [DataTestMethod]
        [DataRow("Read_Normal", "hip_efuse_read", "P", RegisterAssignType.HIP_eFuse_Read, "payload")]
        [DataRow("Read_WithDD", "hip_efuse_read", "dd_payload", RegisterAssignType.HIP_eFuse_Read, "dd_payload")]
        [DataRow("Write", "hip_efuse_write", "P", RegisterAssignType.HIP_eFuse_Write, "payload")]
        public void CheckArgsExceedLimitation_Efuse_Types_And_DD_Suffix(string caseId, string funcName, string lastPayload, RegisterAssignType registerAssignType, string payloadForSubBlock)
        {
            // Arrange
            Function function = new Function
            {
                FunctionName = funcName,
                Parameters = "M_CATENAME,DICTIONARYNAME",
                ArgList = ["A+B", "X+Y"]
            };
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "SheetEFUSE",
                RowNum = 10,
                MiscInfo = "MiscX",
                Pattern = new PatternClass(lastPayload) { RealPatternName = "PatEfuse" }
            };

            // Act
            _service.CheckArgsExceedLimitation(function, pattern);

            // Assert
            HardIpRegAssign? regAssign = _hardIpInputData.HardIpRegAssigns.FirstOrDefault();
            Assert.AreNotEqual(null, regAssign, caseId);
            Assert.AreEqual(registerAssignType, regAssign!.Type, caseId);
            if (payloadForSubBlock.StartsWithIgnoreCase("dd_"))
            {
                StringAssert.Contains(regAssign.SubBlockName, "_DD", caseId);
            }
        }

        [DataTestMethod]
        [DataRow("LeftLess", "A", "X+Y+Z", 3, DisplayName = "01_Efuse_CountAlign_MLess_DictMore")]
        [DataRow("LeftMore", "A+B+C", "X+Y", 3, DisplayName = "02_Efuse_CountAlign_MMore_DictLess")]
        public void CheckArgsExceedLimitation_Efuse_CountAlign(string caseId, string mcatenameArg, string dictArg, int expectedCount)
        {
            // Arrange
            Function function = new Function
            {
                FunctionName = "HardIPFuseRead",
                Parameters = "M_CATENAME,DICTIONARYNAME",
                ArgList = [mcatenameArg, dictArg]
            };
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "SheetEFUSE",
                RowNum = 1,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("P")
            };

            // Act
            _service.CheckArgsExceedLimitation(function, pattern);

            // Assert
            HardIpRegAssign ra = _hardIpInputData.HardIpRegAssigns.Single();
            Assert.AreEqual(expectedCount, ra.RegAssignList.Count, caseId);
        }

        [TestMethod]
        public void CheckArgsExceedLimitation_Efuse_NoDuplicate_When_Same_SubBlock_And_Type()
        {
            // Arrange
            Function function = new Function
            {
                FunctionName = "hip_efuse_read",
                Parameters = "M_CATENAME,DICTIONARYNAME",
                ArgList = ["A+B", "X+Y"]
            };
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "SheetEFUSE",
                RowNum = 1,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("P")
            };

            // Act
            _service.CheckArgsExceedLimitation(function, pattern);

            // Assert
            Assert.AreEqual(1, _hardIpInputData.HardIpRegAssigns.Count);
        }

        [TestMethod]
        public void CheckArgsExceedLimitation_NotEfuse_LengthLessThan6000_NoSplit()
        {
            // Arrange
            Function function = new Function
            {
                FunctionName = "hip_X_X",
                Parameters = "DIGSRC_EQUATION",
                ArgList = ["A+B"]
            };
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet1",
                RowNum = 1,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("P1")
            };

            // Act
            _service.CheckArgsExceedLimitation(function, pattern);

            // Assert
            Assert.AreEqual(0, _hardIpInputData.HardIpRegAssigns.Count);
        }

        [TestMethod]
        public void CheckArgsExceedLimitation_ShouldAddRegAssign_ForHipEfuseRead()
        {
            var function = new Function
            {
                FunctionName = "hip_efuse_read",
                Parameters = "M_CATENAME,DICTIONARYNAME",
                ArgList = ["+A+B", "Code1+Code2"]
            };

            var pattern = new HardIpPattern
            {
                SheetName = "SheetEFUSE",
                RowNum = 10,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("") { RealPatternName = "PatEfuse" }
            };

            _service.CheckArgsExceedLimitation(function, pattern);

            HardIpRegAssign? regAssign = _hardIpInputData.HardIpRegAssigns.FirstOrDefault();
            Assert.AreNotEqual(null, regAssign);
            Assert.AreEqual(RegisterAssignType.HIP_eFuse_Read, regAssign!.Type);
        }

        [TestMethod]
        public void CheckArgsExceedLimitation_ShouldAddRegAssign_NotForHipEfuseRead()
        {
            var function = new Function
            {
                FunctionName = "hip_X_X",
                Parameters = "M_CATENAME,DICTIONARYNAME",
                ArgList = ["+A+B", "Code1+Code2"]
            };

            var pattern = new HardIpPattern
            {
                SheetName = "SheetEFUSE",
                RowNum = 10,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("") { RealPatternName = "PatEfuse" }
            };

            _service.CheckArgsExceedLimitation(function, pattern);

            Assert.AreEqual("'+A+B", function.ArgList[0]);

            HardIpRegAssign? regAssign = _hardIpInputData.HardIpRegAssigns.FirstOrDefault();
            Assert.AreEqual(null, regAssign);
        }

        [DataTestMethod]
        [DataRow("CUS_STR_DIGCAPDATA", "C1:V1,C2:V2", RegisterAssignType.CUS_Str_DigCapData, DisplayName = "01_CUS_Str_DigCapData")]
        [DataRow("DIGSRC_ASSIGNMENT", "X=1;Y=2", RegisterAssignType.DigSrc_Assignment, DisplayName = "02_DigSrc_Assignment_Solo")]
        [DataRow("CALC_EQN", "A:1;B:2", RegisterAssignType.Calc_Eqn, DisplayName = "03_Calc_Eqn_Simple")]
        [DataRow("DIGSRCASSIGNMENT", "A:1;B:2", RegisterAssignType.Calc_Eqn, DisplayName = "04_DIGSRCASSIGNMENT")]
        [DataRow("MAINPROGRAMCUSTOMSTRING", "A:1;B:2", RegisterAssignType.Calc_Eqn, DisplayName = "05_MAINPROGRAMCUSTOMSTRING")]
        [DataRow("TRIMDICTIONARYSTORENAME", "A:1;B:2", RegisterAssignType.Calc_Eqn, DisplayName = "06_TRIMDICTIONARYSTORENAME")]
        public void CheckArgsExceedLimitation_ShouldCreateRegAssign_ForAllTypes(string type, string infoParameter, RegisterAssignType registerAssignType)
        {
            // Arrange
            var function = new Function
            {
                FunctionName = "hip_X_X",
                Parameters = type + "," + type,
                ArgList = [infoParameter, new('x', 6000)] // second arg forces splitting
            };

            var pattern = new HardIpPattern
            {
                SheetName = "Sheet1",
                RowNum = 1,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("P1#1") { RealPatternName = "Pat1" },
                BurstPatterns =
                [
                    new() { DigSrcEquation = "DigSrcEquation" }
                ]
            };

            _hardIpInputData.HardIpRegAssigns.Clear();

            // Act
            _service.CheckArgsExceedLimitation(function, pattern);

            // Assert
            Assert.AreEqual(1, _hardIpInputData.HardIpRegAssigns.Count);
            Assert.AreEqual(infoParameter, function.GetParamValue(type));
        }

        [DataTestMethod]
        [DataRow("01_Range_Equal", "MEASURERANGE", "S1,S2", "V1+V2", true, DisplayName = "01_MeasI_Range_EqualLength")]
        [DataRow("02_Range_Mismatch", "MEASURERANGE", "S1,S2,S3", "V1+V2", false, DisplayName = "02_MeasI_Range_Mismatch")]
        [DataRow("03_ForceV_Equal", "MEASUREFORCEV", "T1,T2", "1.0\n2.0", true, DisplayName = "03_ForceV_EqualLength")]
        [DataRow("04_ForceV_Mismatch", "MEASUREFORCEV", "T1,T2,T3", "1.0\n2.0", false, DisplayName = "04_ForceV_Mismatch")]
        [DataRow("05_ForceI_Equal", "MEASUREFORCEI", "K1,K2", "10\n20", true, DisplayName = "05_ForceI_EqualLength")]
        [DataRow("06_ForceI_Mismatch", "MEASUREFORCEI", "K1,K2,K3", "10\n20", false, DisplayName = "06_ForceI_Mismatch")]
        public void CheckArgsExceedLimitation_MeasureSequence(string caseId, string type, string measureSeq, string values, bool expectEqual)
        {
            // Arrange
            _hardIpSheet.PlanHeaderIdx["registerIndex"] = 5;
            Function function = new Function
            {
                FunctionName = "hip_X_X",
                Parameters = type,
                ArgList = [expectEqual ? values + new string('Z', 6000) : values]
            };

            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet1",
                RowNum = 1,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("P1")
            };

            // Act
            _service.CheckArgsExceedLimitation(function, pattern);

            // Assert
            if (expectEqual)
            {
                Assert.AreEqual(1, _hardIpInputData.HardIpRegAssigns.Count, caseId);
            }
            else
            {
                Assert.AreEqual(0, _hardIpInputData.HardIpRegAssigns.Count, caseId);
            }
        }

        [TestMethod]
        public void CheckArgsExceedLimitation_CalcEqn_With_Alg_And_Normal()
        {
            // Arrange
            _hardIpSheet.PlanHeaderIdx["registerIndex"] = 5;
            Function function = new Function
            {
                FunctionName = "hip_X_X",
                Parameters = "CALC_EQN",
                ArgList = [""]
            };
            function.ArgList[0] = "Alg::MyAlg;M1:X1;M2:X2" + new string('X', 6000);

            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet1",
                RowNum = 1,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("P1")
            };

            // Act
            _service.CheckArgsExceedLimitation(function, pattern);

            // Assert
            HardIpRegAssign? ra = _hardIpInputData.HardIpRegAssigns.FirstOrDefault();
            Assert.AreNotEqual(null, ra);
            Assert.AreEqual(RegisterAssignType.Calc_Eqn, ra!.Type);
            Assert.IsTrue(ra.RegAssignList.Count >= 3);
        }

        [TestMethod]
        public void CheckArgsExceedLimitation_UnsupportedType_GeneratesError()
        {
            // Arrange
            _hardIpSheet.PlanHeaderIdx["registerIndex"] = 5;
            Function function = new Function
            {
                FunctionName = "hip_X_X",
                Parameters = "UNSUPPORTED_TYPE",
                ArgList = [new('U', 6000)]
            };
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet1",
                RowNum = 1,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("P1")
            };

            // Act
            _service.CheckArgsExceedLimitation(function, pattern);

            // Assert
            HardIpRegAssign? ra = _hardIpInputData.HardIpRegAssigns.FirstOrDefault();
            Assert.AreNotEqual(null, ra);
            Assert.AreEqual(RegisterAssignType.NoneType, ra!.Type);
        }

        [DataTestMethod]
        [DataRow("01_NullArg_ToEmpty", null, DisplayName = "01_ArgList_Null_To_EmptyString")]
        [DataRow("02_PlusPrefixed_Quote", "+ABC", DisplayName = "02_ArgList_PlusPrefixed_Adds_Quote")]
        public void CheckArgsExceedLimitation_Robust_Null_And_PlusPrefix(string caseId, string rawArg)
        {
            // Arrange
            _hardIpSheet.PlanHeaderIdx["registerIndex"] = 5;
            Function function = new Function
            {
                FunctionName = "hip_X_X",
                Parameters = "DIGSRC_EQUATION,DIGSRC_EQUATION",
                ArgList = [rawArg, new('X', 6000)]
            };

            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet1",
                RowNum = 1,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("P1")
            };

            // Act
            _service.CheckArgsExceedLimitation(function, pattern);

            // Assert
            Assert.IsTrue(_hardIpInputData.HardIpRegAssigns.Count != 0, caseId);
            Assert.AreEqual(RegisterAssignType.DigSrc_Equation, _hardIpInputData.HardIpRegAssigns.First().Type, caseId);
        }

        [DataTestMethod]
        [DataRow("01_Multiple_Equ_Mixed", "DIGSRC_EQUATION", "D1+D2", "", "", "", DisplayName = "01_IsMultiple_DigSrcEquation_Burst_Mixed")]
        [DataRow("02_Multiple_Assign_Mixed", "DIGSRC_ASSIGNMENT", "", "", "A=1;B=2", "", DisplayName = "02_IsMultiple_DigSrcAssignment_Burst_Mixed")]
        public void CheckArgsExceedLimitation_Multiple_Burst_Split(string caseId, string type, string burst1Equation, string burst2Equation, string burst1Assign, string burst2Assign)
        {
            // Arrange
            _hardIpSheet.PlanHeaderIdx["registerIndex"] = 5;
            Function function = new Function
            {
                FunctionName = "hip_X_X",
                Parameters = type,
                ArgList = [new('Q', 6000)]
            };
            HardIpPattern burstWith = new HardIpPattern
            {
                Pattern = new PatternClass("X"),
                DigSrcEquation = burst1Equation,
                RegisterAssignment = burst1Assign
            };
            HardIpPattern burstEmpty = new HardIpPattern
            {
                Pattern = new PatternClass("Y"),
                DigSrcEquation = burst2Equation,
                RegisterAssignment = burst2Assign
            };
            HardIpPattern pattern = new HardIpPattern
            {
                SheetName = "Sheet1",
                RowNum = 1,
                MiscInfo = "MiscX",
                Pattern = new PatternClass("P1#2"),
                BurstPatterns = [burstWith, burstEmpty]
            };

            // Act
            _service.CheckArgsExceedLimitation(function, pattern);

            // Assert
            string joined = function.ArgList[0];
            string[] parts = joined.Split('\n');
            bool splitAsBurst = joined.Contains('\n');

            if (splitAsBurst)
            {
                Assert.AreEqual(2, parts.Length, caseId);
                StringAssert.StartsWith(parts[0], "Reg_assign:", caseId);
                Assert.AreEqual(string.Empty, parts[1], caseId);
                Assert.IsTrue(_hardIpInputData.HardIpRegAssigns.Count != 0, caseId);
            }
            else
            {
                Assert.AreEqual(1, parts.Length, caseId);
                StringAssert.StartsWith(parts[0], "Reg_assign:", caseId);
                Assert.IsTrue(_hardIpInputData.HardIpRegAssigns.Count != 0, caseId);
            }
        }

        [TestMethod]
        [DataRow("InterposePreInit", InterposeAssignType.InterposePrePat)]
        [DataRow("InterposePreRst", InterposeAssignType.InterposePostPat)]
        [DataRow("InterposePreMeas", InterposeAssignType.InterposePreMeas)]
        [DataRow("InterposePostMeas", InterposeAssignType.InterposePostMeas)]
        public void SetInterPoseFunc_ShouldAddInterposeAssign_WhenVoltageIsNV(string text, InterposeAssignType interposeAssignType)
        {
            var hardIpInputData = new HardIpInputData(new HardIpParaData(EnumBlock.HardIp));
            var hardIpSheet = new HardIpSheet
            {
                SheetName = "HardIp_1",
                PlanHeaderIdx =
                {
                    ["miscInfoIndex"] = 2,
                    ["registerIndex"] = 5,
                    ["miscInfoIndex"] = 6,
                }
            };

            var service = new TestSetValue(hardIpInputData, hardIpSheet);

            var function = new Function
            {
                FunctionName = "FunctionalTest",
                Parameters = text + "," + text,
                ArgList = ["", ""]
            };

            var pattern = new HardIpPattern
            {
                SheetName = "SheetNV",
                RowNum = 5,
                MiscInfo = $"{text}:A;other",
                Pattern = new PatternClass("") { RealPatternName = "PAT_NV" }
            };

            service.SetInterPoseFunc(pattern, function, "NV");

            Assert.IsFalse(hardIpInputData.InterposeAssigns.IsEmpty);
            InterposeAssign assign = hardIpInputData.InterposeAssigns.First();
            Assert.IsTrue(assign.AssignName.Length > 0);
            Assert.AreEqual(interposeAssignType, assign.Type);
            Assert.IsTrue(assign.InterposeAssignList.Count > 0);
        }

        [TestMethod]
        [DataRow("InterposePreInit", InterposeAssignType.StartOfBodyFArgs, DisplayName = "InterposePreInit")]
        [DataRow("InterposePreRst", InterposeAssignType.EndOfBodyFArgs, DisplayName = "InterposePreRst")]
        [DataRow("InterposePreMeas", InterposeAssignType.InterposePreMeas, DisplayName = "InterposePreMeas")]
        [DataRow("InterposePostMeas", InterposeAssignType.InterposePostMeas, DisplayName = "InterposePostMeas")]
        public void SetInterPoseFunc_ShouldAddInterposeAssign_WhenVoltageIsNV_1(string text, InterposeAssignType interposeAssignType)
        {
            var hardIpInputData = new HardIpInputData(new HardIpParaData(EnumBlock.HardIp))
            {
                InterposeAssigns =
                [
                    new() { AssignName = "SHEETNV__PostMeas", BlockName = "SHEETNV_", Type = InterposeAssignType.InterposePreInit }
                ]
            };
            var hardIpSheet = new HardIpSheet
            {
                SheetName = "HardIp_1",
                PlanHeaderIdx =
                {
                    ["miscInfoIndex"] = 2,
                    ["registerIndex"] = 5,
                    ["miscInfoIndex"] = 6,
                }
            };

            var service = new TestSetValue(hardIpInputData, hardIpSheet);

            var function = new Function
            {
                FunctionName = "functional_t_updated",
                Parameters = text + "," + text,
                ArgList = ["", ""]
            };

            var pattern = new HardIpPattern
            {
                SheetName = "SheetNV",
                RowNum = 5,
                MiscInfo = $"{text}:A;other",
                Pattern = new PatternClass("") { RealPatternName = "PAT_NV" }
            };

            service.SetInterPoseFunc(pattern, function, "NV");

            Assert.IsFalse(hardIpInputData.InterposeAssigns.IsEmpty);
            InterposeAssign assign = hardIpInputData.InterposeAssigns.First();
            Assert.IsTrue(assign.AssignName.Length > 0);
            Assert.AreEqual(interposeAssignType, assign.Type);
            Assert.IsTrue(assign.InterposeAssignList.Count > 0);
        }

        [TestMethod]
        public void GetPreMeas_ShouldReturnRegAssign_WhenResultTooLong()
        {
            var pattern = new HardIpPattern
            {
                SheetName = "SheetPM",
                RowNum = 1,
                MiscInfo = "MiscPM",
                Pattern = new PatternClass("dp_brna0_c_fulp_an_aa22_frq_jtg_dyc_allfrv_si_ciopll_t2p1r") { RealPatternName = "PAYLOAD_PM" },
                MeasPins =
                [
                    new()
                    {
                        MeasType = "WiMeas",
                        PinName = "PIN1",
                        SequenceIndex = 1,
                        ForceConditions =
                        [
                            new() { ForcePins = [new() { PinName = "PIN1", ForceValue = new string('X', 6000) }] }
                        ]
                    }
                ],
                TestPlanSequences = [new(1, 1, 1)]
            };

            string result = SearchInfo.GetPreMeas(pattern, _hardIpInputData);

            Assert.IsTrue(result.StartsWith("Reg_assign:"));
            Assert.IsTrue(_hardIpInputData.HardIpRegAssigns.Any(x => x.Type == RegisterAssignType.Interpose_PreMeas));
        }

        [TestMethod]
        public void IsCapBitsMatch_MatchingBits_ReturnsTrue()
        {
            var patternItem = new HardIpPattern
            {
                MeasPins =
                [
                    new() { MeasType = MeasType.MeasC, CapBit = "8" },
                    new() { MeasType = MeasType.MeasC, CapBit = "16" }
                ]
            };

            var patInfo = new HardIpInfo
            {
                CapBitStr = "C_8+C_16"
            };

            bool result = SearchInfo.IsCapBitsMatch(patternItem, patInfo);

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsCapBitsMatch_DifferentBits_ReturnsFalse()
        {
            var patternItem = new HardIpPattern
            {
                MeasPins =
                [
                    new() { MeasType = MeasType.MeasC, CapBit = "8" },
                    new() { MeasType = MeasType.MeasC, CapBit = "15" }
                ]
            };

            var patInfo = new HardIpInfo
            {
                CapBitStr = "C_8+C_16"
            };

            bool result = SearchInfo.IsCapBitsMatch(patternItem, patInfo);

            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsCapBitsMatch_DifferentCounts_ReturnsFalse()
        {
            var patternItem = new HardIpPattern
            {
                MeasPins =
                [
                    new() { MeasType = MeasType.MeasC, CapBit = "8" }
                ]
            };

            var patInfo = new HardIpInfo
            {
                CapBitStr = "C_8+C_16"
            };

            bool result = SearchInfo.IsCapBitsMatch(patternItem, patInfo);

            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsCapBitsMatch_EmptyCapBitStr_ReturnsTrueIfNoMeasC()
        {
            var patternItem = new HardIpPattern
            {
                MeasPins = []
            };

            var patInfo = new HardIpInfo
            {
                CapBitStr = ""
            };

            bool result = SearchInfo.IsCapBitsMatch(patternItem, patInfo);

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsCapBitsMatch_EmptyCapBitStr_ReturnsFalseIfMeasCExist()
        {
            var patternItem = new HardIpPattern
            {
                MeasPins =
                [
                    new() { MeasType = MeasType.MeasC, CapBit = "8" }
                ]
            };

            var patInfo = new HardIpInfo
            {
                CapBitStr = ""
            };

            bool result = SearchInfo.IsCapBitsMatch(patternItem, patInfo);

            Assert.IsFalse(result);
        }

        [TestMethod]
        public void SetValueByParamMapping_EfusePrewrite_AddsFlagToMiscDict()
        {
            // Arrange
            ErrorReportManager.ClearErrors();
            var pattern = new HardIpPattern
            {
                FunctionName = "hip_efuse_write",
                SheetName = "DCTEST_IDS",
                HipPreWriteFlag = "JUDGE_FLAG_01",
                MiscInfoDict = [] // Empty
            };
            var function = new Function { Parameters = "", ArgList = [] };

            // Act
            new SetValueBaseStub(_hardIpInputData, _hardIpSheet).SetValueByParamMapping(function, pattern, "ANY_VOLT");

            // Assert
            Assert.AreEqual(4, ErrorReportManager.GetErrorList().Count);
        }

        [TestMethod]
        [DataRow("HIP_eFuse_Read", "", true, 1)]
        [DataRow("HIP_eFuse_Read", "m_catename", false, 2)]
        [DataRow("HIP_eFuse_Read", "DSPWAVESIZE", false, 2)]
        [DataRow("HIP_eFuse_Read", "m_catename,DSPWAVESIZE", false, 3)]
        [DataRow("HIP_eFuse_Write", "", true, 4)]
        [DataRow("HIP_eFuse_Write", "m_catename", false, 5)]
        [DataRow("HIP_eFuse_Write", "DSPWAVESIZE", false, 5)]
        [DataRow("HIP_eFuse_Write", "m_catename,DSPWAVESIZE", false, 6)]
        public void SetValueByParamMapping_EfuseRead_MissingBitDefTable_AddsError_1(string functionName, string keys, bool efuseBitDefBwnull, int expected)
        {
            // Arrange
            ErrorReportManager.ClearErrors();
            var pattern = new HardIpPattern
            {
                FunctionName = functionName,
                MiscInfoDict = []
            };
            foreach (string key in keys.Split(','))
            {
                pattern.MiscInfoDict[key] = "TEST";
            }

            _hardIpInputData.EfuseBitDefBw = efuseBitDefBwnull ? null : [];
            var function = new Function { Parameters = "" };

            // Act
            new SetValueBaseStub(_hardIpInputData, _hardIpSheet).SetValueByParamMapping(function, pattern, "NV");

            // Assert
            Assert.AreEqual(expected, ErrorReportManager.GetErrorList().Count);
        }

        [TestMethod]
        public void SetValueByParamMapping_ValidMapping_UpdatesArgList()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                FunctionName = "OtherFunc",
                MiscInfoDict = new Dictionary<string, string> { { "CUS_Str_MainProgram", "Val1;Val2" } }
            };

            var function = new Function
            {
                Parameters = "PARAM1,CUS_Str_MainProgram,PARAM3",
                ArgList = ["orig1", "", "orig3"]
            };

            // Act
            new SetValueBaseStub(_hardIpInputData, _hardIpSheet).SetValueByParamMapping(function, pattern, "NV");

            // Assert: Index of CUS_Str_MainProgram is 1
            Assert.AreEqual("Val1;Val2", function.ArgList[1]);
        }

        [TestMethod]
        public void SetValueByParamMapping_ValidMapping_UpdatesArgList_1()
        {
            // Arrange
            LocalSpecs.Options.Device = EnumDevice.RF;
            var pattern = new HardIpPattern
            {
                FunctionName = "OtherFunc",
                MiscInfoDict = new Dictionary<string, string> { { "MeasName", "Val1_2_3_4_5_6_7_8_9;Val2;" } }
            };

            var function = new Function
            {
                Parameters = "PARAM1,MeasName,PARAM3",
                ArgList = ["orig1", "", "orig3"]
            };

            // Act
            new SetValueBaseStub(_hardIpInputData, _hardIpSheet).SetValueByParamMapping(function, pattern, "NV");

            // Assert: Index of CUS_Str_MainProgram is 1
            Assert.AreEqual("Val1_2_3_4_5_6_7_8_NV|Val2|", function.ArgList[1]);
            LocalSpecs.Options.Device = EnumDevice.AP;
        }

        [TestMethod]
        public void SetValueByParamMapping_ValidMapping_UpdatesArgList_2()
        {
            // Arrange
            LocalSpecs.Options.Device = EnumDevice.RF;
            var pattern = new HardIpPattern
            {
                MiscInfoDict = new Dictionary<string, string> { { "XXX", "Val1;Val2" } }
            };

            var function = new Function
            {
                FunctionName = "meas_freqvoltcurr_universal_func",
                Parameters = "PARAM1,XXX,PARAM3",
                ArgList = ["orig1", "", "orig3"]
            };

            // Act
            new SetValueBaseStub(_hardIpInputData, _hardIpSheet).SetValueByParamMapping(function, pattern, "NV");

            // Assert
            Assert.AreEqual("Val1;Val2", function.ArgList[1]);
            LocalSpecs.Options.Device = EnumDevice.AP;
        }

        [TestMethod]
        public void CheckInstArgument_ValueUnderLimit_NoChange()
        {
            // Arrange
            var function = new Function
            {
                Parameters = "Param1",
                ArgList = [new('A', 5000)]
            };
            var pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                RowNum = 10,
                Pattern = new PatternClass("pat")
            };

            // Act
            new SetValueBaseStub(_hardIpInputData, _hardIpSheet).CheckInstArgument(function, pattern);

            // Assert
            Assert.AreEqual(5000, function.ArgList[0].Length);
        }

        [TestMethod]
        public void CheckInstArgument_ValueOverLimit_ClipsTo8000()
        {
            // Arrange
            string longString = new('B', 8500);
            var function = new Function
            {
                Parameters = "LongParam",
                ArgList = [longString]
            };
            var pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                RowNum = 10,
                Pattern = new PatternClass("pat")
            };

            // Act
            new SetValueBaseStub(_hardIpInputData, _hardIpSheet).CheckInstArgument(function, pattern);

            // Assert
            Assert.AreEqual(8000, function.ArgList[0].Length);
            Assert.AreEqual(longString[..8000], function.ArgList[0]);
        }

        [TestMethod]
        public void CheckInstArgument_MultiTimeDomain_SkipsCheck()
        {
            // Arrange
            var function = new Function
            {
                Parameters = "Param1",
                ArgList = [new('C', 9000)]
            };
            // Set IsMultiTimeDomain to true
            var pattern = new HardIpPattern
            {
                SheetName = "TestSheet",
                RowNum = 10,
                Pattern = new PatternClass("pat")
                {
                    RealPatternName = "P1#P2"
                }
            };

            // Act
            new SetValueBaseStub(_hardIpInputData, _hardIpSheet).CheckInstArgument(function, pattern);

            // Assert
            Assert.AreEqual(9000, function.ArgList[0].Length);
        }
    }
}
