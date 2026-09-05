using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility;
using Automation.GenerateIgxl.HardIp.InputObject;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{

    [TestClass]
    public class MiscInfoPreProcessorTests
    {
        private HardIpSheet _sheet = null!;
        private Dictionary<string, HardIpSheet> _planDic = null!;

        [TestInitialize]
        public void Setup()
        {
            _sheet = new HardIpSheet { Rows = [] };
            _planDic = new Dictionary<string, HardIpSheet> { { "Sheet1", _sheet } };
        }

        [TestMethod]
        public void GetTempOpcode_SingleRow_AddsIfAndEndIf()
        {
            // Arrange
            var row = new HardIpPattern { MiscInfo = "" };
            row.MiscInfoDict.Add(HardIpConstData.HighTemp, "SomeValue");
            _sheet.Rows.Add(row);

            // Act
            MiscInfoPreProcessor.PreProcessMiscInfo(_planDic);

            // Assert
            // Logic: !last (true) -> adds If; i+1 < count (false) -> adds EndIf
            Assert.IsTrue(row.MiscInfo.Contains("Opcode:if:Temp_85C"));
            Assert.IsTrue(row.MiscInfo.Contains("Opcode:EndIf"));
        }

        [TestMethod]
        public void GetTempOpcode_ConsecutiveRows_WrapsBlock()
        {
            // Arrange
            var row1 = new HardIpPattern { MiscInfo = "" };
            row1.MiscInfoDict.Add(HardIpConstData.HighTemp, "Value");

            var row2 = new HardIpPattern { MiscInfo = "" };
            row2.MiscInfoDict.Add(HardIpConstData.HighTemp, "Value");

            _sheet.Rows.Add(row1);
            _sheet.Rows.Add(row2);

            // Act
            MiscInfoPreProcessor.PreProcessMiscInfo(_planDic);

            // Assert
            // Row 1 should have If, but NO EndIf because next row also has the key
            Assert.IsTrue(row1.MiscInfo.Contains("Opcode:if:Temp_85C"));
            Assert.IsFalse(row1.MiscInfo.Contains("Opcode:EndIf"));

            // Row 2 should NOT have If (last was true), but SHOULD have EndIf (end of list)
            Assert.IsFalse(row2.MiscInfo.Contains("Opcode:if:Temp_85C"));
            Assert.IsTrue(row2.MiscInfo.Contains("Opcode:EndIf"));
        }

        [TestMethod]
        public void GetCondPatternOpcode_DifferentConditions_ClosesAndOpensNewIf()
        {
            // Arrange
            var row1 = new HardIpPattern { MiscInfo = "" };
            row1.MiscInfoDict.Add(HardIpConstData.ExecCond, "CondA");

            var row2 = new HardIpPattern { MiscInfo = "" };
            row2.MiscInfoDict.Add(HardIpConstData.ExecCond, "CondB");

            _sheet.Rows.Add(row1);
            _sheet.Rows.Add(row2);

            // Act
            MiscInfoPreProcessor.PreProcessMiscInfo(_planDic);

            // Assert
            Assert.AreEqual("Opcode:if:CondA;Opcode:EndIf;", row1.MiscInfo);
            Assert.AreEqual("Opcode:EndIf;", row2.MiscInfo);
        }

        [TestMethod]
        public void GetTempOpcode_MiddleOfRun_NoOpcodeAdded()
        {
            // Arrange - three consecutive rows all carrying the key; the middle row is
            // neither the start (last==true) nor the end (next==true) of the run
            var row1 = new HardIpPattern { MiscInfo = "" };
            row1.MiscInfoDict.Add(HardIpConstData.HighTemp, "Value");
            var row2 = new HardIpPattern { MiscInfo = "" };
            row2.MiscInfoDict.Add(HardIpConstData.HighTemp, "Value");
            var row3 = new HardIpPattern { MiscInfo = "" };
            row3.MiscInfoDict.Add(HardIpConstData.HighTemp, "Value");
            _sheet.Rows.Add(row1);
            _sheet.Rows.Add(row2);
            _sheet.Rows.Add(row3);

            // Act
            MiscInfoPreProcessor.PreProcessMiscInfo(_planDic);

            // Assert
            Assert.AreEqual("", row2.MiscInfo);
        }

        [TestMethod]
        public void GetTempOpcode_EndOfRunWithMoreRowsFollowing_AddsEndIfOnlyThenNothingAfter()
        {
            // Arrange - run of two matching rows followed by a non-matching row
            var row1 = new HardIpPattern { MiscInfo = "" };
            row1.MiscInfoDict.Add(HardIpConstData.HighTemp, "Value");
            var row2 = new HardIpPattern { MiscInfo = "" };
            row2.MiscInfoDict.Add(HardIpConstData.HighTemp, "Value");
            var row3 = new HardIpPattern { MiscInfo = "" };
            _sheet.Rows.Add(row1);
            _sheet.Rows.Add(row2);
            _sheet.Rows.Add(row3);

            // Act
            MiscInfoPreProcessor.PreProcessMiscInfo(_planDic);

            // Assert
            Assert.IsFalse(row2.MiscInfo.Contains("Opcode:if:Temp_85C"));
            Assert.IsTrue(row2.MiscInfo.Contains("Opcode:EndIf"));
            Assert.AreEqual("", row3.MiscInfo);
        }

        [TestMethod]
        public void GetTempOpcode_RowWithoutKey_NothingAdded()
        {
            // Arrange
            var row = new HardIpPattern { MiscInfo = "" };
            _sheet.Rows.Add(row);

            // Act
            MiscInfoPreProcessor.PreProcessMiscInfo(_planDic);

            // Assert
            Assert.AreEqual("", row.MiscInfo);
        }

        [TestMethod]
        public void GetCondPatternOpcode_SameExecCondAsPrevious_NoOpcodeAddedOnRepeat()
        {
            // Arrange - once execCond matches lastExecCond, the guard suppresses further
            // opcode emission for every subsequent row sharing that same value
            var row1 = new HardIpPattern { MiscInfo = "" };
            row1.MiscInfoDict.Add(HardIpConstData.ExecCond, "CondA");
            var row2 = new HardIpPattern { MiscInfo = "" };
            row2.MiscInfoDict.Add(HardIpConstData.ExecCond, "CondA");
            _sheet.Rows.Add(row1);
            _sheet.Rows.Add(row2);

            // Act
            MiscInfoPreProcessor.PreProcessMiscInfo(_planDic);

            // Assert
            Assert.IsTrue(row1.MiscInfo.Contains("Opcode:if:CondA"));
            Assert.AreEqual("", row2.MiscInfo);
        }

        [TestMethod]
        public void GetCondPatternOpcode_NextRowHasDifferentExecCond_StillClosesCurrentIf()
        {
            // Arrange - next row carries the key too, but with a different value, so the
            // current block must still be closed even though `next` (presence) is true
            var row1 = new HardIpPattern { MiscInfo = "" };
            row1.MiscInfoDict.Add(HardIpConstData.ExecCond, "CondA");
            var row2 = new HardIpPattern { MiscInfo = "" };
            row2.MiscInfoDict.Add(HardIpConstData.ExecCond, "CondB");
            _sheet.Rows.Add(row1);
            _sheet.Rows.Add(row2);

            // Act
            MiscInfoPreProcessor.PreProcessMiscInfo(_planDic);

            // Assert
            Assert.AreEqual("Opcode:if:CondA;Opcode:EndIf;", row1.MiscInfo);
        }

        [TestMethod]
        public void GetCondPatternOpcode_ValueEqualsConditionKeyItself_TreatedAsNoExecCond()
        {
            // Arrange - GetExecCond treats a value equal to the condition key itself as "no
            // override", so no opcode should be emitted despite the key being present
            var row = new HardIpPattern { MiscInfo = "" };
            row.MiscInfoDict.Add(HardIpConstData.ExecCond, HardIpConstData.ExecCond);
            _sheet.Rows.Add(row);

            // Act
            MiscInfoPreProcessor.PreProcessMiscInfo(_planDic);

            // Assert
            Assert.AreEqual("", row.MiscInfo);
        }
    }
}
