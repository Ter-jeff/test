using System.Collections.Generic;
using System.IO;
using System.Text;

using FileDiffLib;

using IgxlLib.Const;
using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

using Newtonsoft.Json;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class SubFlowSheetTests
    {
        public string InputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        public string OutputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        public string ExpectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void SubFlowSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "SubFlow1";

            // Act
            var subFlowSheet = new SubFlowSheet(sheetName);

            // Assert
            Assert.IsNotNull(subFlowSheet);
            Assert.AreEqual(sheetName, subFlowSheet.Name);
            Assert.AreEqual("DTFlowtableSheet", subFlowSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.FlowTable, subFlowSheet.IgxlSheetName);
            Assert.AreEqual(0, subFlowSheet.Rows.Count);
        }

        [TestMethod]
        public void SubFlowSheet_Constructor_WithSourceSheet()
        {
            // Arrange
            string sheetName = "SubFlow1";
            string sourceSheet = "SourceSheet";

            // Act
            var subFlowSheet = new SubFlowSheet(sheetName, sourceSheet);

            // Assert
            Assert.IsNotNull(subFlowSheet);
            Assert.AreEqual(sheetName, subFlowSheet.Name);
            Assert.AreEqual(sourceSheet, subFlowSheet.SourceInfo.Name);
        }

        [TestMethod]
        public void SubFlowSheet_AddRow()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            var flowRow = new FlowRow
            {
                Label = "Test1",
                Parameter = "TestParam",
                Opcode = "FuncCall"
            };

            // Act
            subFlowSheet.AddRow(flowRow);

            // Assert
            Assert.AreEqual(1, subFlowSheet.Rows.Count);
            Assert.AreEqual("Test1", subFlowSheet.Rows[0].Label);
            Assert.AreEqual("TestParam", subFlowSheet.Rows[0].Parameter);
        }

        [TestMethod]
        public void SubFlowSheet_AddRows()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            var rows = new List<FlowRow>
            {
                new() { Label = "Test1", Parameter = "Param1", Opcode = "FuncCall" },
                new() { Label = "Test2", Parameter = "Param2", Opcode = "FuncCall" },
                new() { Label = "Test3", Parameter = "Param3", Opcode = "FuncCall" }
            };

            // Act
            subFlowSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, subFlowSheet.Rows.Count);
        }

        [TestMethod]
        public void SubFlowSheet_RemoveRow()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            var row1 = new FlowRow { Label = "Test1", Parameter = "Param1" };
            var row2 = new FlowRow { Label = "Test2", Parameter = "Param2" };
            subFlowSheet.AddRow(row1);
            subFlowSheet.AddRow(row2);

            // Act
            subFlowSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, subFlowSheet.Rows.Count);
            Assert.AreEqual("Test2", subFlowSheet.Rows[0].Label);
        }

        [TestMethod]
        public void SubFlowSheet_InsertRow()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            var row1 = new FlowRow { Label = "Test1", Parameter = "Param1" };
            var row3 = new FlowRow { Label = "Test3", Parameter = "Param3" };
            var rowToInsert = new FlowRow { Label = "Test2", Parameter = "Param2" };
            subFlowSheet.AddRow(row1);
            subFlowSheet.AddRow(row3);

            // Act
            int index = subFlowSheet.InsertRow(1, rowToInsert);

            // Assert
            Assert.AreEqual(2, index);
            Assert.AreEqual(3, subFlowSheet.Rows.Count);
            Assert.AreEqual("Test2", subFlowSheet.Rows[1].Label);
            Assert.AreEqual("Test3", subFlowSheet.Rows[2].Label);
        }

        [TestMethod]
        public void SubFlowSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Assert
            Assert.AreEqual("DTFlowtableSheet", subFlowSheet.SheetType);
        }

        [TestMethod]
        public void SubFlowSheet_IgxlSheetName_IsCorrect()
        {
            // Arrange & Act
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Assert
            Assert.AreEqual(IgxlSheetNames.FlowTable, subFlowSheet.IgxlSheetName);
        }

        [TestMethod]
        public void SubFlowSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Assert
            Assert.IsNotNull(subFlowSheet.GetErrors());
            Assert.AreEqual(0, subFlowSheet.GetErrors().Count);
        }

        [TestMethod]
        public void SubFlowSheet_Name_CanBeSet()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1")
            {
                // Act
                Name = "NewSubFlowName"
            };

            // Assert
            Assert.AreEqual("NewSubFlowName", subFlowSheet.Name);
        }

        [TestMethod]
        public void SubFlowSheet_JobNames_CanBeSet()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            var jobNames = new List<string> { "Job1", "Job2", "Job3" };

            // Act
            subFlowSheet.JobNames = jobNames;

            // Assert
            Assert.AreEqual(3, subFlowSheet.JobNames.Count);
        }

        [TestMethod]
        public void SubFlowSheet_GroupNameInMainFlow_CanBeSet()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1")
            {
                // Act
                GroupNameInMainFlow = EnumGroupInMainFlow.None
            };

            // Assert
            Assert.AreEqual(EnumGroupInMainFlow.None, subFlowSheet.GroupNameInMainFlow);
        }

        [TestMethod]
        public void SubFlowSheet_HasConcurrent_CanBeSet()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1")
            {
                // Act
                HasConcurrent = true
            };

            // Assert
            Assert.IsTrue(subFlowSheet.HasConcurrent);
        }

        [TestMethod]
        public void SubFlowSheet_SplitFromSheet_CanBeSet()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            string splitFrom = "SourceSheet";

            // Act
            subFlowSheet.SplitFromSheet = splitFrom;

            // Assert
            Assert.AreEqual(splitFrom, subFlowSheet.SplitFromSheet);
        }

        [TestMethod]
        public void SubFlowSheet_FlowRow_WithMultipleProperties()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            var flowRow = new FlowRow
            {
                Label = "Test1",
                Enable = "ALWAYS",
                Job = "Job1",
                Part = "Part1",
                Env = "ENV1",
                Opcode = "FuncCall",
                Parameter = "TestParam",
                TName = "TestName",
                TNum = "1",
                LoLim = "0.5",
                HiLim = "1.5",
                Result = "PASS"
            };

            // Act
            subFlowSheet.AddRow(flowRow);

            // Assert
            Assert.AreEqual(1, subFlowSheet.Rows.Count);
            Assert.AreEqual("Test1", subFlowSheet.Rows[0].Label);
            Assert.AreEqual("ALWAYS", subFlowSheet.Rows[0].Enable);
            Assert.AreEqual("Job1", subFlowSheet.Rows[0].Job);
            Assert.AreEqual("Part1", subFlowSheet.Rows[0].Part);
        }

        [TestMethod]
        public void SubFlowSheet_AddReturnRow()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Act
            subFlowSheet.AddReturnRow();

            // Assert
            Assert.AreEqual(1, subFlowSheet.Rows.Count);
            Assert.AreEqual(OpCode.Return, subFlowSheet.Rows[0].Opcode);
        }

        [TestMethod]
        public void SubFlowSheet_AddReturnRow_Multiple()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Act
            subFlowSheet.AddReturnRow();
            subFlowSheet.AddReturnRow();

            // Assert
            Assert.AreEqual(2, subFlowSheet.Rows.Count);
            Assert.AreEqual(OpCode.Return, subFlowSheet.Rows[0].Opcode);
            Assert.AreEqual(OpCode.Return, subFlowSheet.Rows[1].Opcode);
        }

        [TestMethod]
        public void SubFlowSheet_AddStartRows_DefaultSheetName()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Act
            subFlowSheet.AddStartRows();

            // Assert
            Assert.IsTrue(subFlowSheet.Rows.Count > 0);
        }

        [TestMethod]
        public void SubFlowSheet_AddStartRows_CustomSheetName()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Act
            subFlowSheet.AddStartRows("CustomSubFlow");

            // Assert
            Assert.IsTrue(subFlowSheet.Rows.Count > 0);
        }

        [TestMethod]
        public void SubFlowSheet_AddStartRows_WithFlowPrefix()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Act
            subFlowSheet.AddStartRows("Flow_TestFlow");

            // Assert
            Assert.IsTrue(subFlowSheet.Rows.Count > 0);
        }

        [TestMethod]
        public void SubFlowSheet_JobNames_Property()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1")
            {
                // Act
                JobNames = ["Job1", "Job2", "Job3"]
            };

            // Assert
            Assert.AreEqual(3, subFlowSheet.JobNames.Count);
            Assert.AreEqual("Job1", subFlowSheet.JobNames[0]);
        }

        [TestMethod]
        public void SubFlowSheet_GroupNameInMainFlow_Property()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1")
            {
                // Act
                GroupNameInMainFlow = EnumGroupInMainFlow.None
            };

            // Assert
            Assert.AreEqual(EnumGroupInMainFlow.None, subFlowSheet.GroupNameInMainFlow);
        }

        [TestMethod]
        public void SubFlowSheet_HasConcurrent_Property()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1")
            {
                // Act
                HasConcurrent = true
            };

            // Assert
            Assert.IsTrue(subFlowSheet.HasConcurrent);
        }

        [TestMethod]
        public void SubFlowSheet_SplitFromSheet_Property()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1")
            {
                // Act
                SplitFromSheet = "OriginalFlow"
            };

            // Assert
            Assert.AreEqual("OriginalFlow", subFlowSheet.SplitFromSheet);
        }

        [TestMethod]
        public void SubFlowSheet_Multiple_AddReturnRows()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Act
            for (int i = 0; i < 5; i++)
            {
                subFlowSheet.AddReturnRow();
            }

            // Assert
            Assert.AreEqual(5, subFlowSheet.Rows.Count);
        }

        [TestMethod]
        public void SubFlowSheet_WriteContent_WithVersion()
        {
            string subName = "WriteContent";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            subFlowSheet.AddReturnRow();

            // Act
            StringBuilder content = subFlowSheet.WriteContent("3.0");

            // Assert
            string json = JsonConvert.SerializeObject(content, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void SubFlowSheet_InitialState()
        {
            // Arrange & Act
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Assert
            Assert.AreEqual(0, subFlowSheet.Rows.Count);
            Assert.AreEqual(EnumGroupInMainFlow.None, subFlowSheet.GroupNameInMainFlow);
            Assert.IsFalse(subFlowSheet.HasConcurrent);
            Assert.AreEqual("", subFlowSheet.SplitFromSheet);
        }

        [TestMethod]
        public void SubFlowSheet_AddClearFlag_AppendsRowCorrectly()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            string testFlag = "FLAG_RETRY_COUNT";

            // Act
            subFlowSheet.AddClearFlag(testFlag);

            // Assert
            Assert.AreEqual(1, subFlowSheet.Rows.Count);
            Assert.AreEqual(OpCode.FlagClear, subFlowSheet.Rows[0].Opcode);
            Assert.AreEqual(testFlag, subFlowSheet.Rows[0].Parameter);
        }

        [TestMethod]
        public void SubFlowSheet_AddFlagToTrue_WithEnableWord_SetsAllProperties()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            string testFlag = "FLAG_PASS_BIN";
            string enableWd = "EN_WORD_01";

            // Act
            subFlowSheet.AddFlagToTrue(testFlag, enableWd);

            // Assert
            Assert.AreEqual(1, subFlowSheet.Rows.Count);
            Assert.AreEqual(OpCode.FlagTrue, subFlowSheet.Rows[0].Opcode);
            Assert.AreEqual(testFlag, subFlowSheet.Rows[0].Parameter);
            Assert.AreEqual(enableWd, subFlowSheet.Rows[0].Enable);
        }

        [TestMethod]
        public void SubFlowSheet_AddFlagToTrue_WithEmptyEnableWord_LeavesEnableEmpty()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Act
            subFlowSheet.AddFlagToTrue("FLAG_PASS_BIN", "");

            // Assert
            Assert.AreEqual(1, subFlowSheet.Rows.Count);
            Assert.AreEqual("", subFlowSheet.Rows[0].Enable);
        }

        [TestMethod]
        public void SubFlowSheet_GetTestAndLimitRows_CollectsUntilNonLimitOpcode()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Test, Parameter = "MainTest" });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.UseLimit, Parameter = "Limit1" });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Characterize, Parameter = "Char1" });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.FlagClear, Parameter = "ShouldStopHere" });

            // Act
            List<FlowRow> collectedRows = subFlowSheet.GetTestAndLimitRows(0);

            // Assert
            Assert.AreEqual(3, collectedRows.Count);
            Assert.AreEqual("MainTest", collectedRows[0].Parameter);
            Assert.AreEqual("Limit1", collectedRows[1].Parameter);
            Assert.AreEqual("Char1", collectedRows[2].Parameter);
        }

        [TestMethod]
        public void SubFlowSheet_RemoveLimitRows_RemovesOnlyConsecutiveLimitOpcodes()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Test });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.UseLimit });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Characterize });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Return });

            // Act
            subFlowSheet.RemoveLimitRows(0);

            // Assert
            Assert.AreEqual(2, subFlowSheet.Rows.Count);
            Assert.AreEqual(OpCode.Test, subFlowSheet.Rows[0].Opcode);
            Assert.AreEqual(OpCode.Return, subFlowSheet.Rows[1].Opcode);
        }

        [TestMethod]
        public void SubFlowSheet_IsSame_ReturnsFalse_WhenRowCountDiffers()
        {
            // Arrange
            var sheetA = new SubFlowSheet("SubFlowA");
            var sheetB = new SubFlowSheet("SubFlowB");
            sheetA.Rows.Add(new FlowRow { Opcode = OpCode.Test });

            // Act & Assert
            Assert.IsFalse(sheetA.IsSame(sheetB));
        }

        [TestMethod]
        public void SubFlowSheet_IsSame_ReturnsFalse_WhenRowCountDiffers_1()
        {
            // Arrange
            var sheetA = new SubFlowSheet("SubFlowA");
            var sheetB = new SubFlowSheet("SubFlowB");
            sheetA.Rows.Add(new FlowRow { Opcode = OpCode.Test });
            sheetB.Rows.Add(new FlowRow { Opcode = OpCode.Test });

            // Act & Assert
            Assert.IsTrue(sheetA.IsSame(sheetB));
        }

        [TestMethod]
        public void SubFlowSheet_InsertBeforeReturnRow_InsertsBeforeValidActiveReturn()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Test });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Return, IsBackup = false });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.FlagClear });

            var rowsToInsert = new List<FlowRow> { new() { Opcode = OpCode.BinTable, Parameter = "BinA" } };

            // Act
            subFlowSheet.InsertBeforeReturnRow(rowsToInsert);

            // Assert
            Assert.AreEqual(4, subFlowSheet.Rows.Count);
            Assert.AreEqual(OpCode.BinTable, subFlowSheet.Rows[1].Opcode);
            Assert.AreEqual(OpCode.Return, subFlowSheet.Rows[2].Opcode);
        }

        [TestMethod]
        public void SubFlowSheet_InsertBeforeReturnRow_AppendsToEnd_IfNoActiveReturnFound()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Test });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Return, IsBackup = true });

            var rowsToInsert = new List<FlowRow> { new() { Opcode = OpCode.BinTable, Parameter = "BinA" } };

            // Act
            subFlowSheet.InsertBeforeReturnRow(rowsToInsert);

            // Assert
            Assert.AreEqual(3, subFlowSheet.Rows.Count);
            Assert.AreEqual(OpCode.BinTable, subFlowSheet.Rows[2].Opcode);
        }

        [TestMethod]
        public void SubFlowSheet_AddTestAndBinTableRow_BypassBinTableTrue_AddsOneRow()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Act
            subFlowSheet.AddTestAndBinTableRow(testItem: "FUNC_TEST", job: "JOB_A", condition: "VMIN", condName: "COND_1", flag: "efuse_fail", bypassNeedBinTable: true);

            // Assert
            Assert.AreEqual(1, subFlowSheet.Rows.Count);
            Assert.AreEqual(OpCode.Test, subFlowSheet.Rows[0].Opcode);
            Assert.AreEqual("JOB_A", subFlowSheet.Rows[0].Job);
            Assert.AreEqual("VMIN", subFlowSheet.Rows[0].DeviceCondition);
            Assert.AreEqual("COND_1", subFlowSheet.Rows[0].DeviceName);
        }

        [TestMethod]
        public void SubFlowSheet_AddTestAndBinTableRow_BypassBinTableFalse_AddsOneRow()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Act
            subFlowSheet.AddTestAndBinTableRow(testItem: "FUNC_TEST", job: "JOB_A", condition: "VMIN", condName: "COND_1", flag: "", bypassNeedBinTable: false);

            // Assert
            Assert.AreEqual(2, subFlowSheet.Rows.Count);
            Assert.AreEqual(OpCode.Test, subFlowSheet.Rows[0].Opcode);
            Assert.AreEqual("JOB_A", subFlowSheet.Rows[0].Job);
            Assert.AreEqual("VMIN", subFlowSheet.Rows[0].DeviceCondition);
            Assert.AreEqual("COND_1", subFlowSheet.Rows[0].DeviceName);
        }

        [TestMethod]
        public void SubFlowSheet_AddTestAndBinTableRow_NopByEnableWordTrue_CreatesNopOpcode()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Act
            subFlowSheet.AddTestAndBinTableRow(testItem: "FUNC_TEST", job: "JOB_A", condition: "VMIN", condName: "COND_1", flag: "efuse_fail", bypassNeedBinTable: true, nopByEnableWord: true);

            // Assert
            Assert.AreEqual(1, subFlowSheet.Rows.Count);
            Assert.AreEqual(OpCode.Nop, subFlowSheet.Rows[0].Opcode);
        }

        [TestMethod]
        public void SubFlowSheet_AddNopTestAndBinTableRow_CreatesTwoRows()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Act
            subFlowSheet.AddNopTestAndBinTableRow(testItem: "NOP_TEST", job: "JOB_B", condition: "VMAX", condName: "COND_2", flag: "F_FAIL", needBinTable: true);

            // Assert
            Assert.AreEqual(2, subFlowSheet.Rows.Count);
            Assert.AreEqual(OpCode.Nop, subFlowSheet.Rows[0].Opcode);
            Assert.AreEqual(OpCode.BinTable, subFlowSheet.Rows[1].Opcode);
            Assert.AreEqual("JOB_B", subFlowSheet.Rows[1].Job);
        }

        [TestMethod]
        public void SubFlowSheet_GetUsedFlowSheets_RecursivelyFindsCalledSheets()
        {
            // Arrange
            var primarySheet = new SubFlowSheet("Primary");
            primarySheet.Rows.Add(new FlowRow { Opcode = "Call", Parameter = "ChildSheet" });
            primarySheet.Rows.Add(new FlowRow { Opcode = "Call", Parameter = "MissingSheet" });

            var childSheet = new SubFlowSheet("ChildSheet");
            childSheet.Rows.Add(new FlowRow { Opcode = "Call", Parameter = "GrandchildSheet" });

            var grandchildSheet = new SubFlowSheet("GrandchildSheet");

            var pool = new List<SubFlowSheet> { primarySheet, childSheet, grandchildSheet };

            // Act
            List<string> usedSheets = primarySheet.GetUsedFlowSheets(pool);

            // Assert
            Assert.AreEqual(2, usedSheets.Count);
            Assert.IsTrue(usedSheets.Contains("ChildSheet"));
            Assert.IsTrue(usedSheets.Contains("GrandchildSheet"));
        }

        [TestMethod]
        public void SubFlowSheet_IsMatchFlowRow_ReturnsTrueOnMatchingActiveRow()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            subFlowSheet.Rows.Add(new FlowRow { Opcode = "TEST", Parameter = "PARAM_A", Enable = "YES", IsBackup = false });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = "TEST", Parameter = "PARAM_B", Enable = "YES", IsBackup = true });

            var targetMatch = new FlowRow { Opcode = "test", Parameter = "param_a", Enable = "yes" };
            var targetMismatch = new FlowRow { Opcode = "test", Parameter = "param_b", Enable = "yes" };

            // Act & Assert
            Assert.IsTrue(subFlowSheet.IsMatchFlowRow(targetMatch));
            Assert.IsFalse(subFlowSheet.IsMatchFlowRow(targetMismatch));
        }

        [TestMethod]
        public void SubFlowSheet_ReplaceFlowRow_InsertsRowAtMatchingIndexAndReturnsOldRow()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            var originalRow = new FlowRow { Opcode = "SET", Parameter = "VOUT", Enable = "ALWAYS", IsBackup = false };
            subFlowSheet.Rows.Add(originalRow);

            var trackingRow = new FlowRow { Opcode = "SET", Parameter = "VOUT", Enable = "ALWAYS" };

            // Act
            FlowRow result = subFlowSheet.ReplaceFlowRow(trackingRow);

            // Assert
            Assert.IsNotNull(result);
            Assert.AreSame(originalRow, result);
            Assert.AreEqual(2, subFlowSheet.Rows.Count);
        }

        [TestMethod]
        public void SubFlowSheet_RemoveTheSameIf_RemovesConsecutiveRedundantIfBlocks()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            // Mimicking consecutive identical simple IF structures: IF(A)..ENDIF IF(A)..ENDIF
            subFlowSheet.Rows.Add(new FlowRow { Opcode = "If", Parameter = "FLAG_A" });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = "EndIf", Parameter = "FLAG_A" });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = "If", Parameter = "FLAG_A" });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = "EndIf", Parameter = "FLAG_A" });

            // Act
            subFlowSheet.RemoveTheSameIf();

            // Assert
            // It clears the adjacent EndIf (index 1) and consecutive If (index 2) to join the blocks
            Assert.AreEqual(2, subFlowSheet.Rows.Count);
            Assert.AreEqual("If", subFlowSheet.Rows[0].Opcode);
            Assert.AreEqual("EndIf", subFlowSheet.Rows[1].Opcode);
        }

        [TestMethod]
        public void SubFlowSheet_RemoveTheSameIf_RemovesConsecutiveRedundantIfBlocks_1()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            // Mimicking consecutive identical simple IF structures: IF(A)..ENDIF IF(A)..ENDIF
            subFlowSheet.Rows.Add(new FlowRow { Opcode = "If", Parameter = "FLAG_A" });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = "If", Parameter = "FLAG_A" });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = "EndIf", Parameter = "FLAG_A" });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = "EndIf", Parameter = "FLAG_A" });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = "If", Parameter = "FLAG_A" });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = "EndIf", Parameter = "FLAG_A" });

            // Act
            subFlowSheet.RemoveTheSameIf();

            // Assert
            // It clears the adjacent EndIf (index 1) and consecutive If (index 2) to join the blocks
            Assert.AreEqual(6, subFlowSheet.Rows.Count);
            Assert.AreEqual("If", subFlowSheet.Rows[0].Opcode);
            Assert.AreEqual("EndIf", subFlowSheet.Rows[3].Opcode);
        }

        [TestMethod]
        public void SubFlowSheet_FilterFlowJobs_ClearsJobString_WhenAllJobsMatchExactly()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            var targetRow = new FlowRow { Opcode = "Test", Job = "JOB_A,JOB_B" };
            subFlowSheet.Rows.Add(targetRow);

            var hardwareIpJobs = new List<string> { "JOB_B", "JOB_A" };

            // Act
            subFlowSheet.FilterFlowJobs(hardwareIpJobs);

            // Assert
            Assert.AreEqual("", targetRow.Job);
        }

        [TestMethod]
        public void SubFlowSheet_FilterFlowJobs_RetainsJobString_WhenJobsAreSubsetOrMismatched()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            var targetRow = new FlowRow { Opcode = "Test", Job = "JOB_A,JOB_C" };
            subFlowSheet.Rows.Add(targetRow);

            var hardwareIpJobs = new List<string> { "JOB_A", "JOB_B" };

            // Act
            subFlowSheet.FilterFlowJobs(hardwareIpJobs);

            // Assert
            Assert.AreEqual("JOB_A,JOB_C", targetRow.Job);
        }

        [TestMethod]
        public void SubFlowSheet_InsertBinTableBeforeReturnRow_CreatesDistinctBinTableRows()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Return, IsBackup = false });

            // Create list with a duplicate name to verify the Distinct() filter
            var binTableRows = new List<BinTableRow>
            {
                new() { Name = "BIN_A" },
                new() { Name = "BIN_B" },
                new() { Name = "BIN_A" }
            };

            // Act
            subFlowSheet.InsertBinTableBeforeReturnRow(binTableRows);

            // Assert
            Assert.AreEqual(3, subFlowSheet.Rows.Count);
            Assert.AreEqual(OpCode.BinTable, subFlowSheet.Rows[0].Opcode);
            Assert.AreEqual("BIN_A", subFlowSheet.Rows[0].Parameter);
            Assert.AreEqual(OpCode.BinTable, subFlowSheet.Rows[1].Opcode);
            Assert.AreEqual("BIN_B", subFlowSheet.Rows[1].Parameter);
            Assert.AreEqual(OpCode.Return, subFlowSheet.Rows[2].Opcode);
        }

        [TestMethod]
        public void SubFlowSheet_GetFailFlags_ExtractsAndFlattensDistinctFlags()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");

            // Add rows including comma-separated strings and duplicate case variations
            subFlowSheet.Rows.Add(new FlowRow { FailAction = "FLAG_ERR1,FLAG_ERR2" });
            subFlowSheet.Rows.Add(new FlowRow { FailAction = "" });
            subFlowSheet.Rows.Add(new FlowRow { FailAction = "flag_err1" });

            // Act
            List<string> result = subFlowSheet.GetFailFlags();

            // Assert
            // Expecting 2 unique entries (ignoring case variants like 'flag_err1')
            Assert.AreEqual(2, result.Count);
            Assert.IsTrue(result.Contains("FLAG_ERR1") || result.Contains("flag_err1"));
            Assert.IsTrue(result.Contains("FLAG_ERR2"));
        }

        [TestMethod]
        public void SubFlowSheet_AddEndRows_UsesDefaultSheetName_AndStripsFlowPrefix()
        {
            // Arrange
            // Sheet name contains "Flow_" to test the ReplaceStartsWith mechanism
            var subFlowSheet = new SubFlowSheet("Flow_ProductionTest");

            // Act
            subFlowSheet.AddEndRows();

            // Assert
            // Expecting 3 added rows based on the internal orchestration: Print, Footer, and Return
            Assert.AreEqual(3, subFlowSheet.Rows.Count);

            // Check that the stripped name "ProductionTest" was fed into the downstream rows
            Assert.IsTrue(subFlowSheet.Rows[1].Parameter.Contains("ProductionTest"));
            Assert.AreEqual(OpCode.Return, subFlowSheet.Rows[2].Opcode);
        }

        [TestMethod]
        public void SubFlowSheet_AddEndRows_UsesCustomSheetNameOverride()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("DefaultName");

            // Act
            subFlowSheet.AddEndRows("Flow_CustomOverride");

            // Assert
            Assert.AreEqual(3, subFlowSheet.Rows.Count);
            Assert.IsTrue(subFlowSheet.Rows[1].Parameter.Contains("CustomOverride"));
            Assert.IsFalse(subFlowSheet.Rows[1].Parameter.Contains("DefaultName"));
        }

        [TestMethod]
        public void SubFlowSheet_AddFooterRow_AppendsToBottom_WhenRowIdxIsEmpty()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Test, Parameter = "ExistingRow" });

            // Act
            subFlowSheet.AddFooterRow("BlockA", "FLAG_EN", "");

            // Assert
            Assert.AreEqual(2, subFlowSheet.Rows.Count);
            Assert.AreEqual("BlockA_Footer_1", subFlowSheet.Rows[1].Parameter);
            Assert.AreEqual("FLAG_EN", subFlowSheet.Rows[1].Enable);
            Assert.AreEqual(OpCode.Test, subFlowSheet.Rows[1].Opcode);
        }

        [TestMethod]
        public void SubFlowSheet_AddFooterRow_InsertsAtSpecificIndex_WhenRowIdxIsSupplied()
        {
            // Arrange
            var subFlowSheet = new SubFlowSheet("SubFlow1");
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Test, Parameter = "Row0" });
            subFlowSheet.Rows.Add(new FlowRow { Opcode = OpCode.Test, Parameter = "Row1" });

            // Act
            // Targets index 1 to sit right between the two existing items
            subFlowSheet.AddFooterRow("BlockB", "FLAG_EN", "1");

            // Assert
            Assert.AreEqual(3, subFlowSheet.Rows.Count);
            Assert.AreEqual("BlockB_Footer_1", subFlowSheet.Rows[1].Parameter);
            Assert.AreEqual("Row1", subFlowSheet.Rows[2].Parameter);
        }

    }
}
