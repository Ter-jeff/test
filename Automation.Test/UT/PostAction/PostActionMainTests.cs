using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.PostAction;
using Automation.GenerateIgxl.PostAction.PatSetProcessing;
using Automation.Static;

using CommonLib.Extension;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class PostActionMainTests : FunctionTestBase
    {
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void CopyFromExtraSheetsTest()
        {
            string subName = "CopyFromExtraSheets";

            string inputPath = Path.Combine(InputPath, "PostAction", subName);
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            LocalSpecs.TarFolder = outputPath;
            string file = Path.Combine(inputPath, "TestInst_Common.txt");
            Stream stream = File.OpenRead(file);
            InstanceSheet sheet = new ReadInstanceSheet().ReadSheet(stream, "TestInst_Common");
            TestProgram.IgxlWorkBk.AddInsSheet(outputPath, sheet);

            string file1 = Path.Combine(inputPath, "PatSets_ClockOut.txt");
            Stream stream1 = File.OpenRead(file1);
            PatSetSheet sheet1 = new ReadPatSetSheet().ReadSheet(stream1, "PatSets_ClockOut");
            TestProgram.IgxlWorkBk.AddPatSetSheet(outputPath, sheet1);

            string file2 = Path.Combine(inputPath, "Flow_nWire_Default.txt");
            string delete = Path.Combine(inputPath, "Flow_nWire_Default_autogen.txt");
            File.Copy(file2, Path.Combine(outputPath, "Flow_nWire_Default.txt"));
            if (File.Exists(delete))
            {
                File.Delete(delete);
            }
            Stream stream2 = File.OpenRead(file2);
            SubFlowSheet sheet2 = new ReadFlowSheet().ReadSheet(stream2, "Flow_nWire_Default");
            TestProgram.IgxlWorkBk.AddSubFlowSheet(outputPath, sheet2);

            var main = new PostActionMain();
            main.CopyFromExtraSheets(inputPath);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        [Description("ConvertJobToEnableWord should convert Job list to Enable string with || if no !")]
        public void ShouldConvertJobToEnableOr()
        {
            // Arrange
            var subSheet = new SubFlowSheet("MySubFlow")
            {
                Rows = new FlowRows
                {
                    new FlowRow { Opcode = "test", Job = "JOB1,JOB2" }
                }
            };

            var subflows = new Dictionary<string, SubFlowSheet> { { "s1", subSheet } };
            var mainFlows = new Dictionary<string, MainFlow>()
            {
                { "Flow_nWire_Default", new MainFlow("Flow_nWire_Default")
                    {
                        Rows = new FlowRows()
                        {
                            new FlowRow() { Job = "CP1" },
                            new FlowRow() { Job = "CP1" , Enable = "Enable" }
                        }
                    }
                }
            };

            // Act
            var main = new PostActionMain();
            main.ConvertJobToEnableWord(subflows, mainFlows);

            // Assert
            FlowRow row = subSheet.Rows.First();
            Assert.AreEqual("JOB1||JOB2", row.Enable);
            Assert.AreEqual("", row.Job, "Job should be cleared after conversion");
        }

        [TestMethod]
        [Description("ConvertJobToEnableWord should use && when Job contains !")]
        public void ShouldConvertJobToEnableAnd()
        {
            var subSheet = new SubFlowSheet("FlowA")
            {
                Rows = new FlowRows
                {
                    new FlowRow { Opcode = "test", Job = "!JOB1,JOB2" }
                }
            };
            var subflows = new Dictionary<string, SubFlowSheet> { { "x", subSheet } };
            var mainFlows = new Dictionary<string, MainFlow>();

            var main = new PostActionMain();
            main.ConvertJobToEnableWord(subflows, mainFlows);

            FlowRow row = subSheet.Rows.First();
            Assert.AreEqual("!JOB1&&JOB2", row.Enable);
        }

        [TestMethod]
        [Description("ConvertJobToEnableWord should skip Flow_Table_Main_Init_EnableWd")]
        public void ShouldSkipEnableWdFlow()
        {
            var subSheet = new SubFlowSheet("Flow_Table_Main_Init_EnableWd")
            {
                Rows = new FlowRows
                {
                    new FlowRow { Opcode = "test", Job = "JOBX" }
                }
            };
            var subflows = new Dictionary<string, SubFlowSheet> { { "skip", subSheet } };
            var mainFlows = new Dictionary<string, MainFlow>();

            var main = new PostActionMain();
            main.ConvertJobToEnableWord(subflows, mainFlows);

            FlowRow row = subSheet.Rows.First();
            Assert.AreEqual("JOBX", row.Job, "Should not modify skipped flow");
            Assert.AreEqual("", row.Enable, "Enable should remain empty");
        }

        [TestMethod]
        [Description("ConvertJobToEnableWord should append new job enables when Enable already set")]
        public void ShouldAppendToExistingEnable()
        {
            var subSheet = new SubFlowSheet("MyFlow")
            {
                Rows = new FlowRows
                {
                    new FlowRow { Opcode = "test", Job = "JOB1", Enable = "EXISTING" }
                }
            };
            var subflows = new Dictionary<string, SubFlowSheet> { { "f1", subSheet } };
            var mainFlows = new Dictionary<string, MainFlow>();

            var main = new PostActionMain();
            main.ConvertJobToEnableWord(subflows, mainFlows);

            FlowRow row = subSheet.Rows.First();
            Assert.AreEqual("(EXISTING)&&(JOB1)", row.Enable);
        }

        [TestMethod]
        [Description("")]
        public void SplitPatSet()
        {
            string subName = "SplitPatSet";

            string inputPath = Path.Combine(InputPath, "PostAction", subName);
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string[] files = Directory.GetFiles(inputPath);

            string file1 = Path.Combine(inputPath, Path.GetFileName(files[0]));
            Stream stream1 = File.OpenRead(file1);
            PatSetSheet patSetSheet = new ReadPatSetSheet().ReadSheet(stream1, Path.GetFileNameWithoutExtension(files[0]));

            var main = new SplitPatSet(200);
            var dict = new Dictionary<string, PatSetSheet>(StringExtensions.IgnoreCase);

            string? directory = Path.GetDirectoryName(file1);
            string filenameNoExt = Path.GetFileNameWithoutExtension(file1);

            string newPath = Path.Combine(directory!, filenameNoExt);

            dict.Add(newPath, patSetSheet);
            main.SplitPatset(dict);

            foreach (KeyValuePair<string, PatSetSheet> item in dict)
            {
                item.Value.Write(Path.Combine(outputPath, item.Value.Name) + ".txt", "2.3");
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }

        }

        [TestMethod]
        public void PatSetCompare_IdenticalPatSets_ReturnsTrue()
        {
            // Arrange
            var target = new PostActionMain();
            var row1 = new PatSet { PatSetName = "PS1", Domain = "CPU" };
            var row2 = new PatSet { PatSetName = "PS1", Domain = "CPU" };

            // Act
            bool result = target.PatSetCompare(row1, row2);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void PatSetCompare_DifferentPatSetName_ReturnsFalse()
        {
            // Arrange
            var target = new PostActionMain();
            var row1 = new PatSet { PatSetName = "PS1" };
            var row2 = new PatSet { PatSetName = "PS2" };

            // Act
            bool result = target.PatSetCompare(row1, row2);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void PatSetCompare_DifferentRowNumAndDomain_StillReturnsTrue()
        {
            // Arrange - RowNum and Domain are explicitly excluded from the comparison.
            var target = new PostActionMain();
            var row1 = new PatSet { PatSetName = "PS1", Domain = "CPU", RowNum = 1 };
            var row2 = new PatSet { PatSetName = "PS1", Domain = "GFX", RowNum = 2 };

            // Act
            bool result = target.PatSetCompare(row1, row2);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void PatSetCompare_DifferentPatSetRowsCount_ReturnsFalse()
        {
            // Arrange
            var target = new PostActionMain();
            var row1 = new PatSet { PatSetName = "PS1" };
            row1.AddRow(new PatSetRow { File = "pat1" });
            var row2 = new PatSet { PatSetName = "PS1" };

            // Act
            bool result = target.PatSetCompare(row1, row2);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void InstanceCompare_IdenticalRows_ReturnsTrue()
        {
            // Arrange
            var target = new PostActionMain();
            var row1 = new InstanceRow { SheetName = "S", ColumnA = "", TestName = "T1", VbtName = "V1", Args = ["A", "B"] };
            var row2 = new InstanceRow { SheetName = "S", ColumnA = "", TestName = "T1", VbtName = "V1", Args = ["A", "B"] };

            // Act
            bool result = target.InstanceCompare(row1, row2);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void InstanceCompare_DifferentTestName_ReturnsFalse()
        {
            // Arrange
            var target = new PostActionMain();
            var row1 = new InstanceRow { SheetName = "S", ColumnA = "", TestName = "T1" };
            var row2 = new InstanceRow { SheetName = "S", ColumnA = "", TestName = "T2" };

            // Act
            bool result = target.InstanceCompare(row1, row2);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void InstanceCompare_DifferentRowNum_StillReturnsTrue()
        {
            // Arrange - RowNum is explicitly excluded from the comparison.
            var target = new PostActionMain();
            var row1 = new InstanceRow { SheetName = "S", ColumnA = "", TestName = "T1", RowNum = 1 };
            var row2 = new InstanceRow { SheetName = "S", ColumnA = "", TestName = "T1", RowNum = 2 };

            // Act
            bool result = target.InstanceCompare(row1, row2);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void InstanceCompare_DifferentArgs_ReturnsFalse()
        {
            // Arrange
            var target = new PostActionMain();
            var row1 = new InstanceRow { SheetName = "S", ColumnA = "", TestName = "T1", Args = ["A"] };
            var row2 = new InstanceRow { SheetName = "S", ColumnA = "", TestName = "T1", Args = ["B"] };

            // Act
            bool result = target.InstanceCompare(row1, row2);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void InstanceCompare_DifferentInitList_ReturnsFalse()
        {
            // Arrange
            var target = new PostActionMain();
            var row1 = new InstanceRow { SheetName = "S", ColumnA = "", TestName = "T1", InitList = ["Init1"] };
            var row2 = new InstanceRow { SheetName = "S", ColumnA = "", TestName = "T1", InitList = ["Init2"] };

            // Act
            bool result = target.InstanceCompare(row1, row2);

            // Assert
            Assert.IsFalse(result);
        }
    }
}
