using System.IO;

using Automation.GenerateIgxl.PostAction.InstCommon;
using Automation.Static;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class InstCommonMainTests : FunctionTestBase
    {
        private InstCommonMain _instCommonMain = null!;

        [TestInitialize]
        public void Setup()
        {
            _instCommonMain = new InstCommonMain();

            SubFlowSheet subFlowSheet = new SubFlowSheet("Flow_AAA");
            subFlowSheet.AddRow(new FlowRow { Opcode = OpCode.Print });
            subFlowSheet.AddRow(new FlowRow { Opcode = OpCode.Print });
            TestProgram.IgxlWorkBk.AddSubFlowSheet(Path.Combine(OutputPath, "Flow_AAA.txt"), subFlowSheet);
            InstanceSheet instanceSheet = new InstanceSheet("Inst_AAA");
            TestProgram.IgxlWorkBk.AddInsSheet(Path.Combine(OutputPath, "Inst_AAA.txt"), instanceSheet);

            instanceSheet = new InstanceSheet("TestInst_Common", "", true);
            TestProgram.IgxlWorkBk.AddInsSheet(OutputPath, instanceSheet);
        }

        [TestMethod]
        public void WorkFlowWithoutHeader_WhenSheetNameIsNullOrEmpty_ShouldReturnImmediately()
        {
            string subName = "WorkFlowWithoutHeader";
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;

            // Act
            _instCommonMain.WorkFlowWithoutHeader("TestInst_Common");

            TestProgram.IgxlWorkBk.TryGetTestInstCommon(out InstanceSheet? instanceSheet);

            string json = JsonConvert.SerializeObject(instanceSheet, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void WorkFlowWithHeader_WhenWorksheetIsNull_ShouldReturnImmediately()
        {
            string subName = "WorkFlowWithHeader";
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;

            // Act
            _instCommonMain.WorkFlowWithHeader("TestInst_Common");

            TestProgram.IgxlWorkBk.TryGetTestInstCommon(out InstanceSheet? instanceSheet);

            string json = JsonConvert.SerializeObject(instanceSheet, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
