using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.PostAction.ConcurrentSequence;
using Automation.Static;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Concurrent;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class ConcurrentSequenceMainTests : FunctionTestBase
    {
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void WorkFlow_ShouldGenerateConcurrentSequence_WhenInputIsValid()
        {
            string subName = "ConcurrentSequenceMain";
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);
            _ = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            LocalSpecs.TarFolder = outputPath;

            var concurrentFlow = new ConcurrentFlowSheet("ConcurrentFlowSheet");
            concurrentFlow.Rows.Add(new ConcurrentFlowSheetRow
            {
                SequenceName = "Seq1",
                Subflows = ["SubFlow_A"]
            });

            var subFlowSheets = new Dictionary<string, SubFlowSheet>();
            var subFlowA = new SubFlowSheet("")
            {
                Rows = new FlowRows
                {
                new FlowRow
                {
                    Opcode = OpCode.Test,
                    Parameter = "TestInst1"
                }
                }
            };
            subFlowSheets.Add("Path/To/SubFlow_A", subFlowA);

            var insSheets = new Dictionary<string, InstanceSheet>();

            var main = new ConcurrentSequenceMain();
            main.WorkFlow(concurrentFlow, subFlowSheets, insSheets);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
