using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.PostAction.GenGroupFlowSheet;
using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.Static;

using IgxlLib;
using IgxlLib.Enums;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class GroupFlowSheetMainTests : FunctionTestBase
    {
        private GroupFlowSheetMain _processor = null!;

        [TestInitialize]
        public void Setup()
        {
            _processor = new GroupFlowSheetMain();
            TestProgram.IgxlWorkBk = new IgxlWorkBook();
            LocalSpecs.TarFolder = OutputPath;
        }

        [TestMethod]
        public void WorkFlow_FlowScanSa_SetsGroupNameAndGeneratesSheet()
        {
            // Arrange
            string subFlowName = "SubFlow_A";
            string jobName = "Job1";

            var subFlowSheets = new Dictionary<string, SubFlowSheet>
            {
                { subFlowName, new SubFlowSheet(subFlowName) { GroupNameInMainFlow = EnumGroupInMainFlow.FlowScanSa } }
            };

            var sequences = new List<FlowSequence>
            {
                new("") { SubFlowName = subFlowName, Enable = "Yes" }
            };

            var mainFlows = new List<MainFlowBase>
            {
                new() { JobName = jobName, Sequences = sequences }
            };

            // Act
            List<MainFlowBase> result = _processor.WorkFlow(subFlowSheets, mainFlows);

            // Assert
            FlowSequence sequence = result.First().Sequences.First();
            Assert.AreEqual(EnumGroupInMainFlow.FlowScanSa, sequence.GroupNameInMainFlow);
            Assert.IsTrue(sequence.GroupSheetName.Contains("FlowScanSa"));
            Assert.IsTrue(sequence.GroupSheetName.Contains(jobName));
        }

        [TestMethod]
        public void WorkFlow_DuplicateGroups_IncrementsCounter()
        {
            // Arrange
            var sequences = new List<FlowSequence>
            {
                new("") { SubFlowName = "S1", GroupNameInMainFlow = EnumGroupInMainFlow.FlowScanSa },
                new("") { SubFlowName = "Gap", GroupNameInMainFlow = EnumGroupInMainFlow.None },
                new("") { SubFlowName = "S2", GroupNameInMainFlow = EnumGroupInMainFlow.FlowScanSa }
            };

            var mainFlows = new List<MainFlowBase> { new() { JobName = "J", Sequences = sequences } };
            // No lookups needed for this path
            var subFlowSheets = new Dictionary<string, SubFlowSheet>();

            // Act
            List<MainFlowBase> result = _processor.WorkFlow(subFlowSheets, mainFlows);

            // Assert
            var groupNames = result.First().Sequences.Where(x => x.GroupNameInMainFlow != EnumGroupInMainFlow.None).Select(x => x.GroupSheetName).ToList();

            Assert.AreNotEqual(groupNames[0], groupNames[1]);
            Assert.IsTrue(groupNames[1].EndsWith("_1"), "Second group of same name should have suffix _1");
        }
    }
}
