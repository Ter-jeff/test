using Automation.Static;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class LocalSpecsTests : FunctionTestBase
    {
        [DataTestMethod]
        [DataRow("CurrentProject", "CurrentProject")]
        [DataRow("CurrentProject_CPU", "CurrentProject_CPU")]
        [DataRow("CurrentProject_GFX", "CurrentProject_GFX")]
        public void GetProjectNameMapping(string currentProject, string expected)
        {
            LocalSpecs.CurrentProject = currentProject;
            string projectNameMapping = LocalSpecs.GetProjectNameMapping;

            Assert.AreEqual(expected, projectNameMapping);
        }
    }
}
