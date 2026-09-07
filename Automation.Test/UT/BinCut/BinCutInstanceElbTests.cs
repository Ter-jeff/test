using System.IO;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.BinCut.Business.BinCutInstance;
using Automation.Static;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.Test.UT.BinCut
{
    [TestClass]
    public class BinCutInstanceElbTests : BinCutTestBase
    {
        [TestInitialize]
        public override void Setup()
        {
            base.Setup();
            FinalRow = new BinCutFinalInstanceRow
            {
                BinCutInstanceRow = new BinCutInstanceRow { Type = BincutInstanceType.Rtos },
                PatternList = ["Pattern1"],
                FinalJobs = ["Job1"],
                InitList = ["Init1"],
                PayloadList = ["Payload1"]
            };
            Instance = new BinCutInstanceElb(FinalRow, SourceRow, BinCutInputData);
        }

        [TestMethod]
        public void GenerateAcCategory()
        {
            var instanceRow = new InstanceRow
            {
                TestName = "TestInstance",
                DcSelector = "NV"
            };

            Assert.AreEqual("TBD", Instance.GenerateAcCategory(instanceRow));
        }

        [TestMethod]
        public void GenerateDcCategory()
        {
            Assert.AreEqual("Bincut_X_X_X", Instance.GenerateDcCategory());
        }

        [TestMethod]
        public void GetDcSelector_ShouldReturnCorrectSelector()
        {
            Assert.AreEqual("Min", Instance.GetDcSelector("LV"));
            Assert.AreEqual("Max", Instance.GetDcSelector("HV"));
            Assert.AreEqual("Typ", Instance.GetDcSelector("NV"));
        }

        [TestMethod]
        public void GetInstanceName_ShouldContainBinningDomain()
        {
            string name = Instance.GetInstanceName();
            StringAssert.Contains(name, SourceRow.GetBinType());
        }

        [TestMethod]
        public void GenerateFlowRow_ShouldReturnFlowRowWithExpectedValues()
        {
            FlowRow row = Instance.GenerateFlowRow(false, false, false);
            Assert.AreNotEqual(null, row);
            Assert.IsFalse(string.IsNullOrEmpty(row.Parameter));
            Assert.IsTrue(row.Opcode == OpCode.Test || row.Opcode == OpCode.Nop);
        }

        [TestMethod]
        public void GetBinTableRow_ShouldReturnFlowRowWithBinTableOpcode()
        {
            FlowRow row = Instance.GetBinTableRow();
            Assert.AreEqual(OpCode.BinTable, row.Opcode);
            StringAssert.StartsWith(row.Parameter, "Bin_");
        }

        [TestMethod]
        public void SetVbtArg()
        {
            string subName = "BinCutInstanceElb_SetVbtArg";
            var instanceRow = new InstanceRow
            {
                TestName = "TestInstance",
                DcSelector = "NV"
            };
            SetVbtArgTest(subName, instanceRow);
        }

        [TestMethod]
        public void GenerateInstance()
        {
            string subName = "BinCutInstanceElbTests_GenerateInstance";
            string outputPath = Path.Combine(OutputPath, "BinCut", subName);
            string expectPath = Path.Combine(ExpectPath, "BinCut", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            LocalSpecs.CsLibraryFolder = "";
            InstanceRow instanceRow = Instance.GenerateInstance();

            // Assert
            string json = JsonConvert.SerializeObject(instanceRow, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
