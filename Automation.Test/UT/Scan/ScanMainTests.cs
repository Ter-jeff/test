using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.Scan;
using Automation.InputManager.Data;
using Automation.Static;
using Automation.Test.Static;

using CommonLib.Extension;

using FileDiffLib;

using IgxlLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Singleton;
using TestPlanLib.VbtLib;

namespace Automation.Test.UT.Scan
{
    [TestClass]
    public class ScanMainTests : FunctionTestBase
    {
        private static List<Function> _functions = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            _functions = TestSuiteInitialize.Functions;
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void ScanMainTest_VBT()
        {
            string subName = "ScanMain_VBT";
            string outputPath = Path.Combine(OutputPath, "Scan", subName);
            string expectPath = Path.Combine(ExpectPath, "Scan", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            TestProgram.Clear();
            TestProgram.VbtFunctionLib.AddVbtFunctionRange([.. _functions.Where(x => x.Type.EqualsIgnoreCase("VBT"))]);
            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);

            using (var scanMain = new ScanMain())
            {
                scanMain.Execute(null);
                scanMain.Print();
            }

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void ScanMainTest_NET()
        {
            string subName = "ScanMain_NET";
            string outputPath = Path.Combine(OutputPath, "Scan", subName);
            string expectPath = Path.Combine(ExpectPath, "Scan", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            TestProgram.Clear();
            TestProgram.VbtFunctionLib.AddVbtFunctionRange([.. _functions.Where(x => x.Type.EqualsIgnoreCase(".NET"))]);
            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);

            using (var scanMain = new ScanMain())
            {
                scanMain.Execute(null);
                scanMain.Print();
            }

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void GenerateTurboModeCheckFlow_Should_Create_Row_For_Each_Instance()
        {
            // Arrange
            ScanInputData data = CreateScanInputData();
            TestProgram.IgxlWorkBk = new IgxlWorkBook();
            LocalSpecs.TarFolder = OutputPath;
            FolderStructure.CreateFolder();
            // Act
            var sut = new ScanMain();
            sut.GenerateTurboModeCheckFlow(data);

            // Assert
            Assert.AreEqual(2, TestProgram.IgxlWorkBk.InsSheets.First(x => x.Value.Name.Equals("TestInst_TurboModeCheck")).Value.Rows.Count);
        }
        private static ScanInputData CreateScanInputData()
        {
            var turboSheet = new BinCutInstanceSheet("TurboModeCheck");

            turboSheet.Rows.Add(
                new BinCutInstanceRow
                {
                    FlowName = "FLOW1",
                    Instance = "INST1",
                    FunctionName = "FUNC1"
                });

            turboSheet.Rows.Add(
                new BinCutInstanceRow
                {
                    FlowName = "FLOW1",
                    Instance = "INST2",
                    FunctionName = "FUNC1"
                });

            return new ScanInputData
            {
                TurboModeInstanceSheet = turboSheet
            };
        }
    }
}
