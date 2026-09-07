using System.IO;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.BinCut.Business.BinCutInstance;
using Automation.InputManager.Data;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinCut.FlowNew;

namespace Automation.Test.UT.BinCut
{
    public class BinCutTestBase : FunctionTestBase
    {
        public class TestBinCutInstance(BinCutFinalInstanceRow binCutFinalInstanceRow, BinCutSourceItem binCutSourceItem, BinCutInputData binCutInputData) : BinCutInstanceBase(binCutFinalInstanceRow, binCutSourceItem, binCutInputData)
        {
            internal override string GenerateAcCategory(InstanceRow instanceRow) => "AC";

            internal override string GenerateDcCategory() => "DC";

            protected override string GenerateLevel() => "Level";
        }

        protected BinCutSourceItem SourceRow = null!;
        protected BinCutFinalInstanceRow FinalRow = null!;
        protected BinCutInputData BinCutInputData = null!;
        protected BinCutInstanceBase Instance = null!;

        [TestInitialize]
        public virtual void Setup()
        {
            var binCutFlowSheetRow = new BinCutFlowSheetRow("sheetName", []);
            var newBinCutFlowSheetRow = new NewBinCutFlowSheetRow("sheetName", "job")
            {
                TableType = EnumBinCutTableType.Hv
            };
            SourceRow = new BinCutSourceItem(binCutFlowSheetRow, newBinCutFlowSheetRow, EnumColumnName.FUNC, "TEMP SENSOR MONITOR")
            {
                PerformanceMode = "Pmode"
            };
            FinalRow = new BinCutFinalInstanceRow
            {
                BinCutInstanceRow = new BinCutInstanceRow { Type = BincutInstanceType.Rtos, EnableAndDevice = "EnableAndDevice@site" },
                PatternList = ["Pattern1", "DD_DSSC_P2"],
                FinalJobs = ["Job1"],
                InitList = ["Init1"],
                PayloadList = ["Payload1"]
            };
            BinCutInputData = new BinCutInputData();
            Instance = new TestBinCutInstance(FinalRow, SourceRow, BinCutInputData);
        }

        protected void SetVbtArgTest(string subName, InstanceRow instanceRow)
        {
            string outputPath = Path.Combine(OutputPath, "BinCut", subName);
            string expectPath = Path.Combine(ExpectPath, "BinCut", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            Instance.SetVbtArg(ref instanceRow, FinalRow);

            // Assert
            var instanceSheet = new InstanceSheet("InstSheet");
            instanceSheet.Rows.Add(instanceRow);
            instanceSheet.Write(Path.Combine(outputPath, instanceSheet.Name + ".txt"));

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
