using System.IO;

using Automation.GenerateIgxl.PreAction.GenPinMap;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

using OfficeOpenXml;

using TestPlanLib.DataStruct;

namespace Automation.Test.UT.PreAction
{
    [TestClass]
    public class PinMapTests : FunctionTestBase
    {
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void WorkFlow_Format1()
        {
            string subName = "PinMapMain";
            string outputPath = Path.Combine(OutputPath, "PreAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PreAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            Directory.CreateDirectory(outputPath);

            // Arrange
            var pinMapSheet = new PinMapSheet("TestPinMap");
            pinMapSheet.AddPin(new Pin("PIN_A", "I/O"));
            pinMapSheet.AddPin(new Pin("PIN_B", "I/O"));

            var ioConti = new IoContiSheet();
            ioConti.Rows.Add(new IoContiRow { BumpName = "PIN_A" });

            using var package = new ExcelPackage();
            ExcelWorksheet ioPinGroupSheet = package.Workbook.Worksheets.Add("ioPinGroupSheet");
            ioPinGroupSheet.Cells[1, 1].Value = "Pin name contained Pin Group (CP)";
            ioPinGroupSheet.Cells[1, 2].Value = "Pin name contained Pin Group (FT)";
            ioPinGroupSheet.Cells[1, 3].Value = "AAA";

            ioPinGroupSheet.Cells[2, 1].Value = "PIN_A_CP";
            ioPinGroupSheet.Cells[2, 2].Value = "PIN_A_FT";
            ioPinGroupSheet.Cells[2, 3].Value = "X";

            // Act
            var pinMapMain = new PinMapMain();
            PinMapSheet result = pinMapMain.WorkFlow(pinMapSheet, ioPinGroupSheet, ioConti);

            // Assert
            string json = JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail! Check result.json vs expected.");
            }
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void WorkFlow_Format0()
        {
            string subName = "PinMapMain_Format0";
            string outputPath = Path.Combine(OutputPath, "PreAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PreAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            Directory.CreateDirectory(outputPath);

            // Arrange
            var pinMapSheet = new PinMapSheet("TestPinMap");

            var pinA = new Pin("PIN_A", "I/O");
            var pinB = new Pin("PIN_B", "I/O");
            pinMapSheet.AddPin(pinA);
            pinMapSheet.AddPin(pinB);
            var pinGroup = new PinGroup("PIN_C", PinMapConst.TypeIo);
            pinGroup.AddPin(pinA);
            pinGroup.AddPin(pinB);
            pinMapSheet.AddGroup(pinGroup);

            var ioConti = new IoContiSheet();
            ioConti.Rows.Add(new IoContiRow { BumpName = "PIN_A" });

            using var package = new ExcelPackage();
            ExcelWorksheet ioPinGroupSheet = package.Workbook.Worksheets.Add("ioPinGroupSheet");
            ioPinGroupSheet.Cells[1, 1].Value = "Pin Group Name";
            ioPinGroupSheet.Cells[1, 2].Value = "Pin Name";

            ioPinGroupSheet.Cells[2, 1].Value = "Group";
            ioPinGroupSheet.Cells[2, 2].Value = "PIN_A+PIN_C";

            ioPinGroupSheet.Cells[3, 1].Value = "Group";
            ioPinGroupSheet.Cells[3, 2].Value = "PIN_C-PIN_A";

            ioPinGroupSheet.Cells[4, 1].Value = "Group";
            ioPinGroupSheet.Cells[4, 2].Value = "PIN_*";

            ioPinGroupSheet.Cells[5, 1].Value = "Group";
            ioPinGroupSheet.Cells[5, 2].Value = "PIN_A";

            ioPinGroupSheet.Cells[6, 1].Value = "Group";
            ioPinGroupSheet.Cells[6, 2].Value = "PIN_C";

            ioPinGroupSheet.Cells[7, 1].Value = "Group";
            ioPinGroupSheet.Cells[7, 2].Value = "PIN_D";

            // Act
            var pinMapMain = new PinMapMain();
            PinMapSheet result = pinMapMain.WorkFlow(pinMapSheet, ioPinGroupSheet, ioConti);

            // Assert
            string json = JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail! Check result.json vs expected.");
            }
        }
    }
}
