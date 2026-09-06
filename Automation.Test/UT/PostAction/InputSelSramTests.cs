using Automation.GenerateIgxl.PostAction.SelSram;
using Automation.Singleton;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class InputSelSramTests : FunctionTestBase
    {
        [TestMethod]
        public void GetReadbackOtherPatHeaderPos_ShouldFindCorrectColumns()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("TestSheet");

            sheet.Cells[1, 3].Value = "Other Patterns";
            sheet.Cells[1, 5].Value = "SELSRAM bits";
            sheet.Cells[1, 7].Value = "SELSRAM DigCap Readback Pattern";
            sheet.Cells[2, 3].Value = "Other Patterns";
            sheet.Cells[3, 3].Value = "HardIP_*";
            sheet.Cells[4, 4].Value = "SELSRAM DigCap Readback Pattern";

            var target = new InputSelSram();

            // Act
            target.GetReadbackOtherPatHeaderPos(sheet);

            Assert.AreEqual(3, target._otherPatternsPos);
            Assert.AreEqual(5, target._selSramBitsPos);
            Assert.AreEqual(7, target._selSramDigcapPos);
            Assert.AreEqual(2, target._otherPatStartRowPos);

            var selSram = SelSramPatternSingleton.GetInstance();
            selSram.DicOtherPat.Clear();

            target.AddingOtherPat2Dic(sheet);

            string expectedKey = "Other Patterns";
            Assert.IsTrue(selSram.DicOtherPat.ContainsKey(expectedKey.ToUpper()));
        }
    }
}
