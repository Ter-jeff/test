using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.PreAction.GenDSSCSetup;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.PreAction
{
    [TestClass]
    public class DsscSetupMainTests
    {
        [TestInitialize]
        public void Init()
        {
            DsscSetupMain.Initialize();
        }

        // ----------------------------------------------------
        // ParseSgmtSizes
        // ----------------------------------------------------

        [TestMethod]
        public void ParseSgmtSizes_EmptyString_ShouldReturnEmptyList()
        {
            List<int> sizes = DsscItem.ParseSgmtSizes(string.Empty);
            Assert.AreEqual(0, sizes.Count);
        }

        [TestMethod]
        public void ParseSgmtSizes_InvalidPattern_ShouldReturnEmptyList()
        {
            List<int> sizes = DsscItem.ParseSgmtSizes("AAA BBB CCC");
            Assert.AreEqual(0, sizes.Count);
        }

        [TestMethod]
        public void ParseSgmtSizes_MixedValidInvalid_ShouldOnlyReturnValidOnes()
        {
            List<int> sizes = DsscItem.ParseSgmtSizes("sgmt1_4 abc sgmt2_x sgmt3_7");
            Assert.AreEqual(2, sizes.Count);
            Assert.AreEqual(4, sizes[0]);
            Assert.AreEqual(7, sizes[1]);
        }

        // ----------------------------------------------------
        // ParseDigSrcAssignment
        // ----------------------------------------------------

        [TestMethod]
        public void ParseDigSrcAssignment_Empty_ShouldReturnEmpty()
        {
            List<DigSrcRegAssi> list = DsscItem.ParseDigSrcAssignment(string.Empty);
            Assert.AreEqual(0, list.Count);
        }

        [TestMethod]
        public void ParseDigSrcAssignment_MissingValue_ShouldStillParseKey()
        {
            List<DigSrcRegAssi> list = DsscItem.ParseDigSrcAssignment("A:100;B;");
            Assert.AreEqual(2, list.Count);
            Assert.AreEqual("A", list[0].DigSrcReg);
            Assert.AreEqual("100", list[0].DigSrcAssignment);
            Assert.AreEqual("B", list[1].DigSrcReg);
            Assert.AreEqual("", list[1].DigSrcAssignment);
        }

        [TestMethod]
        public void ParseDigSrcAssignment_OnlyColon_ShouldParseEmptyAssign()
        {
            List<DigSrcRegAssi> list = DsscItem.ParseDigSrcAssignment("R0:");
            Assert.AreEqual(1, list.Count);
            Assert.AreEqual("R0", list[0].DigSrcReg);
            Assert.AreEqual("", list[0].DigSrcAssignment);
        }

        // ----------------------------------------------------
        // Save()
        // ----------------------------------------------------

        [TestMethod]
        public void Save_MultipleSheets_ShouldStoreSeparately()
        {
            DsscSetupSheet sheet1 = new DsscSetupSheet("SETUP1", []);
            DsscSetupSheet sheet2 = new DsscSetupSheet("SETUP2", []);

            DsscSetupMain.Save("AAA", sheet1);
            DsscSetupMain.Save("BBB", sheet2);

            Dictionary<string, List<DsscSetupSheet>> dict = DsscSetupMain.GetSheets();

            Assert.AreEqual(2, dict.Count);
            Assert.AreEqual("SETUP1", dict["AAA"][0].SetupName);
            Assert.AreEqual("SETUP2", dict["BBB"][0].SetupName);
        }

        [TestMethod]
        public void Save_TwoDifferentSetupNames_ShouldBothBeStored()
        {
            DsscSetupSheet sheet1 = new DsscSetupSheet("A", []);
            DsscSetupSheet sheet2 = new DsscSetupSheet("B", []);

            DsscSetupMain.Save("SheetA", sheet1);
            DsscSetupMain.Save("SheetA", sheet2);

            Dictionary<string, List<DsscSetupSheet>> dict = DsscSetupMain.GetSheets();

            Assert.AreEqual(2, dict["SheetA"].Count);
        }

        [TestMethod]
        public void Save_DuplicateSetupName_ShouldNotAddTwice()
        {
            DsscSetupSheet sheet = new DsscSetupSheet("AAA", []);

            DsscSetupMain.Save("Group1", sheet);
            DsscSetupMain.Save("Group1", sheet);

            Dictionary<string, List<DsscSetupSheet>> dict = DsscSetupMain.GetSheets();

            Assert.AreEqual(1, dict["Group1"].Count);
        }

        // ----------------------------------------------------
        // ExportAllSheets + cleanup
        // ----------------------------------------------------

        [TestMethod]
        public void ExportAllSheets_MultipleSheets_ShouldGenerateMultipleFiles()
        {
            List<DsscItem> listA = [new DsscItem("Pat1", "EQ", "R1:AA", "PIN1", "sgmt1_4", "CP1", 3, string.Empty, string.Empty)];

            List<DsscItem> listB = [new DsscItem("Pat2", "EQ2", "R2:BB", "PIN2", "sgmt1_7", "CP2", 5, string.Empty, string.Empty)];

            DsscSetupSheet s1 = new DsscSetupSheet("SetupA", listA);
            DsscSetupSheet s2 = new DsscSetupSheet("SetupB", listB);

            DsscSetupMain.Save("X", s1);
            DsscSetupMain.Save("Y", s2);

            string folder = Path.Combine(Path.GetTempPath(), "DsscMulti");
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }

            List<string> result = DsscSetupMain.ExportAllSheets(folder);
            Assert.AreEqual(2, result.Count);

            string fileX = Path.Combine(folder, "X.txt");
            string fileY = Path.Combine(folder, "Y.txt");

            Assert.IsTrue(File.Exists(fileX));
            Assert.IsTrue(File.Exists(fileY));

            string contentX = File.ReadAllText(fileX);
            string contentY = File.ReadAllText(fileY);

            Assert.IsTrue(contentX.Contains("SetupA"));
            Assert.IsTrue(contentY.Contains("SetupB"));

            // cleanup
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }
        }

        [TestMethod]
        public void ExportAllSheets_NoSheets_ShouldReturnEmptyList()
        {
            string folder = Path.Combine(Path.GetTempPath(), "DsscEmpty");
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }

            List<string> result = DsscSetupMain.ExportAllSheets(folder);

            Assert.AreEqual(0, result.Count);

            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }
        }

        [TestMethod]
        public void ExportAllSheets_SetupWithoutItems_ShouldStillGenerateHeader()
        {
            DsscSetupSheet sheet = new DsscSetupSheet("OnlyHeader", []);
            DsscSetupMain.Save("HeaderSheet", sheet);

            string folder = Path.Combine(Path.GetTempPath(), "DsscHeader");
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }

            DsscSetupMain.ExportAllSheets(folder);

            string file = Path.Combine(folder, "HeaderSheet.txt");
            string content = File.ReadAllText(file);

            Assert.IsTrue(content.StartsWith("DsscSetup"));

            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }
        }

        [TestMethod]
        public void ExportAllSheets_ItemNoDigSrcRegAssi_ShouldPrintBlankColumns()
        {
            DsscItem item = new DsscItem(
                "PAT",
                "EQ",
                string.Empty,
                "PIN",
                "sgmt1_4",
                "CP",
                1,
                string.Empty,
                string.Empty
            );

            DsscSetupSheet sheet = new DsscSetupSheet("NoReg", [item]);
            DsscSetupMain.Save("SheetNoReg", sheet);

            string folder = Path.Combine(Path.GetTempPath(), "DsscNoReg");
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }

            DsscSetupMain.ExportAllSheets(folder);

            string file = Path.Combine(folder, "SheetNoReg.txt");
            string content = File.ReadAllText(file);

            Assert.IsTrue(content.Contains("\t\t"));

            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }
        }
    }
}
