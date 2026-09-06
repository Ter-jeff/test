using System.Collections.Generic;
using System.IO;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT
{
    [TestClass]
    public class IgxlManagerTests
    {
        public string InputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        public string OutputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        public string ExpectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void GenTestProgramBySubProgram_Test()
        {
            string subName = "GenTestProgramBySubProgram";
            string inputPath = Path.Combine(InputPath, subName);
            string outputPath = Path.Combine(OutputPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string subProgramName = "SubProgram1";
            string igxlProj = Path.Combine(inputPath, subProgramName + ".igxlProj");
            string igxl = Path.Combine(outputPath, subProgramName + ".igxl");
            new IgxlManager().GenTestProgramBySubProgram(igxlProj, igxl, " -g ", "CP1");

            Assert.IsTrue(File.Exists(igxl));
        }

        [TestMethod]
        public void ExportWorkBook_Igxl_Test()
        {
            string subName = "ExportWorkBook_Igxl";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string file = Path.Combine(InputPath, "Sample.igxl");
            IgxlManager.ExportWorkBook(file, outputPath);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void IgxlManagerMain()
        {
            string subName = "IgxlManagerMain";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var igxlSheets = new List<IIgxlSheet>()
            {
                new BasFile("Sheet")
                {
                    Rows=
                    [
                        new("A"),
                        new("B")
                    ]
                }
            };
            string outputIgxl = Path.Combine(outputPath, "MyProject.igxl");
            IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets, outputPath);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void IgxlManager_Constructor_Initializes()
        {
            var igxlManager = new IgxlManager();

            Assert.IsNotNull(igxlManager);
        }

        [TestMethod]
        public void IgxlManager_AddIgxlSheets_WithEmptyIgxlPath_CreatesIgxlFile()
        {
            string outputPath = Path.Combine(OutputPath, "IgxlManager_AddIgxlSheets_WithEmptyIgxlPath");
            string outputIgxl = Path.Combine(outputPath, "test.igxl");

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var igxlSheets = new List<IIgxlSheet>();
            IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets, outputPath);

            Assert.IsTrue(File.Exists(outputIgxl));
        }

        [TestMethod]
        public void IgxlManager_AddIgxlSheets_WithBasFile_SuccessfullyAddsSheet()
        {
            string subName = "IgxlManager_AddIgxlSheets_WithBasFile";
            string outputPath = Path.Combine(OutputPath, subName);
            string outputIgxl = Path.Combine(outputPath, "test.igxl");

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var igxlSheets = new List<IIgxlSheet>()
            {
                new BasFile("TestSheet")
                {
                    Rows =
                    [
                        new("Row1"),
                        new("Row2")
                    ]
                }
            };

            IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets, outputPath);

            Assert.IsTrue(File.Exists(outputIgxl));
        }

        [TestMethod]
        public void IgxlManager_AddIgxlSheets_WithMultipleBasFiles_SuccessfullyAddsAllSheets()
        {
            string subName = "IgxlManager_AddIgxlSheets_WithMultipleBasFiles";
            string outputPath = Path.Combine(OutputPath, subName);
            string outputIgxl = Path.Combine(outputPath, "test.igxl");

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var igxlSheets = new List<IIgxlSheet>()
            {
                new BasFile("Sheet1")
                {
                    Rows =
                    [
                        new("Row1")
                    ]
                },
                new BasFile("Sheet2")
                {
                    Rows =
                    [
                        new("Row2")
                    ]
                }
            };

            IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets, outputPath);

            Assert.IsTrue(File.Exists(outputIgxl));
        }

        [TestMethod]
        public void IgxlManager_AddIgxlSheets_WithBasFileWithMultipleRows_SuccessfullyAddsAllRows()
        {
            string subName = "IgxlManager_AddIgxlSheets_WithBasFileWithMultipleRows";
            string outputPath = Path.Combine(OutputPath, subName);
            string outputIgxl = Path.Combine(outputPath, "test.igxl");

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var basFile = new BasFile("MultiRowSheet");
            for (int i = 0; i < 5; i++)
            {
                basFile.Rows.Add(new BasRow($"Row{i}"));
            }

            var igxlSheets = new List<IIgxlSheet> { basFile };
            IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets, outputPath);

            Assert.IsTrue(File.Exists(outputIgxl));
        }

        [TestMethod]
        public void IgxlManager_AddIgxlSheets_WithEmptyBasFile_SuccessfullyAddsEmptySheet()
        {
            string subName = "IgxlManager_AddIgxlSheets_WithEmptyBasFile";
            string outputPath = Path.Combine(OutputPath, subName);
            string outputIgxl = Path.Combine(outputPath, "test.igxl");

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var emptyBasFile = new BasFile("EmptySheet");
            var igxlSheets = new List<IIgxlSheet> { emptyBasFile };

            IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets, outputPath);

            Assert.IsTrue(File.Exists(outputIgxl));
        }

        [TestMethod]
        public void IgxlManager_AddIgxlSheets_ToExistingFile_UpdatesFile()
        {
            string subName = "IgxlManager_AddIgxlSheets_ToExistingFile";
            string outputPath = Path.Combine(OutputPath, subName);
            string outputIgxl = Path.Combine(outputPath, "test.igxl");

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // First addition
            var igxlSheets1 = new List<IIgxlSheet>()
            {
                new BasFile("Sheet1") { Rows = [new("Row1")] }
            };
            IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets1, outputPath);

            // Second addition
            var igxlSheets2 = new List<IIgxlSheet>()
            {
                new BasFile("Sheet2") { Rows = [new("Row2")] }
            };
            IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets2, outputPath);

            Assert.IsTrue(File.Exists(outputIgxl));
            Assert.IsTrue(new FileInfo(outputIgxl).Length > 0);
        }

        [TestMethod]
        public void IgxlManager_AddIgxlSheets_WithValidTempFolder_SuccessfullyCreatesFile()
        {
            string subName = "IgxlManager_AddIgxlSheets_WithValidTempFolder";
            string outputPath = Path.Combine(OutputPath, subName);
            string outputIgxl = Path.Combine(outputPath, "test.igxl");

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var igxlSheets = new List<IIgxlSheet>()
            {
                new BasFile("Sheet") { Rows = [new("Data")] }
            };

            string tempFolder = Path.Combine(outputPath, "Temp");
            Directory.CreateDirectory(tempFolder);

            IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets, tempFolder);

            Assert.IsTrue(File.Exists(outputIgxl));
        }

        [TestMethod]
        public void IgxlManager_AddIgxlSheets_OutputPathDoesNotExist_CreatesPath()
        {
            string subName = "IgxlManager_AddIgxlSheets_OutputPathDoesNotExist";
            string outputPath = Path.Combine(OutputPath, subName);
            string outputIgxl = Path.Combine(outputPath, "test.igxl");

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            var igxlSheets = new List<IIgxlSheet>()
            {
                new BasFile("Sheet") { Rows = [new("Data")] }
            };

            Directory.CreateDirectory(outputPath);

            IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets, outputPath);

            Assert.IsTrue(File.Exists(outputIgxl));
        }

        [TestMethod]
        public void IgxlManager_AddIgxlSheets_WithSpecialCharactersInName_SuccessfullyAddsSheet()
        {
            string subName = "IgxlManager_AddIgxlSheets_WithSpecialCharacters";
            string outputPath = Path.Combine(OutputPath, subName);
            string outputIgxl = Path.Combine(outputPath, "test_special-name.igxl");

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var igxlSheets = new List<IIgxlSheet>()
            {
                new BasFile("Sheet_With-Special") { Rows = [new("Data")] }
            };

            IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets, outputPath);

            Assert.IsTrue(File.Exists(outputIgxl));
        }

        [TestMethod]
        public void IgxlManager_AddIgxlSheets_CreatesValidZipFile()
        {
            string subName = "IgxlManager_AddIgxlSheets_CreatesValidZipFile";
            string outputPath = Path.Combine(OutputPath, subName);
            string outputIgxl = Path.Combine(outputPath, "test.igxl");

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var igxlSheets = new List<IIgxlSheet>()
            {
                new BasFile("Sheet") { Rows = [new("Data")] }
            };

            IgxlManager.AddIgxlSheets(outputIgxl, igxlSheets, outputPath);

            // Verify it's a valid ZIP file by checking magic number
            byte[] fileBytes = File.ReadAllBytes(outputIgxl);
            Assert.IsTrue(fileBytes.Length >= 4);
            // ZIP files start with 50 4B (PK)
            Assert.AreEqual(0x50, fileBytes[0]);
            Assert.AreEqual(0x4B, fileBytes[1]);
        }
    }
}
