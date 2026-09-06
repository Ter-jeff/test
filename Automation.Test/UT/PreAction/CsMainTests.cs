using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

using Automation.GenerateIgxl.PreAction.ReadBasLib;

using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.VbtLib;

namespace Automation.Test.UT.PreAction
{
    [TestClass]
    public class CsMainTests
    {
        private string _tempRoot = null!;
        private CsMain _csMain = null!;

        [TestInitialize]
        public void Setup()
        {
            _tempRoot = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(_tempRoot);
            _csMain = new CsMain();
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (Directory.Exists(_tempRoot))
            {
                Directory.Delete(_tempRoot, true);
            }
        }

        [TestMethod]
        public void WorkFlow()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Assert.Inconclusive("This test requires Windows OS: requires vswhere.exe and MSBuild.exe");
            }

            // Arrange
            string path = Path.Combine(Directory.GetCurrentDirectory(), "Input", "PreAction", "Automation.sln");
            string dllFolderPath = Path.Combine(_tempRoot, "central_library_cs", "bin");
            Directory.CreateDirectory(dllFolderPath);

            // Act
            (List<Function> functions, ReferenceSheet? referenceSheet) = _csMain.WorkFlow(_tempRoot, path);

            // Assert
            Assert.AreEqual(0, functions.Count);
            Assert.AreNotEqual(null, referenceSheet);
        }

        [TestMethod]
        public void WorkFlow_ShouldReturnEmpty_WhenInvalidPath()
        {
            // Arrange
            string invalidPath = Path.Combine(_tempRoot, "notexist");

            // Act
            (List<Function> functions, ReferenceSheet? referenceSheet) = _csMain.WorkFlow(_tempRoot, invalidPath);

            // Assert
            Assert.AreEqual(0, functions.Count);
            Assert.AreEqual(null, referenceSheet);
        }

        [TestMethod]
        public void WorkFlow_ShouldHandleDllFolder()
        {
            // Arrange
            string dllDir = Path.Combine(_tempRoot, "MyDlls");
            Directory.CreateDirectory(dllDir);

            // Act
            (List<Function> functions, ReferenceSheet? _) = _csMain.WorkFlow(_tempRoot, dllDir);

            // Assert
            Assert.AreEqual(0, functions.Count);
        }

        [TestMethod]
        public void CopyDirectory_ShouldCopyAllFilesAndSubfolders()
        {
            // Arrange
            string source = Path.Combine(_tempRoot, "src");
            string dest = Path.Combine(_tempRoot, "dst");
            Directory.CreateDirectory(source);
            File.WriteAllText(Path.Combine(source, "a.txt"), "hello");
            Directory.CreateDirectory(Path.Combine(source, "sub"));
            File.WriteAllText(Path.Combine(source, "sub", "b.txt"), "world");

            // Act
            _csMain.CopyDirectory(source, dest, true, false);

            // Assert
            Assert.IsTrue(File.Exists(Path.Combine(dest, "a.txt")));
            Assert.IsTrue(File.Exists(Path.Combine(dest, "sub", "b.txt")));
        }

        [TestMethod]
        public void ReadLocalLib_ShouldThrow_WhenFolderMissing()
        {
            // Arrange
            string missingFolder = Path.Combine(_tempRoot, "nope");
            Assert.ThrowsException<Exception>(() =>
            {
                CsMain.ReadLocalLib(missingFolder);
            });
        }

        [TestMethod]
        public void ReadLocalLib_ShouldReturnEmpty_WhenFolderExistsNoDll()
        {
            // Arrange
            string emptyFolder = Path.Combine(_tempRoot, "dlls");
            Directory.CreateDirectory(emptyFolder);

            // Act
            (List<Function> functions, ReferenceSheet? refSheet) = CsMain.ReadLocalLib(emptyFolder);

            // Assert
            Assert.AreNotEqual(null, functions);
            Assert.AreNotEqual(null, refSheet);
            Assert.AreEqual(0, functions.Count);
            Assert.AreEqual(0, refSheet!.Rows.Count);
        }
    }
}
