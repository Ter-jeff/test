using System;
using System.Collections.Generic;
using System.IO;

using Automation.Static;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Common
{
    [TestClass]
    [TestCategory("FolderStructure")]
    public class FolderStructureTests
    {
        private string _tempRoot = null!;

        [TestInitialize]
        public void Setup()
        {
            _tempRoot = Path.Combine(Path.GetTempPath(), "FolderStructureTests_" + Guid.NewGuid());
            _ = Directory.CreateDirectory(_tempRoot);

            LocalSpecs.TarFolder = _tempRoot;
            FolderStructure.TarDir = _tempRoot;
        }

        [TestCleanup]
        public void Cleanup()
        {
            try
            {
                if (Directory.Exists(_tempRoot))
                {
                    Directory.Delete(_tempRoot, true);
                }
            }
            catch { /* ignore */ }

            FolderStructure.Clear();
        }

        [DataTestMethod]
        [Description("01_Verify that folder properties correctly return full paths starting with TarDir.")]
        [DataRow(nameof(FolderStructure.DirIgLink), DisplayName = "01_DirIgLink_ReturnsFullPath")]
        [DataRow(nameof(FolderStructure.DirLogs), DisplayName = "02_DirLogs_ReturnsFullPath")]
        [DataRow(nameof(FolderStructure.DirXmlFiles), DisplayName = "03_DirXmlFiles_ReturnsFullPath")]
        [DataRow(nameof(FolderStructure.DirCommon), DisplayName = "04_DirCommon_ReturnsFullPath")]
        [DataRow(nameof(FolderStructure.DirEfuse), DisplayName = "05_DirEfuse_ReturnsFullPath")]
        [DataRow(nameof(FolderStructure.DirHardIp), DisplayName = "06_DirHardIp_ReturnsFullPath")]
        [DataRow(nameof(FolderStructure.DirMixedSignal), DisplayName = "07_DirMixedSignal_ReturnsFullPath")]
        [DataRow(nameof(FolderStructure.DirReference), DisplayName = "08_DirReference_ReturnsFullPath")]
        public void Contact_ShouldReturn_ValidFullPath(string propertyName)
        {
            // Arrange
            var accessors = new Dictionary<string, Func<string>>
            {
                [nameof(FolderStructure.DirIgLink)] = () => FolderStructure.DirIgLink,
                [nameof(FolderStructure.DirLogs)] = () => FolderStructure.DirLogs,
                [nameof(FolderStructure.DirXmlFiles)] = () => FolderStructure.DirXmlFiles,
                [nameof(FolderStructure.DirCommon)] = () => FolderStructure.DirCommon,
                [nameof(FolderStructure.DirEfuse)] = () => FolderStructure.DirEfuse,
                [nameof(FolderStructure.DirHardIp)] = () => FolderStructure.DirHardIp,
                [nameof(FolderStructure.DirMixedSignal)] = () => FolderStructure.DirMixedSignal,
                [nameof(FolderStructure.DirReference)] = () => FolderStructure.DirReference
            };
            Assert.IsTrue(accessors.ContainsKey(propertyName), $"Property {propertyName} should exist.");

            // Act
            string value = accessors[propertyName]();

            // Assert
            Assert.IsFalse(string.IsNullOrEmpty(value), $"{propertyName} returned null or empty.");
            StringAssert.StartsWith(value, _tempRoot, $"{propertyName} should start with TarDir.");
        }

        [TestMethod]
        [Description("02_Verify that CreateFolder() successfully creates all expected directories.")]
        [TestProperty("DisplayName", "08_CreateFolder_CreatesExpectedDirectories")]
        public void CreateFolder_ShouldCreateExpectedDirectories()
        {
            // Act
            FolderStructure.CreateFolder();

            // Assert
            Assert.IsTrue(Directory.Exists(FolderStructure.DirLogs), "Logs folder not created.");
            Assert.IsTrue(Directory.Exists(FolderStructure.DirAcSpec), "AC_Spec folder not created.");
        }

        [TestMethod]
        [Description("03_Verify that Clear() resets static fields correctly.")]
        [TestProperty("DisplayName", "09_Clear_ResetsInternalFields")]
        public void Clear_ShouldResetInternalFields()
        {
            // Arrange
            FolderStructure.TarDir = _tempRoot;

            // Act
            FolderStructure.Clear();

            // Assert
            string value = FolderStructure.TarDir;
            StringAssert.StartsWith(value, _tempRoot);
        }

        [TestMethod]
        [Description("04_Verify that Contact() combines TarDir and subpath properly.")]
        [TestProperty("DisplayName", "10_Contact_CombinesTarDirAndSubPath")]
        public void Contact_ShouldCombineTarDirAndSubPath()
        {
            // Arrange
            string subPath = "SubFolder";

            // Act
            string result = FolderStructure.Contact(subPath);

            // Assert
            Assert.AreEqual(Path.Combine(_tempRoot, subPath), result);
        }

        [TestMethod]
        [Description("Covers all lines in CreateFolder including delete and create blocks.")]
        public void CreateFolder_AllLinesCovered()
        {
            // Arrange
            _ = Directory.CreateDirectory(FolderStructure.DirXmlFiles);
            _ = Directory.CreateDirectory(FolderStructure.DirLogs);

            // Act
            FolderStructure.CreateFolder();

            // Assert
            Assert.IsTrue(Directory.Exists(FolderStructure.DirLogs));
            Assert.IsTrue(Directory.Exists(FolderStructure.DirXmlFiles));

            Assert.IsTrue(Directory.Exists(FolderStructure.DirAcSpec));
        }
    }
}
