using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Static;
using Automation.Utility.Basic;

using CommonLib.Enums;
using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class VbtToolServiceTests
    {
        private string _tempDir = null!;

        private string _tempDirLib = null!;
        private string _tempDirCodingLib = null!;
        private string _tempDirVbtGenTool = null!;

        [TestInitialize]
        public void Setup()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(_tempDir);

            _tempDirLib = Path.Combine(Path.GetTempPath(), "Lib_" + Path.GetRandomFileName());
            _tempDirCodingLib = Path.Combine(Path.GetTempPath(), "CodingLib_" + Path.GetRandomFileName());
            _tempDirVbtGenTool = Path.Combine(Path.GetTempPath(), "VbtGen_" + Path.GetRandomFileName());

            Directory.CreateDirectory(_tempDirLib);
            Directory.CreateDirectory(_tempDirCodingLib);
            Directory.CreateDirectory(_tempDirVbtGenTool);

            File.WriteAllText(Path.Combine(_tempDirLib, "file1.bas"), "dummy content");
            File.WriteAllText(Path.Combine(_tempDirLib, "file2.bas"), "dummy content");
            File.WriteAllText(Path.Combine(_tempDirCodingLib, "coding.bas"), "dummy content");
            File.WriteAllText(Path.Combine(_tempDirVbtGenTool, "tool.bas"), "dummy content");
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }

            Directory.Delete(_tempDirLib, true);
            Directory.Delete(_tempDirCodingLib, true);
            Directory.Delete(_tempDirVbtGenTool, true);
        }

        private FileInfo CreateTempBasFile(string fileName, params string[] lines)
        {
            string path = Path.Combine(_tempDir, fileName);
            File.WriteAllLines(path, lines);
            return new FileInfo(path);
        }

        [TestMethod]
        [DataRow(EnumDevice.RF, "Public Const AllDCVIPinlist = \"DCVI_Power\"", "Public Const AllPowerPinlist = \"DCVS_Power\"", DisplayName = "01_ModifyCommonBas_Device_RF")]
        [DataRow(EnumDevice.AP, "Public Const AllDCVIPinlist = \"All_DCVI\"", "Public Const AllDCVIPinlist = \"All_Power\"", DisplayName = "02_ModifyCommonBas_Device_AP")]
        public void ModifyCommonBas_ShouldUpdateConstantsBasedOnDevice(EnumDevice enumDevice, string expectedDcvi, string expectedPower)
        {
            // Arrange
            LocalSpecs.Options.Device = enumDevice;
            var lines = new List<string>
            {
                "Public Const AllDCVIPinlist = \"OldValue1\"",
                "Public Const AllPowerPinlist = \"OldValue2\""
            };

            FileInfo commonBas = CreateTempBasFile("LIB_Common_GlobalConstant.bas", [.. lines]);
            var fileInfos = new List<FileInfo> { commonBas };

            // Act
            VbtToolService.ModifyCommonBas(fileInfos);

            // Assert
            List<string> resultLines = [.. File.ReadAllLines(commonBas.FullName)];

            Assert.IsTrue(resultLines.Any(l => l.EqualsIgnoreCase(expectedDcvi)), "DCVI line mismatch");
            Assert.IsTrue(resultLines.Any(l => l.EqualsIgnoreCase(expectedPower)), "Power line mismatch");
        }

        [TestMethod]
        public void GetVbFiles_ShouldReturnLibFiles_WhenDeviceIsAP()
        {
            LocalSpecs.Options.Device = EnumDevice.AP;

            List<FileInfo> files = VbtToolService.GetVbFiles(_tempDirLib, _tempDirCodingLib);

            Assert.AreEqual(2, files.Count);
            Assert.IsTrue(files.Any(f => f.Name == "file1.bas"));
            Assert.IsTrue(files.Any(f => f.Name == "file2.bas"));
        }

        [TestMethod]
        public void GetVbFiles_ShouldIncludeCodingLib_WhenDeviceIsLCD()
        {
            LocalSpecs.Options.Device = EnumDevice.LCD;

            List<FileInfo> files = VbtToolService.GetVbFiles(_tempDirLib, _tempDirCodingLib);

            Assert.AreEqual(3, files.Count);
            Assert.IsTrue(files.Any(f => f.Name == "file1.bas"));
            Assert.IsTrue(files.Any(f => f.Name == "file2.bas"));
            Assert.IsTrue(files.Any(f => f.Name == "coding.bas"));
        }
    }
}
