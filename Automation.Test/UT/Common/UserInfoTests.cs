using System;
using System.IO;

using CommonLib.Extension;

using CommonReaderLib;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

using OfficeOpenXml;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class UserInfoTests : FunctionTestBase
    {
        [TestMethod]
        public void UserInfoTest()
        {
            AssertOnlyWindowsOS("expected output JSON contains Windows-style backslash paths");
            string subName = "UserInfo";
            string outputPath = Path.Combine(OutputPath, "Common", subName);
            string expectPath = Path.Combine(ExpectPath, "Common", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string inputInfo = Path.Combine(InputPath, "borneo_documents", "borneo_InputInfo.csv");
            using var excelPackage = new ExcelPackage();
            ExcelWorksheet ws = excelPackage.Workbook.Worksheets.Add("InputInfo");
            ws.Cells[1, 1].Value = "Flow Name";
            _ = ws.Cells[1, 1].PrintExcelRowByList(inputInfo.CsvConvertToLists());

            InputInfoSheet? sheet = new InputInfoReader(inputInfo).ReadSheet(ws);
            var userInfo = new UserInfo(sheet);
            userInfo.Initialize();

            // Assert
            string json = JsonConvert.SerializeObject(userInfo, Formatting.Indented);
            string currentDirectory = Directory.GetCurrentDirectory().Replace("\\", "\\\\");
            string modifiedJson = json.Replace(currentDirectory, "");
            modifiedJson = modifiedJson.Replace("/", "\\\\");
            File.WriteAllText(Path.Combine(outputPath, "result.json"), modifiedJson);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void ResolveConfigPath_CurrentValueNonEmpty_ReturnsCurrentValueUnchanged()
        {
            // Act
            string result = UserInfo.ResolveConfigPath("ExistingValue.xlsx", "Z:\\missing", "Proj", "Settings\\{0}.xlsx", "Settings\\Default.xlsx");

            // Assert
            Assert.AreEqual("ExistingValue.xlsx", result);
        }

        [TestMethod]
        public void ResolveConfigPath_NoFilesFound_ReturnsDefaultRelativeToCurrentDirectory()
        {
            // Act
            string result = UserInfo.ResolveConfigPath("", "Z:\\definitely_missing_repo", "Proj", "Settings\\{0}_Missing_Pattern_XYZ.xlsx", "Settings\\Default_Missing_XYZ.xlsx");

            // Assert
            Assert.AreEqual(Path.Combine(Directory.GetCurrentDirectory(), "Settings\\Default_Missing_XYZ.xlsx"), result);
        }

        [TestMethod]
        public void ResolveConfigPath_ProjectFileExistsInRepoFolder_ReturnsProjectFile()
        {
            // Arrange - create a real file in a scratch repo folder matching the project pattern
            string repoFolder = Path.Combine(Path.GetTempPath(), $"UserInfoTests_{Guid.NewGuid():N}");
            Directory.CreateDirectory(repoFolder);
            string projectFile = Path.Combine(repoFolder, "Config_ProjA.xlsx");
            File.WriteAllText(projectFile, "dummy");

            try
            {
                // Act
                string result = UserInfo.ResolveConfigPath("", repoFolder, "ProjA", "Config_{0}.xlsx", "Config_Default.xlsx");

                // Assert
                Assert.AreEqual(projectFile, result);
            }
            finally
            {
                Directory.Delete(repoFolder, true);
            }
        }
    }
}
