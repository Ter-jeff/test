using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenNonIgxlSheet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.DataStruct;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class SpecialTsValueFileGeneratorTests
    {
        private string _outputDir = null!;

        [TestInitialize]
        public void Setup()
        {
            _outputDir = Path.Combine(Path.GetTempPath(), "SpecialTsValueFileTests");
            if (Directory.Exists(_outputDir))
            {
                Directory.Delete(_outputDir, true);
            }

            _ = Directory.CreateDirectory(_outputDir);
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (Directory.Exists(_outputDir))
            {
                Directory.Delete(_outputDir, true);
            }
        }

        [TestMethod]
        public void EditTestSettingDt_ShouldReturnNull_WhenNoSpecialCategoryFound()
        {
            // Arrange
            TestSettingData ts = CreateTestSettingData(["VDD"], ["1.0", "2.0", "3.0"]);

            var gen = new SpecialTsValueFileGenerator([ts], _outputDir);

            // Act
            TestSettingData? result = gen.EditTestSettingDt(ts);

            // Assert
            Assert.AreEqual(null, result);
        }

        [TestMethod]
        public void EditTestSettingDt_ShouldReturnData_WhenSpecialCategoryDetected()
        {
            // Arrange
            TestSettingData ts = CreateTestSettingData(["VDD"], ["1.0A", "2.0", "3.0"]);

            var gen = new SpecialTsValueFileGenerator([ts], _outputDir);

            // Act
            TestSettingData? result = gen.EditTestSettingDt(ts);

            // Assert
            Assert.AreNotEqual(null, result);
        }

        [TestMethod]
        public void GenerateSpecialTsValueFile_ShouldWriteOutputFile_WhenSpecialCategoriesExist()
        {
            // Arrange
            TestSettingData ts = CreateTestSettingData(["CAT1"], ["MC601 E1 + 10%", "1.0", "2.0"]);

            var gen = new SpecialTsValueFileGenerator([ts], _outputDir);

            // Act
            List<string> files = gen.GenerateSpecialTsValueFile();

            // Assert
            Assert.AreEqual(1, files.Count);
            Assert.IsTrue(File.Exists(files.First()));

            string content = File.ReadAllText(files.First());
            Assert.IsTrue(content.Contains("Rev:"));
            Assert.IsTrue(content.Contains("MC601_E1 + 10%"));
        }

        private static TestSettingData CreateTestSettingData(string[] categoryNames, string[] values)
        {
            var dcCategories = new List<DcCategoryName>();
            for (int i = 0; i < categoryNames.Length; i++)
            {
                dcCategories.Add(new DcCategoryName(categoryNames[i])
                {
                    ColumnIndex = i,
                    ValueType = CategoryValueType.NV
                });
            }

            var row = new TestSettingRow
            {
                PowerPinName = "PIN1",
                DcCategoryValues = [.. categoryNames.Select((c, idx) => new DcCategoryValue(c)
                {
                    ColumnIndex = idx,
                    Nv = new DcCategoryItem { Value = values.ElementAtOrDefault(idx) ?? "0" },
                    Hv = new DcCategoryItem { Value = "0" },
                    Lv = new DcCategoryItem { Value = "0" }
                })]
            };

            return new TestSettingData
            {
                SheetName = "TestSheet",
                TestSettingVersion = "V1",
                DcCategorys = dcCategories,
                DataRows = [row]
            };
        }
    }
}
