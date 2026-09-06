using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.EFuse.Business;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Efuse.Input;

namespace Automation.Test.UT.Efuse
{
    [TestClass]
    public class EfuseGenerateTests
    {
        [TestMethod]
        public void FilterEfuseConfigSheets_Should_Filter_1()
        {
            // Arrange
            var source = new List<EfuseConfigMainSheet>
            {
                new("EFUSE_CONFIG_MAIN_A"),
                new("EFUSE_CONFIG_MAIN_B"),
                new("EFUSE_CONFIG_MAIN_C"),
                new("EFUSE_CONFIG_MAIN_D")
            };

            string configType = "Config";
            bool hasMultipleBlock = false;

            // Act
            List<EfuseConfigMainSheet> result = EfuseGenerate.FilterEfuseConfigSheets(source, configType, hasMultipleBlock);

            // Assert
            Assert.AreEqual(source.Count, result.Count);
        }

        [TestMethod]
        public void FilterEfuseConfigSheets_Should_Filter_2()
        {
            // Arrange
            var source = new List<EfuseConfigMainSheet>
            {
                new("EFUSE_CONFIG_MAIN_A_GFX"),
                new("EFUSE_CONFIG_MAIN_B_GFX"),
                new("EFUSE_CONFIG_MAIN_C_GFX"),
                new("EFUSE_CONFIG_MAIN_D_GFX"),
                new("EFUSE_CONFIG_MAIN_A_CPU"),
                new("EFUSE_CONFIG_MAIN_B_CPU"),
                new("EFUSE_CONFIG_MAIN_C_CPU"),
                new("EFUSE_CONFIG_MAIN_D_CPU"),
            };

            string configType = "CPU";
            bool hasMultipleBlock = true;
            // Act
            List<EfuseConfigMainSheet> result = EfuseGenerate.FilterEfuseConfigSheets(source, configType, hasMultipleBlock);

            // Assert
            Assert.AreEqual(4, result.Count);
            Assert.IsTrue(result.All(x => x.SheetName.Contains("CPU")));
        }

        [TestMethod]
        public void FilterEfuseConfigSheets_Should_Filter_3()
        {
            // Arrange
            var source = new List<EfuseConfigMainSheet>
            {
                new("EFUSE_CONFIG_MAIN_A_GFX"),
                new("EFUSE_CONFIG_MAIN_B_GFX"),
                new("EFUSE_CONFIG_MAIN_C_GFX"),
                new("EFUSE_CONFIG_MAIN_D_GFX"),
                new("EFUSE_CONFIG_MAIN_A_CPU"),
                new("EFUSE_CONFIG_MAIN_B_CPU"),
                new("EFUSE_CONFIG_MAIN_C_CPU"),
                new("EFUSE_CONFIG_MAIN_D_CPU"),
            };
            string configType = "GFX";
            bool hasMultipleBlock = true;
            // Act
            List<EfuseConfigMainSheet> result = EfuseGenerate.FilterEfuseConfigSheets(source, configType, hasMultipleBlock);

            // Assert
            Assert.AreEqual(4, result.Count);
            Assert.IsTrue(result.All(x => x.SheetName.Contains("GFX")));
        }
    }
}
