using System;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.PostAction.GenIgxlProj;
using Automation.Static;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.PostAction
{
    /// <summary>
    /// Covers <see cref="UfpBasMutator"/>, which injects <c>#Const isUFP = True/False</c>
    /// into .bas sources before the IGXL packager runs. Inherits <see cref="FunctionTestBase"/>
    /// so <c>TestPlanStatic.Equipments</c> (and its dependencies) resolve deterministically.
    /// </summary>
    [TestClass]
    public class UfpBasMutatorTests : FunctionTestBase
    {
        private static readonly string[] _dspFileLines = ["Attribute VB_Name = \"DSP\"", "Sub Foo()"];
        private static readonly string[] _plainFileLines = ["row1", "row2"];

        private string _dir = null!;

        [TestInitialize]
        public void Setup()
        {
            _dir = Path.Combine(Path.GetTempPath(), $"ufptest-{Guid.NewGuid():N}");
            Directory.CreateDirectory(_dir);
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (Directory.Exists(_dir))
            {
                Directory.Delete(_dir, true);
            }
        }

        private static string ExpectedConstLine()
        {
            return TestPlanStatic.Equipments.Contains(EnumEquipment.UltraFlexPlus)
                ? "#Const isUFP = True"
                : "#Const isUFP = False";
        }

        private string WriteBas(string name, params string[] lines)
        {
            string path = Path.Combine(_dir, name);
            File.WriteAllLines(path, lines);
            return path;
        }

        [TestMethod]
        public void Apply_ShouldInsertConstAtLineTwo_WhenAbsent()
        {
            string path = WriteBas("Module.bas", "Attribute VB_Name = \"Module\"", "Sub Foo()");

            UfpBasMutator.Apply([path]);

            string[] result = File.ReadAllLines(path);
            Assert.AreEqual("Attribute VB_Name = \"Module\"", result[0]);
            Assert.AreEqual(ExpectedConstLine(), result[1]);
        }

        [TestMethod]
        public void Apply_ShouldReplaceExistingConst()
        {
            // Existing value is the opposite spelling so the mutator must rewrite it.
            string path = WriteBas("Module.bas", "Attribute VB_Name = \"Module\"", "#Const isUFP = Maybe", "Sub Foo()");

            UfpBasMutator.Apply([path]);

            string text = File.ReadAllText(path);
            Assert.IsTrue(text.Contains(ExpectedConstLine()));
            // Exactly one directive remains.
            Assert.AreEqual(1, File.ReadAllLines(path).Count(l => l.Contains("#Const isUFP = ")));
        }

        [TestMethod]
        public void Apply_ShouldSkipDspFiles()
        {
            string path = WriteBas("DSP_Kernel.bas", "Attribute VB_Name = \"DSP\"", "Sub Foo()");

            UfpBasMutator.Apply([path]);

            CollectionAssert.AreEqual(_dspFileLines, File.ReadAllLines(path));
        }

        [TestMethod]
        public void Apply_ShouldIgnoreNonBasFiles()
        {
            string path = WriteBas("Sheet.txt", "row1", "row2");

            UfpBasMutator.Apply([path]);

            CollectionAssert.AreEqual(_plainFileLines, File.ReadAllLines(path));
        }
    }
}
