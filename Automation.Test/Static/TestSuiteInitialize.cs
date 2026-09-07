using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.PreAction.ReadBasLib;
using Automation.Static;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.VbtLib;

namespace Automation.Test.Static
{
    [TestClass]
    public static class TestSuiteInitialize
    {
        public static List<Function> Functions { get; private set; } = null!;
        private static readonly string _kPath = Path.Combine(Directory.GetCurrentDirectory(), "K");

        [AssemblyInitialize]
        public static void AssemblyInit(TestContext testContext)
        {
            LocalSpecs.BasLibraryFolder = Path.Combine(_kPath, "central_library_vb");
            LocalSpecs.CsLibraryFolder = Path.Combine(_kPath, "Lib");
            var basMain = new BasMain();
            Functions = basMain.ReadLib(LocalSpecs.CsLibraryFolder, LocalSpecs.BasLibraryFolder);
        }

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
        }
    }
}
