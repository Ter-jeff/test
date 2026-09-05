using System.Collections.Generic;
using System.IO;

using Automation.Reader;
using Automation.Singleton;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.PatternListCsvFile;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class BasicTests : FunctionTestBase
    {
        [TestMethod]
        public void CompileCsvTest()
        {
            string subName = "CompileCsvTest";
            string outputPath = Path.Combine(OutputPath, "Basic", subName);
            string expectPath = Path.Combine(ExpectPath, "Basic", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string compiledPatFile = Path.Combine(KPath, "borneo", "A0_V04A", "CompilePat", "borneo_CompiledPat.csv");
            string patternFolder = Path.Combine(KPath, "borneo", "A0_V04A", "patx");
            string timeSetFolder = Path.Combine(KPath, "borneo", "A0_V04A", "TimeSet");
            string csv = Path.Combine(InputPath, "borneo_documents", "A0_V04A", "borneo_A0_pattern_dashboard_20251002_1524_15.csv");
            PatternListSingleton.Initialize();
            var patListCsv = new InputPatternListCsv(new FileInfo(csv));
            patListCsv.Compile(compiledPatFile, patternFolder, timeSetFolder, true, out Dictionary<string, CompileItem> _, null);
            string compiled = patListCsv.FullName;
            string outputFile = Path.Combine(outputPath, Path.GetFileNameWithoutExtension(compiled) + ".txt");
            File.Copy(compiled, outputFile);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
