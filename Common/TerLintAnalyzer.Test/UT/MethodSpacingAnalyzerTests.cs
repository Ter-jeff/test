using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.MethodSpacing;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class MethodSpacingAnalyzerTests : CSharpAnalyzerVerifier<MethodSpacingAnalyzer>
    {
        [TestMethod]
        public async Task CorrectSpacing_NoDiagnostic()
        {
            // Exactly one blank line (2 newlines) between methods
            string test = @"
    class Test
    {
        void MethodA() { }

        void MethodB() { }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task MissingBlankLine_ReportsDiagnostic()
        {
            // Immediate next line (1 newline) - should fail
            string test = @"
    class Test
    {
        void MethodA() { }
        void [|MethodB|]() { }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task TooManyBlankLines_ReportsDiagnostic()
        {
            // Two blank lines (3 newlines) - should fail
            string test = @"
    class Test
    {
        void MethodA() { }


        void [|MethodB|]() { }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task MethodsWithComments_RespectsSpacing()
        {
            // One blank line above the comment for MethodB
            string test = @"
    class Test
    {
        void MethodA() { }

        // This is a comment
        void MethodB() { }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task MethodsWithComments_ReportsDiagnostic()
        {
            // One blank line above the comment for MethodB
            string test = @"
    class Test
    {
        void MethodA() { }
        // This is a comment
        void [|MethodB|]() { }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task MethodsWithRegion_RespectsSpacing()
        {
            string test = @"
    class Test
    {
        #region AAAA
        void MethodA() { }
        #endregion

        void MethodB() { }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task MethodsWithEndRegionThenNewRegionAndDocComment_RespectsSpacing()
        {
            // Repro: previously the analyzer walked past the #endregion/#region pair and
            // re-accumulated unrelated blank-line gaps from deeper in the trivia, reporting
            // a false positive even though there is exactly one blank line before the
            // #endregion boundary (the region-start marks the real content boundary, so
            // anything after it, like the doc comment, is irrelevant to the count).
            string test = @"
    class Test
    {
        #region MethodA Section
        void MethodA() { }

        #endregion
        #region Read Header

        /// <summary>
        /// ReadHeader
        /// </summary>
        void MethodB() { }
        #endregion
    }";
            await VerifyAnalyzerAsync(test);
        }
    }
}
