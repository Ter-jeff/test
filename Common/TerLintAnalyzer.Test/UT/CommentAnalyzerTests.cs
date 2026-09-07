using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class CommentAnalyzerTests : CSharpAnalyzerVerifier<CommentAnalyzer>
    {
        [TestMethod]
        public async Task TrailingComment_ReportsDiagnostic()
        {
            // Comment on the same line as a semicolon
            string test = @"
    class Test
    {
        void Method()
        {
            int x = 1; [|// This is a trailing comment|]
        }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task CommentOnOwnLine_NoDiagnostic()
        {
            // Comment is on its own line (has leading trivia/newlines)
            string test = @"
    class Test
    {
        void Method()
        {
            // This is a valid comment
            int x = 1;
        }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task MultiLineComment_NoDiagnostic()
        {
            // Your logic specifically checks for SingleLineCommentTrivia
            string test = @"
    class Test
    {
        void Method()
        {
            int x = 1; /* Multi-line trailing */
        }
    }";
            await VerifyAnalyzerAsync(test);
        }
    }
}
