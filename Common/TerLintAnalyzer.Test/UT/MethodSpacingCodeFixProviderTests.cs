using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.MethodSpacing;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class MethodSpacingCodeFixProviderTests : CSharpCodeFixVerifier<MethodSpacingAnalyzer, MethodSpacingCodeFixProvider>
    {
        [TestMethod]
        public async Task MissingBlankLine_FixedToOneBlankLine()
        {
            // Case 1: No space between methods
            string test = @"
    class Test
    {
        void MethodA() { }
        void [|MethodB|]() { }
    }";

            string fixedCode = @"
    class Test
    {
        void MethodA() { }

        void MethodB() { }
    }";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task ExcessiveBlankLines_FixedToOneBlankLine()
        {
            // Case 2: Three newlines (two blank lines)
            string test = @"
    class Test
    {
        void MethodA() { }


        void [|MethodB|]() { }
    }";

            string fixedCode = @"
    class Test
    {
        void MethodA() { }

        void MethodB() { }
    }";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task MethodWithComments_PreservesCommentsAndFixesSpacing()
        {
            // Case 3: No space before a method that has a comment
            string test = @"
    class Test
    {
        void MethodA() { }
        // This is a comment
        void [|MethodB|]() { }
    }";

            // Your FixSpacingAsync logic adds a NewLine, then the existing comment
            string fixedCode = @"
    class Test
    {
        void MethodA() { }

        // This is a comment
        void MethodB() { }
    }";

            await VerifyCodeFixAsync(test, fixedCode);
        }
    }
}
