using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.ExcessiveSpacing;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class ExcessiveSpacingCodeFixProviderTests : CSharpCodeFixVerifier<ExcessiveSpacingAnalyzer, ExcessiveSpacingCodeFixProvider>
    {
        [TestMethod]
        public async Task TwoBlankLines_CodeFix_RemovesExtraLine()
        {
            const string test = @"class Test
{
    void M()
    {
        int x = 1;

     [|
|]        int y = 2;
    }
}";

            const string fixedCode = @"class Test
{
    void M()
    {
        int x = 1;

        int y = 2;
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task ExcessiveBlankLinesNearEndRegion_CodeFix_PreservesEndRegion()
        {
            // Repro: the old fixer deleted everything between the first and last newline
            // in the trivia list, which destroyed the #endregion directive itself instead
            // of just collapsing the excess blank lines around it.
            const string test = @"class Test
{
    void M1() { }
    #region data
    int x = 1;

     [|
|]        int y = 2;
    #endregion
}";

            const string fixedCode = @"class Test
{
    void M1() { }
    #region data
    int x = 1;

        int y = 2;
    #endregion
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }
    }
}
