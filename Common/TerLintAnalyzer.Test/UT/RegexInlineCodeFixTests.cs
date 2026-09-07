using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.RegexInline;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class RegexInlineCodeFixTests : CSharpCodeFixVerifier<RegexInlineAnalyzer, RegexInlineCodeFixProvider>
    {
        [TestMethod]
        public async Task StaticRegexIsMatch_ExtractsToField()
        {
            // 1. The code with the issue (marked with [| |])
            string testCode = @"
using System.Text.RegularExpressions;
class MyClass {
    void Method() {
        bool result = [|Regex.IsMatch(""input"", ""^[a-z]+$"")|];
    }
}";

            // 2. The code exactly as it should look after the fix
            string fixedCode = @"
using System.Text.RegularExpressions;
class MyClass {
    private static readonly Regex _regex = new Regex(""^[a-z]+$"");

    void Method() {
        bool result = _regex.IsMatch(""input"");
    }
}";

            await VerifyCodeFixAsync(testCode, fixedCode);
        }

        [TestMethod]
        public async Task StaticRegexIsMatch_WithExistingField_GeneratesUniqueName()
        {
            string testCode = @"
using System.Text.RegularExpressions;
class MyClass {
    private static readonly Regex _regex = new Regex(""abc"");
    void Method() {
        bool result = [|Regex.IsMatch(""input"", ""xyz"")|];
    }
}";

            // Expects '_regex1' because '_regex' is taken
            string fixedCode = @"
using System.Text.RegularExpressions;
class MyClass {
    private static readonly Regex _regex1 = new Regex(""xyz"");
    private static readonly Regex _regex = new Regex(""abc"");
    void Method() {
        bool result = _regex1.IsMatch(""input"");
    }
}";

            await VerifyCodeFixAsync(testCode, fixedCode);
        }
    }
}
