using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.PreferStringMethods;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class PreferStringMethodsAnalyzerTests : CSharpAnalyzerVerifier<PreferStringMethodsAnalyzer>
    {
        [TestMethod]
        public async Task SimplePattern_ReportsDiagnostic()
        {
            // Simple text should trigger the "Contains/Replace" suggestion
            string test = @"
    using System.Text.RegularExpressions;
    class Test {
        void Method() {
            var r = new Regex([|""SimpleText""|]);
        }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task AnchoredPattern_ReportsDiagnostic()
        {
            // Testing ^ and $ anchors for StartsWith, EndsWith, and Equals
            string test = @"
    using System.Text.RegularExpressions;
    class Test {
        void Method() {
            var r1 = new Regex([|""^Start""|]);
            var r2 = new Regex([|""End$""|]);
            var r3 = new Regex([|""^Exact$""|]);
        }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task RegexMetaCharacters_NoDiagnostic()
        {
            // Patterns with meta-characters like . * + ( ) should be ignored
            string test = @"
    using System.Text.RegularExpressions;
    class Test {
        void Method() {
            var r1 = new Regex(""Simple.Text"");
            var r2 = new Regex(""word+"");
            var r3 = new Regex(@""\d+"");
        }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task NonRegexType_NoDiagnostic()
        {
            // Ensure it doesn't trigger on other types with similar constructor signatures
            string test = @"
    class NotRegex { public NotRegex(string p) {} }
    class Test {
        void Method() {
            var r = new NotRegex(""^Start"");
        }
    }";
            await VerifyAnalyzerAsync(test);
        }
    }
}
