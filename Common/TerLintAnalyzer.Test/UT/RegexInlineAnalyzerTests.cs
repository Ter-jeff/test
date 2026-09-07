using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.RegexInline;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class RegexInlineAnalyzerTests : CSharpAnalyzerVerifier<RegexInlineAnalyzer>
    {
        [TestMethod]
        public async Task StaticRegexIsMatch_WithLiteral_ReportsDiagnostic()
        {
            // The code that SHOULD trigger the analyzer
            string test = @"
using System.Text.RegularExpressions;
class Test {
    void Method() {
        bool result = [|Regex.IsMatch(""input"", ""^[a-z]+$"")|];
    }
}";
            // This helper automatically checks that the diagnostic is reported at the [| markup |]
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task StaticRegexIsMatch_WithVariable_DoesNotReportDiagnostic()
        {
            // The code that should NOT trigger the analyzer (dynamic pattern)
            string test = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string pattern) {
        bool result = Regex.IsMatch(""input"", pattern);
    }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task InstanceRegexIsMatch_DoesNotReportDiagnostic()
        {
            string test = @"
using System.Text.RegularExpressions;
class Test {
    void Method() {
        var r = new Regex(""^[a-z]+$"");
        bool result = r.IsMatch(""input""); 
    }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task OtherMethodNamedIsMatch_DoesNotReportDiagnostic()
        {
            string test = @"
class Test {
    void Method() {
        bool result = IsMatch(""input"", ""pattern"");
    }
    bool IsMatch(string a, string b) => true;
}";
            await VerifyAnalyzerAsync(test);
        }
    }

}
