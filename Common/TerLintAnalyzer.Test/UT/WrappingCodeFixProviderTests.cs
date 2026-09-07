using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.Rules;
using TerLintAnalyzer.Ter.Wrapping;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class WrappingCodeFixProviderTests : CSharpCodeFixVerifier<WrappingAnalyzer, WrappingCodeFixProvider>
    {
        [TestMethod]
        public async Task WrappedArguments_CollapsedToSingleLine()
        {
            string test = $@"
class Test {{
    void Foo(int a, int b, string c) {{ }}
    void Method() {{
        {{|{TerRules.GeneralNoWrappingRule.Id}:Foo|}}(1,
            2,
            ""three"");
    }}
}}";

            string fixedCode = @"
class Test {
    void Foo(int a, int b, string c) { }
    void Method() {
        Foo(1, 2, ""three"");
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task WrappedArguments_OpenParenAtEndOfLine_CollapsedWithoutStraySpace()
        {
            // Reproduces the BinCutInputManagerTests.cs pattern: the opening paren is the last
            // character on its line, so the naive whitespace collapse used to leave a stray
            // space right after "(" (and before the matching ")").
            string test = $@"
class Test {{
    void Foo(int a, int b, int c, ref int d) {{ }}
    void Method() {{
        int e = 0;
        {{|{TerRules.GeneralNoWrappingRule.Id}:Foo|}}(
            1,
            2,
            3,
            ref e
        );
    }}
}}";

            string fixedCode = @"
class Test {
    void Foo(int a, int b, int c, ref int d) { }
    void Method() {
        int e = 0;
        Foo(1, 2, 3, ref e);
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task WrappedObjectInitializerArgument_NoDiagnostic()
        {
            // Reproduces the BvComparerBaseHelpers.cs pattern: result.Add(new BvResult { ... })
            // with the object initializer spread across multiple lines. Wrapping a brace
            // initializer across lines is normal, readable style, so it should not be flagged.
            string test = @"
class Result {
    public string PinName;
    public double Voltage;
}
class Test {
    void Method(System.Collections.Generic.List<Result> result) {
        result.Add(new Result()
        {
            PinName = ""a"",
            Voltage = 1.0
        });
    }
}";

            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task WrappedCallOnMemberChain_PreservesSurroundingIndentation()
        {
            string test = $@"
class Printer {{
    public static void PrintDifference(string a, string b, string c) {{ }}
}}
class Test {{
    void Method(string site, string expected, string actual) {{
        if (true)
        {{
            Printer.{{|{TerRules.GeneralNoWrappingRule.Id}:PrintDifference|}}(site,
                expected, actual);
        }}
    }}
}}";

            string fixedCode = @"
class Printer {
    public static void PrintDifference(string a, string b, string c) { }
}
class Test {
    void Method(string site, string expected, string actual) {
        if (true)
        {
            Printer.PrintDifference(site, expected, actual);
        }
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }
    }
}
