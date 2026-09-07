using System.Threading.Tasks;

using Microsoft.CodeAnalysis.Testing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class StringAnalyzerTests : CSharpAnalyzerVerifier<StringAnalyzer>
    {
        [TestMethod]
        public async Task VariableIsAssigned_NoDiagnostic1()
        {
            string test = @"
using System;
class Test {
    void Method(string s1, string s2) {
bool result = s1.ToUpper() == s2;
    }
}";
            DiagnosticResult expected = Diagnostic(StringRules.ToUpperToLowerEqualityOperator).WithSpan(5, 15, 5, 33).WithArguments("ToUpper", "==", "");
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task StaticEquals_Ordinal_ReportsDiagnostic()
        {
            string test = @"
            class Test {
                void Method(string a, string b) {
                    bool r = string.Equals(a, b, System.StringComparison.Ordinal);
                }
            }";
            DiagnosticResult expected = Diagnostic(StringRules.StringEqualsCaseSensitive).WithSpan(4, 30, 4, 82).WithArguments("Ordinal");
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task StaticEquals_OrdinalIgnoreCase_ReportsDiagnostic()
        {
            string test = @"
            class Test {
                void Method(string a, string b) {
                    bool r = string.Equals(a, b, System.StringComparison.OrdinalIgnoreCase);
                }
            }";
            DiagnosticResult expected = Diagnostic(StringRules.StringEqualsCaseInsensitive).WithSpan(4, 30, 4, 92).WithArguments("OrdinalIgnoreCase");
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task InstanceEquals_NoArgs_ReportsDiagnostic()
        {
            string test = @"
            class Test {
                void Method(string a, string b) {
                    bool r = a.Equals(b);
                }
            }";
            DiagnosticResult expected = Diagnostic(StringRules.EqualsCaseSensitive).WithSpan(4, 30, 4, 41);
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task InstanceEquals_IgnoreCase_ReportsDiagnostic()
        {
            string test = @"
            class Test {
                void Method(string a, string b) {
                    bool r = a.Equals(b, System.StringComparison.CurrentCultureIgnoreCase);
                }
            }";
            DiagnosticResult expected = Diagnostic(StringRules.EqualsCaseInsensitive).WithSpan(4, 30, 4, 91).WithArguments("CurrentCultureIgnoreCase");
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task ToUpper_ChainedWithContains_ReportsDiagnostic()
        {
            string test = @"
            class Test {
                void Method(string a) {
                    bool r = a.ToUpper().Contains(""XYZ"");
                }
            }";
            DiagnosticResult expected = Diagnostic(StringRules.ToUpperToLowerComparison).WithSpan(4, 30, 4, 41).WithArguments("ToUpper", "Contains");
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task ToLower_WithEqualityOperator_ReportsDiagnostic()
        {
            string test = @"
            class Test {
                void Method(string a, string b) {
                    bool r = a.ToLower() == b;
                }
            }";
            DiagnosticResult expected = Diagnostic(StringRules.ToUpperToLowerEqualityOperator).WithSpan(4, 30, 4, 46).WithArguments("ToLower", "==", "");
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task StringComparer_PropertyAccess_ReportsDiagnostic()
        {
            string test = @"
            using System;
            class Test {
                void Method() {
                    var c = StringComparer.OrdinalIgnoreCase;
                }
            }";
            DiagnosticResult expected = Diagnostic(StringRules.StringComparerUsage).WithSpan(5, 29, 5, 61);
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task StringComparer_IndexOf_ReportsDiagnostic()
        {
            string test = @"
            using System;
            class Test {
                void Method() {
                    if (""srcFile"".ToLower().IndexOf("".txt"", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                    }
                }
            }";
            DiagnosticResult expected = Diagnostic(StringRules.ToUpperToLowerComparison).WithSpan(5, 25, 5, 44).WithArguments("ToLower", "IndexOf");
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task InstanceEquals_null_ReportsDiagnostic()
        {
            string test = @"
            class Test {
                void Method(string a, string b) {
                    bool r = a == null;
                }
            }";
            DiagnosticResult expected = Diagnostic(StringRules.StringIsNullOrEmptyUsage).WithSpan(4, 30, 4, 39).WithArguments("a", "string.IsNullOrEmpty(a)");
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task PropertyNullCheck_InLoop_ReportsDiagnostic()
        {
            string test = @"
class PatternRow
{
    public string PatternSet { get; set; }
}

class Test
{
    void Process(PatternRow patRow)
    {
        while (true)
        {
            if (patRow.PatternSet == null) { continue; }
        }
    }
}";

            DiagnosticResult expected = Diagnostic(StringRules.StringIsNullOrEmptyUsage).WithSpan(13, 17, 13, 42).WithArguments("patRow.PatternSet", "string.IsNullOrEmpty(patRow.PatternSet)");

            await VerifyAnalyzerAsync(test, expected);
        }
    }
}
