using System.Threading.Tasks;

using Microsoft.CodeAnalysis.Testing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.PathSeparator;
using TerLintAnalyzer.Ter.Rules;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class PathSeparatorAnalyzerTests : CSharpAnalyzerVerifier<PathSeparatorAnalyzer>
    {
        [TestMethod]
        public async Task BackslashInPath_ReportsTer402()
        {
            string test = @"
class Test {
    void Method() {
        string path = ""Settings\\Basic\\Config.txt"";
    }
}";
            DiagnosticResult expected = Diagnostic(TerRules.PathBackslashInLiteral).WithSpan(4, 23, 4, 52);
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task DriveLetterPath_ReportsTer402()
        {
            string test = @"
class Test {
    void Method() {
        string path = ""C:\\Users\\test\\file.txt"";
    }
}";
            DiagnosticResult expected = Diagnostic(TerRules.PathBackslashInLiteral).WithSpan(4, 23, 4, 50);
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task RelativePath_ReportsTer402()
        {
            string test = @"
class Test {
    void Method() {
        string path = "".\\output\\file.txt"";
    }
}";
            DiagnosticResult expected = Diagnostic(TerRules.PathBackslashInLiteral).WithSpan(4, 23, 4, 44);
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task UncPath_NoDiagnostic()
        {
            string test = @"
class Test {
    void Method() {
        string path = ""\\\\server\\share\\file.txt"";
    }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task RegistryKey_NoDiagnostic()
        {
            string test = @"
class Test {
    void Method() {
        string key = ""SOFTWARE\\Microsoft\\Windows"";
    }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task RegexPattern_NoDiagnostic()
        {
            string test = @"
class Test {
    void Method() {
        string pattern = ""\\d+\\.\\w+"";
    }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task SingleBackslash_NoDiagnostic()
        {
            string test = @"
class Test {
    void Method() {
        string sep = ""\\"";
    }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task ReplaceContext_NoDiagnostic()
        {
            string test = @"
class Test {
    void Method(string path) {
        string result = path.Replace(""/"", ""\\"");
    }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task PathCombineFirstArg_NoDiagnostic()
        {
            string test = @"
using System.IO;
class Test {
    void Method() {
        string result = Path.Combine(""C:\\base"", ""file.txt"");
    }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task ConcatenationWithBackslash_ReportsTer403()
        {
            string test = @"
class Test {
    void Method(string folder) {
        string path = folder + ""\\"" + ""file.txt"";
    }
}";
            DiagnosticResult expected = Diagnostic(TerRules.PathBackslashInConcatenation).WithSpan(4, 23, 4, 49);
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task InterpolatedStringWithBackslash_ReportsTer403()
        {
            string test = @"
class Test {
    void Method(string folder) {
        string path = $""{folder}\\subdir"";
    }
}";
            DiagnosticResult expected = Diagnostic(TerRules.PathBackslashInConcatenation).WithSpan(4, 23, 4, 42);
            await VerifyAnalyzerAsync(test, expected);
        }

        [TestMethod]
        public async Task NormalString_NoDiagnostic()
        {
            string test = @"
class Test {
    void Method() {
        string msg = ""Hello World"";
    }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task HkeyRegistryPath_NoDiagnostic()
        {
            string test = @"
class Test {
    void Method() {
        string key = ""HKEY_LOCAL_MACHINE\\SOFTWARE\\Test"";
    }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task RegexPatternConst_NoDiagnostic()
        {
            // Repro: a verbatim regex pattern must not be misidentified as a Windows path
            // just because it contains backslash-escape sequences.
            string test = @"
class Test {
    private const string Pattern = @""^(CP|FT)\d_EVS\d_"";
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task PathWithLetterAfterBackslash_ReportsTer402()
        {
            // A real path segment that happens to start with a letter from the regex
            // word-boundary/class-escape indicator set (e.g. "\Basic") must still be flagged.
            string test = @"
class Test {
    void Method() {
        string path = ""Settings\\Basic\\Config.txt"";
    }
}";
            DiagnosticResult expected = Diagnostic(TerRules.PathBackslashInLiteral).WithSpan(4, 23, 4, 52);
            await VerifyAnalyzerAsync(test, expected);
        }
    }
}
