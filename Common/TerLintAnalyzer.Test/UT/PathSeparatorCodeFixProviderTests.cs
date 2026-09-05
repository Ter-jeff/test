using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.PathSeparator;
using TerLintAnalyzer.Ter.Rules;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class PathSeparatorCodeFixProviderTests : CSharpCodeFixVerifier<PathSeparatorAnalyzer, PathSeparatorCodeFixProvider>
    {
        [TestMethod]
        public async Task BackslashPath_FixedToPathCombine()
        {
            string test = $@"
class Test {{
    void Method() {{
        string path = {{|{TerRules.PathBackslashInLiteral.Id}:""folder\\sub\\file.txt""|}};
    }}
}}";

            string fixedCode = @"using System.IO;

class Test {
    void Method() {
        string path = Path.Combine(""folder"", ""sub"", ""file.txt"");
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task DriveLetterPath_FixedToPathCombine()
        {
            string test = $@"
using System.IO;
class Test {{
    void Method() {{
        string path = {{|{TerRules.PathBackslashInLiteral.Id}:""C:\\Users\\file.txt""|}};
    }}
}}";

            string fixedCode = @"
using System.IO;
class Test {
    void Method() {
        string path = Path.Combine(""C:\\Users"", ""file.txt"");
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }
    }
}
