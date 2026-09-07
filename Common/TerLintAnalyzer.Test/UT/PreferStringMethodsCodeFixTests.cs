using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.PreferStringMethods;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class PreferStringMethodsCodeFixTests : CSharpCodeFixVerifier<PreferStringMethodsAnalyzer, PreferStringMethodsCodeFixProvider>
    {
        private const string ExtensionSource = @"
namespace CommonLib.Extension {
    public static class StringExtensions {
        public static bool ContainsIgnoreCase(this string s, string value) => true;
        public static bool EndsWithIgnoreCase(this string s, string value) => true;
        public static bool StartsWithIgnoreCase(this string s, string value) => true;
        public static bool EqualsIgnoreCase(this string s, string value) => true;
    }
}";

        [TestMethod]
        public async Task Test_RegexEquals_FixedToStringEquals()
        {
            const string test = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string input) {
        var rx = new Regex([|""^exact$""|]);
        bool result = rx.IsMatch(input);
    }
}";

            const string fixedCode = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string input) {
        
        bool result = input.Equals(""exact"");
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_RegexStartsWith_FixedToStringStartsWith()
        {
            const string test = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string input) {
        var rx = new Regex([|""^abc""|]);
        bool result = rx.IsMatch(input);
    }
}";

            const string fixedCode = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string input) {
        
        bool result = input.StartsWith(""abc"");
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_RegexEndsWith_FixedToStringEndsWith()
        {
            const string test = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string input) {
        var rx = new Regex([|""suffix$""|]);
        bool result = rx.IsMatch(input);
    }
}";

            const string fixedCode = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string input) {
        
        bool result = input.EndsWith(""suffix"");
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_RegexContains_FixedToStringContains()
        {
            const string test = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string input) {
        var rx = new Regex([|""middle""|]);
        bool result = rx.IsMatch(input);
    }
}";

            const string fixedCode = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string input) {
        
        bool result = input.Contains(""middle"");
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_RegexEqualsIgnoreCase_FixedToStringEquals()
        {
            const string test = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string input) {
        var rx = new Regex([|""^exact$""|], RegexOptions.IgnoreCase);
        bool result = rx.IsMatch(input);
    }
}" + ExtensionSource;

            const string fixedCode = @"
using System.Text.RegularExpressions;
using CommonLib.Extension;
class Test {
    void Method(string input) {
        
        bool result = input.EqualsIgnoreCase(""exact"");
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_RegexStartsWithIgnoreCase_FixedToStringStartsWith()
        {
            const string test = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string input) {
        var rx = new Regex([|""^abc""|], RegexOptions.IgnoreCase);
        bool result = rx.IsMatch(input);
    }
}" + ExtensionSource;

            const string fixedCode = @"
using System.Text.RegularExpressions;
using CommonLib.Extension;
class Test {
    void Method(string input) {
        
        bool result = input.StartsWithIgnoreCase(""abc"");
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_RegexEndsWithIgnoreCase_FixedToStringEndsWith()
        {
            const string test = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string input) {
        var rx = new Regex([|""suffix$""|], RegexOptions.IgnoreCase);
        bool result = rx.IsMatch(input);
    }
}" + ExtensionSource;

            const string fixedCode = @"
using System.Text.RegularExpressions;
using CommonLib.Extension;
class Test {
    void Method(string input) {
        
        bool result = input.EndsWithIgnoreCase(""suffix"");
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_RegexContainsIgnoreCase_FixedToStringContains()
        {
            const string test = @"
using System.Text.RegularExpressions;
class Test {
    void Method(string input) {
        var rx = new Regex([|""middle""|], RegexOptions.IgnoreCase);
        bool result = rx.IsMatch(input);
    }
}" + ExtensionSource;

            const string fixedCode = @"
using System.Text.RegularExpressions;
using CommonLib.Extension;
class Test {
    void Method(string input) {
        
        bool result = input.ContainsIgnoreCase(""middle"");
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_RegexContainsIgnoreCase_FixedToStringContains_Field()
        {
            const string test = @"
using System.Text.RegularExpressions;
class Test {
    private static readonly Regex _regex = new Regex([|""BI""|], RegexOptions.IgnoreCase);
    void Method(string input) {
        bool result = _regex.IsMatch(input);
    }
}" + ExtensionSource;

            const string fixedCode = @"
using System.Text.RegularExpressions;
using CommonLib.Extension;
class Test {
    
    void Method(string input) {
        bool result = input.ContainsIgnoreCase(""BI"");
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_RegexContainsIgnoreCase_FixedToStringContains_Field_2()
        {
            const string test = @"
using System.Text.RegularExpressions;
class Test {
    private static readonly Regex _regex = new Regex([|""BI""|], RegexOptions.IgnoreCase | RegexOptions.Compiled);
    void Method(string input) {
        bool result = _regex.IsMatch(input);
    }
}" + ExtensionSource;

            const string fixedCode = @"
using System.Text.RegularExpressions;
using CommonLib.Extension;
class Test {
    
    void Method(string input) {
        bool result = input.ContainsIgnoreCase(""BI"");
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }
    }
}
