using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class StringCodeFixProviderTests : CSharpCodeFixVerifier<StringAnalyzer, StringCodeFixProvider>
    {
        private const string ExtensionSource = @"
#pragma warning disable
namespace CommonLib.Extension {
    public static class StringExtensions {
        public static bool EqualsIgnoreCase(this string s, string value) => true;
        public static bool ContainsIgnoreCase(this string s, string value) => true;
        public static bool StartsWithIgnoreCase(this string s, string value) => true;
        public static bool EndsWithIgnoreCase(this string s, string value) => true;
        public static int IndexOfIgnoreCase(this string s, string value) => -1;
        public static StringComparer IgnoreCase => StringComparer.OrdinalIgnoreCase;
    }
}";

        [TestMethod]
        public async Task Test_StaticEquals_FixedToEqualsIgnoreCase()
        {
            string test = $@"
using System;
class Test {{
    void Method(string s1, string s2) {{
        bool r = {{|{StringRules.StringEqualsCaseInsensitive.Id}:string.Equals(s1, s2, StringComparison.OrdinalIgnoreCase)|}};
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
using CommonLib.Extension;
class Test {
    void Method(string s1, string s2) {
        bool r = s1.EqualsIgnoreCase(s2);
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_ToUpperContains_FixedToContainsIgnoreCase()
        {
            string test = $@"
using System;
class Test {{
    void Method(string s) {{
        bool r = {{|{StringRules.ToUpperToLowerComparison.Id}:s.ToUpper()|}}.Contains(""abc"");
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
using CommonLib.Extension;
class Test {
    void Method(string s) {
        bool r = s.ContainsIgnoreCase(""abc"");
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_EqualityOperator_FixedToEqualsIgnoreCase()
        {
            string test = $@"
using System;
class Test {{
    void Method(string s1, string s2) {{
        bool r = {{|{StringRules.ToUpperToLowerEqualityOperator.Id}:s1.ToLower() == s2|}};
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
using CommonLib.Extension;
class Test {
    void Method(string s1, string s2) {
        bool r = s1.EqualsIgnoreCase(s2);
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_NotEqualsOperator_FixedToNegatedEqualsIgnoreCase()
        {
            string test = $@"
using System;
class Test {{
    void Method(string s1, string s2) {{
        bool r = {{|{StringRules.ToUpperToLowerEqualityOperator.Id}:s1.ToUpper() != s2|}};
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
using CommonLib.Extension;
class Test {
    void Method(string s1, string s2) {
        bool r = !s1.EqualsIgnoreCase(s2);
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_StringComparer_FixedToIgnoreCase()
        {
            string test = $@"
using System;
class Test {{
    void Method() {{
        var c = {{|{StringRules.StringComparerUsage.Id}:StringComparer.OrdinalIgnoreCase|}};
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
using CommonLib.Extension;
class Test {
    void Method() {
        var c = StringExtensions.IgnoreCase;
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_StringComparer_FixedToIgnoreCase_2()
        {
            // Added 'using System.Collections.Generic;'
            string test = $@"
using System;
using System.Collections.Generic;
class Test {{
    void Method() {{
        Dictionary<string, string> pinNameList = new Dictionary<string, string>({{|{StringRules.StringComparerUsage.Id}:StringComparer.OrdinalIgnoreCase|}});
    }}
}}" + ExtensionSource;

            // Ensure fixed code also includes the using and matches the structure
            string fixedCode = @"
using System;
using System.Collections.Generic;
using CommonLib.Extension;
class Test {
    void Method() {
        Dictionary<string, string> pinNameList = new Dictionary<string, string>(StringExtensions.IgnoreCase);
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_StaticEqualsCaseSensitive_FixedToEqualityOperator()
        {
            // Rule: StringEqualsCaseSensitive (IGLint301)
            // Goal: string.Equals(a, b) -> a == b
            string test = $@"
using System;
class Test {{
    void Method(string s1, string s2) {{
        bool r = {{|{StringRules.StringEqualsCaseSensitive.Id}:string.Equals(s1, s2)|}};
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
class Test {
    void Method(string s1, string s2) {
        bool r = s1 == s2;
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_StaticEqualsCaseSensitive_FixedToEqualityOperator_not()
        {
            // Rule: StringEqualsCaseSensitive (IGLint301)
            // Goal: string.Equals(a, b) -> a == b
            string test = $@"
using System;
class Test {{
    void Method(string s1, string s2) {{
        bool r = !{{|{StringRules.StringEqualsCaseSensitive.Id}:string.Equals(s1, s2)|}};
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
class Test {
    void Method(string s1, string s2) {
        bool r = s1 != s2;
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_InstanceEqualsCaseSensitive_FixedToEqualityOperator()
        {
            // Rule: EqualsCaseSensitive (IGLint303)
            // Goal: s1.Equals(s2) -> s1 == s2
            string test = $@"
using System;
class Test {{
    void Method(string s1, string s2) {{
        bool r = {{|{StringRules.EqualsCaseSensitive.Id}:s1.Equals(s2)|}};
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
class Test {
    void Method(string s1, string s2) {
        bool r = s1 == s2;
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_InstanceEqualsCaseInsensitive_FixedToEqualsIgnoreCase()
        {
            // Rule: EqualsCaseInsensitive (IGLint304)
            // Goal: s1.Equals(s2, StringComparison.OrdinalIgnoreCase) -> s1.EqualsIgnoreCase(s2)
            string test = $@"
using System;
class Test {{
    void Method(string s1, string s2) {{
        bool r = {{|{StringRules.EqualsCaseInsensitive.Id}:s1.Equals(s2, StringComparison.OrdinalIgnoreCase)|}};
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
using CommonLib.Extension;
class Test {
    void Method(string s1, string s2) {
        bool r = s1.EqualsIgnoreCase(s2);
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_InstanceEqualsCaseInsensitive_FixedToStartsWithIgnoreCase()
        {
            // Rule: EqualsCaseInsensitive (IGLint304)
            // Goal: s1.Equals(s2, StringComparison.OrdinalIgnoreCase) -> s1.EqualsIgnoreCase(s2)
            string test = $@"
using System;
class Test {{
    void Method(string s1, string s2) {{
        bool r = {{|{StringRules.EqualsCaseInsensitive.Id}:s1.StartsWith(s2, StringComparison.OrdinalIgnoreCase)|}};
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
using CommonLib.Extension;
class Test {
    void Method(string s1, string s2) {
        bool r = s1.StartsWithIgnoreCase(s2);
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_InstanceIndexCaseInsensitive_FixedToEqualsIgnoreCase()
        {
            string test = $@"
using System;
class Test {{
    void Method(string s1, string s2) {{
        if ({{|{StringRules.ToUpperToLowerComparison.Id}:""srcFile"".ToLower()|}}.IndexOf("".txt"", StringComparison.OrdinalIgnoreCase) != -1)
        {{
        }}
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
using CommonLib.Extension;
class Test {
    void Method(string s1, string s2) {
        if (""srcFile"".IndexOfIgnoreCase("".txt"") != -1)
        {
        }
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_InstanceEqualsCaseSensitive_NestedAsSecondArgument_FixedToEqualityOperator()
        {
            // Repro: s2.Equals(s3) nested as a non-first argument inside another call
            // must only replace the inner Equals call, not the outer invocation's arguments.
            string test = $@"
using System;
using System.Collections.Generic;
class Test {{
    void Method(Dictionary<string, bool> dic, string s1, string s2, string s3) {{
        dic.Add(s1, {{|{StringRules.EqualsCaseSensitive.Id}:s2.Equals(s3)|}});
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
using System.Collections.Generic;
class Test {
    void Method(Dictionary<string, bool> dic, string s1, string s2, string s3) {
        dic.Add(s1, s2 == s3);
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_InstanceIndexCaseInsensitive_FixedToEqualsIgnoreCase_Linq()
        {
            string test = $@"
using System;
using System.Collections.Generic;
class Test {{
    void Method(string s1, string s2) {{
        int rowind = new List<string>().FindIndex(row => {{|{StringRules.ToUpperToLowerEqualityOperator.Id}:row.ToLower() == ""call""|}} && row == ""call2"");
    }}
}}" + ExtensionSource;

            string fixedCode = @"
using System;
using System.Collections.Generic;
using CommonLib.Extension;
class Test {
    void Method(string s1, string s2) {
        int rowind = new List<string>().FindIndex(row => row.EqualsIgnoreCase(""call"") && row == ""call2"");
    }
}" + ExtensionSource;

            await VerifyCodeFixAsync(test, fixedCode);
        }
    }
}
