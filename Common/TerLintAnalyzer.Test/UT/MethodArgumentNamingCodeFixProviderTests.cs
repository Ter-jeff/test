using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.MethodArgumentNaming;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class MethodArgumentNamingCodeFixProviderTests : CSharpCodeFixVerifier<MethodArgumentNamingAnalyzer, MethodArgumentNamingCodeFixProvider>
    {
        [TestMethod]
        public async Task ArgumentNaming_ValidNames_NoDiagnostic()
        {
            string test = @"
using System.Collections.Generic;
public class User {}
public interface IService {}
public class MyClass {
    public void ValidMethod(User user, IService service, List<User> users, User[] userArray) { }
}";

            // FIX: Pass the same string twice to indicate no changes are expected
            await VerifyCodeFixAsync(test, test);
        }

        [TestMethod]
        public async Task ArgumentNaming_InvalidNames_FixesNaming()
        {
            string test = @"
using System.Collections.Generic;
public class User {}
public class MyClass {
    public void InvalidMethod(User [|usr|], List<User> [|items|]) { }
}";

            string fixedCode = @"
using System.Collections.Generic;
public class User {}
public class MyClass {
    public void InvalidMethod(User user, List<User> users) { }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task ArgumentNaming_MultipleArgumentsOfSameType_Ignored()
        {
            string test = @"
public class User {}
public class MyClass {
    public void MultiMethod(User user1, User user2) { }
}";

            // FIX: Pass the same string twice to indicate no changes are expected
            await VerifyCodeFixAsync(test, test);
        }
    }
}
