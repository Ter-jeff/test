using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.MethodArgumentNaming;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class MethodArgumentNamingAnalyzerTests : CSharpAnalyzerVerifier<MethodArgumentNamingAnalyzer>
    {
        [TestMethod]
        public async Task ArgumentNaming_ValidNames_NoDiagnostic()
        {
            string test = @"
    using System.Collections.Generic;
    namespace TestNamespace {
        public class User {}
        public interface IService {}
        public class MyClass {
            public void ValidMethod(User user, IService service, List<User> users, User[] userArray) { }
        }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task ArgumentNaming_InvalidNames_ReportsDiagnostics()
        {
            string test = @"
    using System.Collections.Generic;
    namespace TestNamespace {
        public class User {}
        public class MyClass {
            public void InvalidMethod(User [|usr|], List<User> [|items|]) { }
        }
    }";
            // This expects ArgumentNamingRule to fire on 'usr' (expected 'user') 
            // and 'items' (expected 'users')
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task ArgumentNaming_MultipleArgumentsOfSameType_Ignored()
        {
            string test = @"
    namespace TestNamespace {
        public class User {}
        public class MyClass {
            public void MultiMethod(User user1, User user2) { }
        }
    }";
            // Should not report because Group.Count() > 1
            await VerifyAnalyzerAsync(test);
        }
    }
}
