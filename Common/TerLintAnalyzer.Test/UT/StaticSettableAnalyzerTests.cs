using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.StaticSettable;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class StaticSettableAnalyzerTests : CSharpAnalyzerVerifier<StaticSettableAnalyzer>
    {
        [TestMethod]
        public async Task Warning_When_FileNoContainsStatic_And_MemberIsSettable()
        {
            const string testCode = @"
            public class Config {
                public static int [|MyProperty|] { get; set; }
            }";

            var test = new Test
            {
                TestState =
                {
                    Sources = { (filename: "app_static_config.cs", content: testCode) },
                }
            };

            await test.RunAsync();
        }

        [TestMethod]
        public async Task Warning_When_FileNoContainsStatic_And_MemberIsSettable_Filed()
        {
            const string testCode = @"
            public class Config {
                public static int [|MyProperty|];
            }";

            var test = new Test
            {
                TestState =
                {
                    Sources = { (filename: "app_static_config.cs", content: testCode) },
                }
            };

            await test.RunAsync();
        }

        [TestMethod]
        public async Task NoWarning_When_FilePath_ContainStatic()
        {
            const string testCode = @"
            public class User {
                public static string SharedSecret;
            }";

            var test = new Test
            {
                TestState =
                {
                    Sources = { (filename: "/Static/UserModule.cs", content: testCode) },
                }
            };

            await test.RunAsync();
        }

        [TestMethod]
        public async Task NoWarning_When_StaticMember_IsReadonly()
        {
            const string testCode = @"
            public class Settings {
                public static readonly int MaxRetries = 5;
            }";

            var test = new Test
            {
                TestState =
                {
                    Sources = { (filename: "app_static_config.cs", content: testCode) },
                }
            };

            await test.RunAsync();
        }
    }
}
