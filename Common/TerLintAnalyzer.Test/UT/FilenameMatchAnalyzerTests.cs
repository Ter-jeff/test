using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.FilenameMatch;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class FilenameMatchAnalyzerTests : CSharpAnalyzerVerifier<FilenameMatchAnalyzer>
    {
        [TestMethod]
        public async Task ClassNameMatchesFileName_NoDiagnostic()
        {
            // The helper allows us to specify the filename for the test
            string testCode = @"
    namespace MyNamespace
    {
        class MyClass { }
    }";
            var test = new Test
            {
                TestState =
                {
                    Sources = { (filename: "MyClass.cs", content: testCode) },
                }
            };

            await test.RunAsync();
        }

        [TestMethod]
        public async Task ClassNameDoesNotMatchFileName_ReportsDiagnostic()
        {
            // File is named "WrongName.cs", but class is "MyClass"
            string testCode = @"
    namespace MyNamespace
    {
        class [|MyClass|] { }
    }";
            var test = new Test
            {
                TestState =
                {
                    Sources = { (filename: "WrongName.cs", content: testCode) },
                }
            };

            // We pass the expected diagnostic arguments (className, fileName)
            // to the verifier if your Rule has placeholders like {0} and {1}
            await test.RunAsync();
        }

        [TestMethod]
        public async Task SecondClassInFile_NoDiagnostic()
        {
            // Your logic specifically skips classes that aren't the first member
            string testCode = @"
        class MyClass { }
        class SecondaryClass { }";

            var test = new Test
            {
                TestState =
                {
                    Sources = { (filename: "MyClass.cs", content: testCode) },
                }
            };

            await test.RunAsync();
        }

        [TestMethod]
        public async Task NestedClass_NoDiagnostic()
        {
            string testCode = @"
        class OuterClass
        {
            class InnerClass { }
        }";

            var test = new Test
            {
                TestState =
                {
                    Sources = { (filename: "OuterClass.cs", content: testCode) },
                }
            };

            await test.RunAsync();
        }
    }
}
