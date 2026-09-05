using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.Rules;
using TerLintAnalyzer.Ter.Wrapping;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class MethodWrappingAnalyzerTests : CSharpAnalyzerVerifier<WrappingAnalyzer>
    {
        [TestMethod]
        public async Task NoMethodWrapping_SingleLine_NoDiagnostic()
        {
            string test = @"
    namespace TestNamespace {
        public class MyClass {
            public void SingleLine(int a, int b) { }
        }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task NoMethodWrapping_MultiLineSignature_ReportsDiagnostic()
        {
            string test = $@"
    namespace TestNamespace {{
        public class MyClass {{
            public void {{|{TerRules.NoMethodWrappingRule.Id}:MultiLine|}}(
                int a, 
                int b) 
            {{ }}
        }}
    }}";
            // Reports NoMethodWrappingRule because the identifier to parameter list spans lines
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task NoMethodWrapping_ExpressionBodyMultiLine_ReportsDiagnostic()
        {
            string test = $@"
    namespace TestNamespace {{
        public class MyClass {{
            public int {{|{TerRules.NoMethodWrappingRule.Id}:Add|}}(int a, 
                int b) => a + b;
        }}
    }}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task GeneralNoWrapping_InvocationArgumentsSpanLines_ReportsDiagnostic()
        {
            string test = $@"
    namespace TestNamespace {{
        public class MyClass {{
            void Foo(int a, int b) {{ }}
            void Method() {{
                {{|{TerRules.GeneralNoWrappingRule.Id}:Foo|}}(1,
                    2);
            }}
        }}
    }}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task GeneralNoWrapping_InvocationSingleLine_NoDiagnostic()
        {
            string test = @"
    namespace TestNamespace {
        public class MyClass {
            void Foo(int a, int b) { }
            void Method() {
                Foo(1, 2);
            }
        }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task GeneralNoWrapping_InvocationArgumentIsNewExpressionWithInitializer_NoDiagnostic()
        {
            string test = @"
    namespace TestNamespace {
        public class Foo {
            public string Name;
        }
        public class MyClass {
            void Add(Foo foo) { }
            void Method() {
                Add(new Foo
                {
                    Name = ""Test""
                });
            }
        }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task GeneralNoWrapping_InvocationArgumentIsImplicitNewExpressionWithInitializer_NoDiagnostic()
        {
            string test = @"
    namespace TestNamespace {
        public class Foo {
            public string Name;
        }
        public class MyClass {
            void Add(Foo foo) { }
            void Method() {
                Add(new()
                {
                    Name = ""Test""
                });
            }
        }
    }";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task NoMethodWrapping_WithAttribute_NoDiagnostic()
        {
            // The attribute is on a different line, but the method identifier 
            // through the parameter list is on a single line.
            string test = @"
    namespace TestNamespace
    {
        public class MyClass
        {
            [System.Obsolete]
            public void MethodWithAttribute(int a, int b) 
            { 
            }
        }
    }";

            await VerifyAnalyzerAsync(test);
        }
    }
}
