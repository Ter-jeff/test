using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.EnumNaming;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class EnumNamingAnalyzerTests : CSharpAnalyzerVerifier<EnumNamingAnalyzer>
    {
        [TestMethod]
        public async Task EnumWithoutPrefix_ReportsDiagnostic()
        {
            // The [| |] markup indicates where we expect the diagnostic to be reported
            string test = @"
    namespace MyNamespace
    {
        public enum [|Status|]
        {
            Active,
            Inactive
        }
    }";

            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task EnumWithPrefix_NoDiagnostic()
        {
            // Valid prefix, should not report anything
            string test = @"
    namespace MyNamespace
    {
        public enum EnumStatus
        {
            Active
        }
    }";

            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task ClassWithNonMatchingName_NoDiagnostic()
        {
            // Rule only applies to Enums, so classes should be ignored
            string test = @"
    namespace MyNamespace
    {
        public class Status { }
    }";

            await VerifyAnalyzerAsync(test);
        }
    }
}
