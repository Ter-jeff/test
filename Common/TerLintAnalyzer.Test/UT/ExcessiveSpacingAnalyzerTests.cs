using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.ExcessiveSpacing;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class ExcessiveSpacingAnalyzerTests : CSharpAnalyzerVerifier<ExcessiveSpacingAnalyzer>
    {
        [TestMethod]
        public async Task SingleBlankLine_NoDiagnostic()
        {
            // 1 empty line = 2 newlines. No diagnostic expected.
            string test = @"class Test
{
    void M1() { }

    void M2() { }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task TwoBlankLines_ReportsDiagnostic()
        {
            string test = @"class Test
{
    void M1() { }
    
[|
|]    
    void M2() { }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task ManyBlankLines_ReportsDiagnostic()
        {
            string test = @"class Test
{

[|
|]
     void M1() { }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task TwoBlankLinesInsideMethod_ReportsDiagnostic()
        {
            string test = @"class Test
{
    void M()
    {
        int x = 1;

[|
|]
    }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task SpacesOnEmptyLines_StillReportsDiagnostic()
        {
            string test = @"class Test
{
    void M1() { }

[|
|]    void M2() { }
}";
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task SpacesOnRegion_StillReportsDiagnostic()
        {
            string test = @"class Test
{
    void M1() { }

    void M2() 
    {
        #region data1

        #endregion

        #region data2
        #endregion
    }
}";
            await VerifyAnalyzerAsync(test);
        }
    }
}
