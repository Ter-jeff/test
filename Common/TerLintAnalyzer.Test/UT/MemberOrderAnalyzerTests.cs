using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.MemberOrder;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class MemberOrderAnalyzerTests : CSharpAnalyzerVerifier<MemberOrderAnalyzer>
    {
        [TestMethod]
        public async Task CorrectOrder_NoDiagnostic()
        {
            const string test = @"
    using System.Text.RegularExpressions;
    class MyClass
    {
        public const string MyConst = ""Test"";
        private static readonly Regex MyRegex = new Regex(@""\d"");
        public int MyField;
        public string MyProperty { get; set; }
        public MyClass() { }
        public void MyMethod() { }
    }";

            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task OutOfOrder_MethodBeforeCtor_ReportsDiagnostic()
        {
            // The [| ... |] syntax indicates where the diagnostic (warning) is expected.
            const string test = @"
    class MyClass
    {
        public void MyMethod() { }
        [|public MyClass() { }|]
    }";

            // This verifies that TER002 is reported on the Constructor
            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task OutOfOrder_FieldBeforeRegex_ReportsDiagnostic()
        {
            const string test = @"
    using System.Text.RegularExpressions;
    class MyClass
    {
        public int MyField;
        [|private static readonly Regex MyRegex = null;|]
    }";

            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task Interface_CorrectOrder_NoDiagnostic()
        {
            const string test = @"
    using System.Collections.Generic;
    public interface IProdCharItem { 
        string PayloadValue { get; } 
        int RowNum { set; get; }       // Property (Weight 3)
        List<string> GetInitList();    // Method (Weight 5)
    }";

            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task Interface_IncorrectOrder_ReportsDiagnostic()
        {
            // This simulates the RowNum failure you saw
            const string test = @"
    using System.Collections.Generic;
    public interface IProdCharItem { 
        List<string> GetInitList();    // Method (Weight 5)
        [|int RowNum { set; get; }|]    // Property (Weight 3) - Should FAIL here
    }";

            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task ClassInClass_NoDiagnostic()
        {
            const string test = @"
    using System.Text.RegularExpressions;
    class MyClass
    {
        class MyClass1
        {
        }
        public int MyField;
    }";

            await VerifyAnalyzerAsync(test);
        }

        [TestMethod]
        public async Task PropertyChanged_MethodBeforeEvent_ReportsDiagnostic()
        {
            const string test = @"
using System.ComponentModel;
using System.Runtime.CompilerServices;
class ViewModel : INotifyPropertyChanged
{
    protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
    [|public event PropertyChangedEventHandler PropertyChanged;|]
}";
            await VerifyAnalyzerAsync(test);
        }
    }
}
