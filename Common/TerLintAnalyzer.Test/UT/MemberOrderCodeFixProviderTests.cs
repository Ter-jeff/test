using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.MemberOrder;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class MemberOrderCodeFixProviderTests : CSharpCodeFixVerifier<MemberOrderAnalyzer, MemberOrderCodeFixProvider>
    {
        [TestMethod]
        public async Task PropertyAfterMethod_MovedBeforeMethod()
        {
            const string test = @"
class MyClass
{
    public void MyMethod() { }
    [|public string MyProperty { get; set; }|]
}";

            const string fixedCode = @"
class MyClass
{
    public string MyProperty { get; set; }
    public void MyMethod() { }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task FieldAfterMethod_MovedBeforeMethod()
        {
            const string test = @"
class MyClass
{
    public void MyMethod() { }
    [|public int MyField;|]
}";

            const string fixedCode = @"
class MyClass
{
    public int MyField;
    public void MyMethod() { }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task CtorAfterMethod_MovedBeforeMethod()
        {
            const string test = @"
class MyClass
{
    public void MyMethod() { }
    [|public MyClass() { }|]
}";

            const string fixedCode = @"
class MyClass
{
    public MyClass() { }
    public void MyMethod() { }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }
    }
}
