using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class CommentCodeFixProviderTests : CSharpCodeFixVerifier<CommentAnalyzer, CommentCodeFixProvider>
    {
        [TestMethod]
        public async Task Test_MoveTrailingCommentToLineAbove()
        {
            // Use the [| |] syntax to mark the diagnostic span
            string test = @"
class Test
{
    void Method()
    {
        int x = 1;[|// This is a trailing comment|]
    }
}";

            // Ensure the indentation of the moved comment matches the statement below it
            string fixedCode = @"
class Test
{
    void Method()
    {
        // This is a trailing comment
        int x = 1;
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_MoveTrailingCommentWithNoLeadingTrivia()
        {
            // Specifically testing the case where the comment is directly against the semicolon
            string test = @"
class Test {
    void Method() {
        string s = ""test"";[|//move me|]
    }
}";

            string fixedCode = @"
class Test {
    void Method() {
        //move me
        string s = ""test"";
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_MoveTrailingComment_MethodCall()
        {
            // Testing that it works for expression statements, not just declarations
            string test = @"
using System;
class Test {
    void Method() {
        Console.WriteLine(""Hello"");[|//Log message|]
    }
}";

            string fixedCode = @"
using System;
class Test {
    void Method() {
        //Log message
        Console.WriteLine(""Hello"");
    }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_MoveTrailingComment_Field()
        {
            // Use 'Empty' or escape the quotes to avoid CS1010/CS1002
            string test = @"
using System;
class Test {
    public int x1 = 1;[|// This is a trailing comment|]
    public string Org { get; set; } = string.Empty; [|//If not exist TW will added|]
}";

            string fixedCode = @"
using System;
class Test {
    // This is a trailing comment
    public int x1 = 1;
    //If not exist TW will added
    public string Org { get; set; } = string.Empty; 
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Test_MoveTrailingComment_PreservesPrecedingBlankLine()
        {
            // Repro: a blank line separating this field from the previous one must stay
            // attached to the previous field; the moved comment must land directly above
            // its own field, not above the blank line (which would visually attach it to
            // the previous field instead).
            string test = @"
class Test {
    public int BinnedIdx = -1;

    public int StartRowIdx = -1;[|//Doamin|]
    public int DomainIdx = -1;
}";

            string fixedCode = @"
class Test {
    public int BinnedIdx = -1;

    //Doamin
    public int StartRowIdx = -1;
    public int DomainIdx = -1;
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }
    }
}
