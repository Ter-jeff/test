using System.Linq;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace TerLintAnalyzer.Utility
{
    internal static class Using
    {
        internal static SyntaxNode AddUsingDirective(SyntaxNode root, string name)
        {
            if (!(root is CompilationUnitSyntax compilationUnit))
            {
                return root;
            }

            if (compilationUnit.Usings.Any(u => u.Name.ToString() == name))
            {
                return root;
            }

            UsingDirectiveSyntax newUsing = SyntaxFactory.UsingDirective(SyntaxFactory.ParseName(name)).NormalizeWhitespace().WithTrailingTrivia(SyntaxFactory.CarriageReturnLineFeed);
            return compilationUnit.AddUsings(newUsing);
        }
    }
}
