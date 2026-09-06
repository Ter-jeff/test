using System.Collections.Generic;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace TerLintAnalyzer.Ter.RegexInline
{
    internal static class RegexInlineCodeFixProviderHelpers
    {

        internal static FieldDeclarationSyntax CreateRegexField(string name, ExpressionSyntax pattern, IEnumerable<ArgumentSyntax> extraArgs)
        {
            var constructorArgs = new List<ArgumentSyntax> { SyntaxFactory.Argument(pattern) };
            constructorArgs.AddRange(extraArgs);

            return SyntaxFactory.FieldDeclaration(
                SyntaxFactory.VariableDeclaration(SyntaxFactory.IdentifierName("Regex"))
                .AddVariables(SyntaxFactory.VariableDeclarator(SyntaxFactory.Identifier(name))
                    .WithInitializer(SyntaxFactory.EqualsValueClause(
                        SyntaxFactory.ObjectCreationExpression(SyntaxFactory.IdentifierName("Regex"))
                        .WithArgumentList(SyntaxFactory.ArgumentList(SyntaxFactory.SeparatedList(constructorArgs)))))))
                .AddModifiers(
                    SyntaxFactory.Token(SyntaxKind.PrivateKeyword),
                    SyntaxFactory.Token(SyntaxKind.StaticKeyword),
                    SyntaxFactory.Token(SyntaxKind.ReadOnlyKeyword))
                .WithTrailingTrivia(SyntaxFactory.ElasticCarriageReturnLineFeed);
        }
    }
}
