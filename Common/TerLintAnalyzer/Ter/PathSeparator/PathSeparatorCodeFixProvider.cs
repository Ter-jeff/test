using System.Collections.Immutable;
using System.Composition;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CodeActions;
using Microsoft.CodeAnalysis.CodeFixes;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.PathSeparator
{
    [ExportCodeFixProvider(LanguageNames.CSharp, Name = nameof(PathSeparatorCodeFixProvider)), Shared]
    public class PathSeparatorCodeFixProvider : CodeFixProvider
    {
        public sealed override ImmutableArray<string> FixableDiagnosticIds => ImmutableArray.Create(TerRules.PathBackslashInLiteral.Id);

        public sealed override FixAllProvider GetFixAllProvider() => WellKnownFixAllProviders.BatchFixer;

        public sealed override async Task RegisterCodeFixesAsync(CodeFixContext context)
        {
            SyntaxNode root = await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(false);
            Diagnostic diagnostic = context.Diagnostics.First();
            Microsoft.CodeAnalysis.Text.TextSpan span = diagnostic.Location.SourceSpan;

            SyntaxNode node = root?.FindNode(span);
            if (!(node is LiteralExpressionSyntax literal))
            {
                return;
            }

            string value = literal.Token.ValueText;
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            string[] segments = Regex.Split(value, @"(?<!:)\\");
            if (segments.Length < 2 || segments.All(string.IsNullOrEmpty))
            {
                return;
            }

            context.RegisterCodeFix(CodeAction.Create(title: "Use Path.Combine()", createChangedDocument: ct => ReplaceWithPathCombine(context.Document, root, literal, segments), equivalenceKey: "UsePathCombine"), diagnostic);
        }

        private static Task<Document> ReplaceWithPathCombine(
            Document document, SyntaxNode root, LiteralExpressionSyntax literal,
            string[] segments)
        {
            string[] nonEmpty = segments.Where(s => !string.IsNullOrEmpty(s)).ToArray();
            if (nonEmpty.Length == 0)
            {
                return Task.FromResult(document);
            }

            SeparatedSyntaxList<ArgumentSyntax> arguments = SyntaxFactory.SeparatedList(
                nonEmpty.Select(s =>
                    SyntaxFactory.Argument(
                        SyntaxFactory.LiteralExpression(
                            SyntaxKind.StringLiteralExpression,
                            SyntaxFactory.Literal(s)))));

            InvocationExpressionSyntax pathCombine = SyntaxFactory.InvocationExpression(
                SyntaxFactory.MemberAccessExpression(
                    SyntaxKind.SimpleMemberAccessExpression,
                    SyntaxFactory.IdentifierName("Path"),
                    SyntaxFactory.IdentifierName("Combine")),
                SyntaxFactory.ArgumentList(arguments));

            SyntaxNode newRoot = root.ReplaceNode(literal, pathCombine.WithTriviaFrom(literal));

            if (newRoot is CompilationUnitSyntax compilationUnit)
            {
                bool hasSystemIo = compilationUnit.Usings.Any(u => u.Name?.ToString() == "System.IO");
                if (!hasSystemIo)
                {
                    UsingDirectiveSyntax usingDirective = SyntaxFactory.UsingDirective(
                        SyntaxFactory.QualifiedName(
                            SyntaxFactory.IdentifierName("System"),
                            SyntaxFactory.IdentifierName("IO")))
                        .WithTrailingTrivia(SyntaxFactory.CarriageReturnLineFeed);

                    newRoot = compilationUnit.AddUsings(usingDirective);
                }
            }

            return Task.FromResult(document.WithSyntaxRoot(newRoot));
        }
    }
}
