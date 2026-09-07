using System.Collections.Immutable;
using System.Composition;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CodeActions;
using Microsoft.CodeAnalysis.CodeFixes;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.MethodSpacing
{
    [ExportCodeFixProvider(LanguageNames.CSharp, Name = nameof(MethodSpacingCodeFixProvider)), Shared]
    public class MethodSpacingCodeFixProvider : CodeFixProvider
    {
        public sealed override ImmutableArray<string> FixableDiagnosticIds => ImmutableArray.Create(TerRules.MethodSpacingRule.Id);

        public sealed override FixAllProvider GetFixAllProvider() => WellKnownFixAllProviders.BatchFixer;

        public sealed override async Task RegisterCodeFixesAsync(CodeFixContext context)
        {
            SyntaxNode root = await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(false);
            Diagnostic diagnostic = context.Diagnostics.First();
            TextSpan diagnosticSpan = diagnostic.Location.SourceSpan;

            // Find the method declaration identified by the analyzer
            MethodDeclarationSyntax methodDeclaration = root.FindToken(diagnosticSpan.Start).Parent.AncestorsAndSelf().OfType<MethodDeclarationSyntax>().First();

            // Register the fix action that will appear in the Light Bulb
            context.RegisterCodeFix(CodeAction.Create(title: "Fix method spacing", createChangedDocument: c => FixSpacingAsync(context.Document, methodDeclaration, c), equivalenceKey: nameof(MethodSpacingCodeFixProvider)), diagnostic);
        }

        private async Task<Document> FixSpacingAsync(Document document, MethodDeclarationSyntax methodDeclaration, CancellationToken cancellationToken)
        {
            SyntaxTriviaList leadingTrivia = methodDeclaration.GetLeadingTrivia();

            // 1. Capture the indentation (the whitespace immediately before the method keyword)
            // If no whitespace exists, we use an empty string to avoid null issues
            SyntaxTrivia indentation = leadingTrivia.LastOrDefault(t => t.IsKind(SyntaxKind.WhitespaceTrivia));
            if (indentation == default)
            {
                indentation = SyntaxFactory.Whitespace("");
            }

            // 2. Filter out existing leading blank lines/spaces
            // This leaves us with just the "content" (comments, XML docs, etc.)
            var contentTrivia = leadingTrivia
                .SkipWhile(t => t.IsKind(SyntaxKind.WhitespaceTrivia) || t.IsKind(SyntaxKind.EndOfLineTrivia))
                .ToList();

            // 3. Construct the new trivia list:
            // [Blank Line] + [Indent for Comment] + [Comment] + [Newline/Indent for Method]
            SyntaxTriviaList newTrivia = SyntaxFactory.TriviaList(SyntaxFactory.EndOfLine("\r\n"));

            if (contentTrivia.Count > 0)
            {
                // Add the indent so the comment isn't at Column 0
                newTrivia = newTrivia.Add(indentation);
                newTrivia = newTrivia.AddRange(contentTrivia);
            }
            else
            {
                // If there is no comment, just add the indent for the method keyword
                newTrivia = newTrivia.Add(indentation);
            }

            // 4. Apply the new trivia to the method and update the document
            MethodDeclarationSyntax newMethod = methodDeclaration.WithLeadingTrivia(newTrivia);
            SyntaxNode root = await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(false);
            SyntaxNode newRoot = root.ReplaceNode(methodDeclaration, newMethod);

            return document.WithSyntaxRoot(newRoot);
        }
    }
}
