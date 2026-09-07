using System.Collections.Immutable;
using System.Composition;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CodeActions;
using Microsoft.CodeAnalysis.CodeFixes;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.Wrapping
{
    [ExportCodeFixProvider(LanguageNames.CSharp, Name = nameof(WrappingCodeFixProvider)), Shared]
    public class WrappingCodeFixProvider : CodeFixProvider
    {
        public sealed override ImmutableArray<string> FixableDiagnosticIds => ImmutableArray.Create(TerRules.GeneralNoWrappingRule.Id);

        public sealed override FixAllProvider GetFixAllProvider() => WellKnownFixAllProviders.BatchFixer;

        public sealed override async Task RegisterCodeFixesAsync(CodeFixContext context)
        {
            SyntaxNode root = await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(false);
            Diagnostic diagnostic = context.Diagnostics.First();
            SyntaxNode node = root.FindNode(diagnostic.Location.SourceSpan, getInnermostNodeForTie: true);

            // GeneralNoWrappingRule also covers property/class/namespace declarations, but only
            // the InvocationExpression case can actually fire (the others compare a location to
            // itself). Only offer a fix when we can locate the wrapped invocation.
            InvocationExpressionSyntax invocation = node.FirstAncestorOrSelf<InvocationExpressionSyntax>();
            if (invocation == null || ContainsComment(invocation))
            {
                return;
            }

            context.RegisterCodeFix(CodeAction.Create(title: "Collapse call onto a single line", createChangedDocument: c => CollapseToSingleLineAsync(context.Document, invocation, c), equivalenceKey: nameof(WrappingCodeFixProvider)), diagnostic);
        }

        // Bail out rather than risk swallowing a comment (e.g. a trailing "//" note on an
        // argument) into the middle of the collapsed line.
        private static bool ContainsComment(SyntaxNode node)
        {
            return node.DescendantTrivia().Any(t => t.IsKind(SyntaxKind.SingleLineCommentTrivia) || t.IsKind(SyntaxKind.MultiLineCommentTrivia));
        }

        private async Task<Document> CollapseToSingleLineAsync(Document document, InvocationExpressionSyntax invocation, CancellationToken cancellationToken)
        {
            // A line break's indentation is split across the trailing trivia of the token
            // before it and the leading trivia of the token after it, so collapsing each
            // token's trivia independently leaves stray whitespace behind. Operating on the
            // node's own text (which excludes the invocation's own outer leading/trailing
            // trivia) sidesteps that boundary problem entirely.
            string collapsedText = Regex.Replace(invocation.ToString(), @"[ \t]*\r?\n\s*", " ");
            // The collapse above can leave a stray space where a line break used to sit right
            // after an opening paren/bracket or right before a closing one (e.g. "Foo(\n  x)" ->
            // "Foo( x)"); strip it so the result matches the codebase's no-space-inside-parens style.
            collapsedText = Regex.Replace(collapsedText, @"(?<=[(\[]) ", "");
            collapsedText = Regex.Replace(collapsedText, @" (?=[)\]])", "");
            ExpressionSyntax collapsed = SyntaxFactory.ParseExpression(collapsedText).WithTriviaFrom(invocation);

            SyntaxNode root = await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(false);
            SyntaxNode newRoot = root.ReplaceNode(invocation, collapsed);

            return document.WithSyntaxRoot(newRoot);
        }
    }
}
