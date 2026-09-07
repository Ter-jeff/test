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

namespace TerLintAnalyzer
{
    [ExportCodeFixProvider(LanguageNames.CSharp, Name = nameof(CommentCodeFixProvider)), Shared]
    public class CommentCodeFixProvider : CodeFixProvider
    {
        public sealed override ImmutableArray<string> FixableDiagnosticIds => ImmutableArray.Create(CommentRules.TrailingComment.Id);

        public sealed override FixAllProvider GetFixAllProvider() => WellKnownFixAllProviders.BatchFixer;

        public sealed override async Task RegisterCodeFixesAsync(CodeFixContext context)
        {
            SyntaxNode root = await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(false);
            Diagnostic diagnostic = context.Diagnostics.First();
            TextSpan diagnosticSpan = diagnostic.Location.SourceSpan;

            context.RegisterCodeFix(CodeAction.Create(title: "Move trailing comment", createChangedDocument: c => MoveTrailingCommentAsync(context.Document, root.FindTrivia(diagnosticSpan.Start), c), equivalenceKey: nameof(CommentCodeFixProvider)), diagnostic);
        }

        private async Task<Document> MoveTrailingCommentAsync(Document document, SyntaxTrivia trivia, CancellationToken cancellationToken)
        {
            SyntaxNode root = await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(false);
            SyntaxToken token = trivia.Token;

            // 1. Find the target "line" node (Statement or Field/Property)
            SyntaxNode targetNode = token.Parent;
            while (targetNode != null &&
                   !(targetNode is StatementSyntax) &&
                   !(targetNode is MemberDeclarationSyntax))
            {
                targetNode = targetNode.Parent;
            }

            if (targetNode == null)
            {
                return document;
            }

            // 2. The existing leading trivia may include blank-line spacing carried over
            // from the previous member (e.g. a blank line separating fields). The final
            // whitespace trivia in that list is the indentation immediately before the
            // target node's own keyword; the moved comment must be inserted right before
            // that indentation, not before any earlier blank-line spacing, or it ends up
            // visually attached to the previous member instead of this one.
            SyntaxTriviaList leadingTrivia = targetNode.GetLeadingTrivia();
            int lastIndentIndex = -1;
            for (int i = leadingTrivia.Count - 1; i >= 0; i--)
            {
                if (leadingTrivia[i].IsKind(SyntaxKind.WhitespaceTrivia))
                {
                    lastIndentIndex = i;
                    break;
                }
            }

            SyntaxTrivia indentation = lastIndentIndex >= 0 ? leadingTrivia[lastIndentIndex] : SyntaxFactory.Whitespace("");
            SyntaxTriviaList precedingTrivia = lastIndentIndex >= 0
                ? SyntaxFactory.TriviaList(leadingTrivia.Take(lastIndentIndex))
                : leadingTrivia;

            // 3. Clean the semicolon (remove comment and its preceding space)
            SyntaxTriviaList trailing = token.TrailingTrivia;
            int index = trailing.IndexOf(trivia);
            SyntaxTriviaList cleanedTrailing = trailing.Remove(trivia);

            if (index > 0 && trailing[index - 1].IsKind(SyntaxKind.WhitespaceTrivia))
            {
                cleanedTrailing = cleanedTrailing.Remove(trailing[index - 1]);
            }

            // 4. Build the new leading trivia: [preserved earlier spacing] + [indent] +
            // [comment] + [newline] + [indent] (for the target node itself)
            SyntaxTriviaList newLeadingTrivia = precedingTrivia
                .Add(indentation)
                .Add(trivia)
                .Add(SyntaxFactory.CarriageReturnLineFeed)
                .Add(indentation);

            // 5. Apply changes: Clean the token and prepend the new trivia block
            SyntaxNode updatedNode = targetNode
                .ReplaceToken(token, token.WithTrailingTrivia(cleanedTrailing))
                .WithLeadingTrivia(newLeadingTrivia);

            return document.WithSyntaxRoot(root.ReplaceNode(targetNode, updatedNode));
        }
    }
}
