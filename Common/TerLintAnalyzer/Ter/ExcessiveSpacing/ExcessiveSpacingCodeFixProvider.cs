using System.Collections.Generic;
using System.Collections.Immutable;
using System.Composition;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CodeActions;
using Microsoft.CodeAnalysis.CodeFixes;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Text;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.ExcessiveSpacing
{
    [ExportCodeFixProvider(LanguageNames.CSharp, Name = nameof(ExcessiveSpacingCodeFixProvider)), Shared]
    public class ExcessiveSpacingCodeFixProvider : CodeFixProvider
    {
        public sealed override ImmutableArray<string> FixableDiagnosticIds => ImmutableArray.Create(TerRules.ExcessiveSpacingRule.Id);

        public sealed override FixAllProvider GetFixAllProvider()
            => WellKnownFixAllProviders.BatchFixer;

        public sealed override async Task RegisterCodeFixesAsync(CodeFixContext context)
        {
            SyntaxNode root = await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(false);
            Diagnostic diagnostic = context.Diagnostics.First();
            TextSpan diagnosticSpan = diagnostic.Location.SourceSpan;

            context.RegisterCodeFix(CodeAction.Create(title: "Remove excessive blank lines", createChangedDocument: c => RemoveExcessiveNewlinesAsync(context.Document, root.FindTrivia(diagnosticSpan.Start), c), equivalenceKey: nameof(ExcessiveSpacingCodeFixProvider)), diagnostic);
        }

        private async Task<Document> RemoveExcessiveNewlinesAsync(Document document, SyntaxTrivia trivia, CancellationToken cancellationToken)
        {
            SyntaxNode root = await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(false);
            SyntaxToken token = trivia.Token;

            // Determine if the trivia is leading or trailing
            bool isLeading = token.LeadingTrivia.Contains(trivia);
            SyntaxTriviaList triviaList = isLeading ? token.LeadingTrivia : token.TrailingTrivia;

            // Logic: rebuild the list, collapsing any run of > 2 consecutive newlines down to 2
            // (i.e. at most one blank line). Any non-newline, non-whitespace trivia (comments,
            // #region/#endregion directives, etc.) must be preserved as-is and resets the run,
            // since it isn't part of a blank-line sequence.
            //
            // When operating on leading trivia, the previous token's trailing trivia already
            // consumed one end-of-line (ending its own statement's line), so the run starts
            // at 1 here, not 0 - matching how the analyzer counts across the token boundary.
            SyntaxTriviaList newTriviaList = CollapseExcessiveBlankLines(triviaList, startCount: isLeading ? 1 : 0);

            SyntaxToken newToken = isLeading
                ? token.WithLeadingTrivia(newTriviaList)
                : token.WithTrailingTrivia(newTriviaList);

            SyntaxNode newRoot = root.ReplaceToken(token, newToken);
            return document.WithSyntaxRoot(newRoot);
        }

        private SyntaxTriviaList CollapseExcessiveBlankLines(SyntaxTriviaList list, int startCount)
        {
            var result = new List<SyntaxTrivia>();
            // Whitespace-only trivia seen since the last emitted EOL/content; buffered because
            // we don't yet know if it merely decorates a blank line about to be dropped, or is
            // the final indentation before real content that must be kept.
            var pendingWhitespace = new List<SyntaxTrivia>();
            int consecutiveEol = startCount;

            foreach (SyntaxTrivia trivia in list)
            {
                if (trivia.IsKind(SyntaxKind.EndOfLineTrivia))
                {
                    consecutiveEol++;
                    if (consecutiveEol > 2)
                    {
                        // Excess blank line: drop its newline and any whitespace decorating it.
                        pendingWhitespace.Clear();
                        continue;
                    }

                    result.AddRange(pendingWhitespace);
                    pendingWhitespace.Clear();
                    result.Add(trivia);
                }
                else if (trivia.IsKind(SyntaxKind.WhitespaceTrivia))
                {
                    pendingWhitespace.Add(trivia);
                }
                else
                {
                    // Real content (comment, region directive, etc.): flush any whitespace
                    // preceding it, keep it untouched, and reset the run.
                    result.AddRange(pendingWhitespace);
                    pendingWhitespace.Clear();
                    result.Add(trivia);
                    consecutiveEol = 0;
                }
            }

            result.AddRange(pendingWhitespace);

            return SyntaxFactory.TriviaList(result);
        }

    }
}
