using System.Collections.Immutable;
using System.Linq;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Diagnostics;

namespace TerLintAnalyzer
{
    /// <summary>
    /// Custom Roslyn analyzer
    /// </summary>
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class CommentAnalyzer : DiagnosticAnalyzer
    {
        /// <inheritdoc />
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(CommentRules.TrailingComment);

        /// <inheritdoc />
        public override void Initialize(AnalysisContext analysisContext)
        {
            analysisContext.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            analysisContext.EnableConcurrentExecution();

            analysisContext.RegisterSyntaxTreeAction(AnalyzeSyntaxTree);
        }

        private static void AnalyzeSyntaxTree(SyntaxTreeAnalysisContext context)
        {
            SyntaxTrivia[] comments = context.Tree.GetRoot().DescendantTrivia().Where(x => x.IsKind(SyntaxKind.SingleLineCommentTrivia) || x.IsKind(SyntaxKind.MultiLineCommentTrivia)).ToArray();
            foreach (SyntaxTrivia comment in comments)
            {
                CheckForTrailingComment(context, comment);
            }
        }

        private static void CheckForTrailingComment(SyntaxTreeAnalysisContext context, SyntaxTrivia comment)
        {
            // No leading trivia means it is not on its own line
            if (comment.IsKind(SyntaxKind.SingleLineCommentTrivia) && !comment.Token.HasLeadingTrivia && comment.Token.IsKind(SyntaxKind.SemicolonToken))
            {
                Location location = comment.GetLocation();
                context.ReportDiagnostic(Diagnostic.Create(CommentRules.TrailingComment, location));
            }
        }
    }
}
