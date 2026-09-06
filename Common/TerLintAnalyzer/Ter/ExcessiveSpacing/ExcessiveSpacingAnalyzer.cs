using System.Collections.Immutable;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Diagnostics;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.ExcessiveSpacing
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class ExcessiveSpacingAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(TerRules.ExcessiveSpacingRule);

        public override void Initialize(AnalysisContext analysisContext)
        {
            analysisContext.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            analysisContext.EnableConcurrentExecution();

            analysisContext.RegisterSyntaxTreeAction(AnalyzeSyntaxTree);
        }

        private static void AnalyzeSyntaxTree(SyntaxTreeAnalysisContext context)
        {
            SyntaxNode root = context.Tree.GetRoot(context.CancellationToken);
            int consecutiveNewlines = 0;

            foreach (SyntaxToken token in root.DescendantTokens())
            {
                foreach (SyntaxTrivia leading in token.LeadingTrivia)
                {
                    if (leading.IsKind(SyntaxKind.EndOfLineTrivia))
                    {
                        consecutiveNewlines++;

                        if (consecutiveNewlines == 3)
                        {
                            context.ReportDiagnostic(Diagnostic.Create(TerRules.ExcessiveSpacingRule, leading.GetLocation()));
                        }
                    }
                    else if (IsComment(leading))
                    {
                        consecutiveNewlines = 0;
                    }
                }

                consecutiveNewlines = 0;

                foreach (SyntaxTrivia trailing in token.TrailingTrivia)
                {
                    if (trailing.IsKind(SyntaxKind.EndOfLineTrivia))
                    {
                        consecutiveNewlines++;

                        if (consecutiveNewlines == 3)
                        {
                            context.ReportDiagnostic(Diagnostic.Create(TerRules.ExcessiveSpacingRule, trailing.GetLocation()));
                        }
                    }
                    else if (IsComment(trailing))
                    {
                        consecutiveNewlines = 0;
                    }
                }
            }
        }

        private static bool IsComment(SyntaxTrivia trivia)
        {
            return trivia.IsKind(SyntaxKind.SingleLineCommentTrivia) ||
                   trivia.IsKind(SyntaxKind.MultiLineCommentTrivia) ||
                   trivia.IsKind(SyntaxKind.DocumentationCommentExteriorTrivia) ||
                   trivia.IsKind(SyntaxKind.SingleLineDocumentationCommentTrivia) ||
                   trivia.IsKind(SyntaxKind.MultiLineDocumentationCommentTrivia) ||
                   trivia.IsKind(SyntaxKind.RegionDirectiveTrivia) ||
                   trivia.IsKind(SyntaxKind.EndRegionDirectiveTrivia);
        }
    }
}
