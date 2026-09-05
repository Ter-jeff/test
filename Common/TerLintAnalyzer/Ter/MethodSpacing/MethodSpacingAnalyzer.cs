using System.Collections.Immutable;
using System.Linq;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.MethodSpacing
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class MethodSpacingAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(TerRules.MethodSpacingRule);

        public override void Initialize(AnalysisContext analysisContext)
        {
            analysisContext.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            analysisContext.EnableConcurrentExecution();

            analysisContext.RegisterSyntaxNodeAction(AnalyzeMethodSpacing, SyntaxKind.MethodDeclaration);
        }

        private static void AnalyzeMethodSpacing(SyntaxNodeAnalysisContext context)
        {
            var methodDeclaration = (MethodDeclarationSyntax)context.Node;

            // Find the sibling immediately before this method
            SyntaxNode previousNode = methodDeclaration.Parent.ChildNodes().TakeWhile(node => node != methodDeclaration).LastOrDefault();

            // Rule: Only check if the previous sibling is also a method
            if (previousNode is MethodDeclarationSyntax)
            {
                // Only the blank-line run immediately before the method's leading content
                // (its first attached comment/region-start/attribute, if any) matters.
                // An #endregion is treated as transparent (it closes a section for the
                // *previous* method, not decoration for this one) so it's skipped like
                // whitespace. Note its own directive line absorbs its trailing end-of-line
                // internally (Roslyn attaches it to the directive's EndOfDirectiveToken),
                // so no separate EndOfLineTrivia is seen for that line. Any other non-blank
                // trivia (comment, region-start, attribute) marks the real content boundary:
                // stop there, don't keep accumulating unrelated blank-line gaps from deeper
                // in the trivia list (e.g. inside a nested region that follows).
                SyntaxTriviaList leadingTrivia = methodDeclaration.GetLeadingTrivia();
                int actualNewlines = 0;
                foreach (SyntaxTrivia trivia in leadingTrivia)
                {
                    if (trivia.IsKind(SyntaxKind.EndOfLineTrivia))
                    {
                        actualNewlines++;
                    }
                    else if (trivia.IsKind(SyntaxKind.WhitespaceTrivia) || trivia.IsKind(SyntaxKind.EndRegionDirectiveTrivia))
                    {
                        // Skip: doesn't count as blank, doesn't stop the scan.
                    }
                    else
                    {
                        break;
                    }
                }

                if (actualNewlines != 1)
                {
                    context.ReportDiagnostic(Diagnostic.Create(TerRules.MethodSpacingRule, methodDeclaration.Identifier.GetLocation()));
                }
            }
        }
    }
}
