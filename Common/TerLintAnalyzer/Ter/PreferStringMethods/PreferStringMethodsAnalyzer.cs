using System.Collections.Immutable;
using System.Linq;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.PreferStringMethods
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class PreferStringMethodsAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(TerRules.PreferStringMethodsRule);

        public override void Initialize(AnalysisContext analysisContext)
        {
            analysisContext.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            analysisContext.EnableConcurrentExecution();

            analysisContext.RegisterSyntaxNodeAction(AnalyzeObjectCreation, SyntaxKind.ObjectCreationExpression);
        }

        private void AnalyzeObjectCreation(SyntaxNodeAnalysisContext context)
        {
            var creation = (ObjectCreationExpressionSyntax)context.Node;
            if (IsRegexType(context, creation) && creation.ArgumentList?.Arguments.Count > 0)
            {
                // Constructor: new Regex(pattern, options) -> index 0 is pattern
                AnalyzePattern(context, creation.ArgumentList.Arguments[0].Expression);
            }
        }

        private void AnalyzePattern(SyntaxNodeAnalysisContext context, ExpressionSyntax patternExpr)
        {
            Optional<object> constant = context.SemanticModel.GetConstantValue(patternExpr);
            if (constant.HasValue && constant.Value is string pattern && !string.IsNullOrWhiteSpace(pattern))
            {
                string suggestion;
                string cleanPattern;

                // 1. Determine the suggestion and clean the pattern of anchors
                if (pattern.StartsWith("^") && pattern.EndsWith("$"))
                {
                    suggestion = "Equals";
                    cleanPattern = pattern.Substring(1, pattern.Length - 2);
                }
                else if (pattern.StartsWith("^"))
                {
                    suggestion = "StartsWith";
                    cleanPattern = pattern.Substring(1);
                }
                else if (pattern.EndsWith("$"))
                {
                    suggestion = "EndsWith";
                    cleanPattern = pattern.Substring(0, pattern.Length - 1);
                }
                else
                {
                    suggestion = "Contains/Replace";
                    cleanPattern = pattern;
                }

                // 2. Validate the pattern is actually "simple" text
                if (IsSimpleText(cleanPattern))
                {
                    context.ReportDiagnostic(Diagnostic.Create(TerRules.PreferStringMethodsRule, patternExpr.GetLocation(), pattern, suggestion));
                }
            }
        }

        private bool IsSimpleText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            // Allows alphanumeric, spaces, underscores (to catch _CS_), and common dashes
            // Rejects regex meta-chars: . $ ^ { [ ( | ) * + ? \
            return text.All(c => char.IsLetterOrDigit(c) || c == ' ' || c == '_' || c == '-');
        }

        private bool IsRegexType(SyntaxNodeAnalysisContext context, SyntaxNode node)
        {
            ISymbol symbol = context.SemanticModel.GetSymbolInfo(node).Symbol;
            INamedTypeSymbol type = null;

            if (symbol is IMethodSymbol methodSymbol)
            {
                type = methodSymbol.ContainingType;
            }
            else if (symbol is INamedTypeSymbol typeSymbol)
            {
                type = typeSymbol;
            }

            return type?.ToDisplayString() == "System.Text.RegularExpressions.Regex";
        }
    }
}
