using System.Collections.Immutable;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.RegexInline
{

    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class RegexInlineAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(TerRules.RegexInlineRule);

        public override void Initialize(AnalysisContext analysisContext)
        {
            analysisContext.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            analysisContext.EnableConcurrentExecution();

            analysisContext.RegisterSyntaxNodeAction(AnalyzeInvocation, SyntaxKind.InvocationExpression);
        }

        private void AnalyzeInvocation(SyntaxNodeAnalysisContext context)
        {
            var invocation = (InvocationExpressionSyntax)context.Node;
            if (invocation.Expression is MemberAccessExpressionSyntax memberAccess &&
                memberAccess.Name.Identifier.Text == "IsMatch" &&
                context.SemanticModel.GetSymbolInfo(memberAccess).Symbol is IMethodSymbol method &&
                method.ContainingType.ToString() == "System.Text.RegularExpressions.Regex" &&
                method.IsStatic)
            {
                // Only flag if the pattern (2nd argument) is a literal string
                if (invocation.ArgumentList.Arguments.Count >= 2 &&
                    invocation.ArgumentList.Arguments[1].Expression is LiteralExpressionSyntax)
                {
                    context.ReportDiagnostic(Diagnostic.Create(TerRules.RegexInlineRule, invocation.GetLocation()));
                }
            }
        }
    }
}
