using System.Collections.Immutable;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.PathSeparator
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class PathSeparatorAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(
            TerRules.PathBackslashInLiteral,
            TerRules.PathBackslashInConcatenation
        );

        public override void Initialize(AnalysisContext context)
        {
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            context.EnableConcurrentExecution();

            context.RegisterSyntaxNodeAction(AnalyzeStringLiteral, SyntaxKind.StringLiteralExpression);
            context.RegisterSyntaxNodeAction(AnalyzeInterpolatedString, SyntaxKind.InterpolatedStringExpression);
            context.RegisterSyntaxNodeAction(AnalyzeAddExpression, SyntaxKind.AddExpression);
        }

        private static void AnalyzeStringLiteral(SyntaxNodeAnalysisContext context)
        {
            var literal = (LiteralExpressionSyntax)context.Node;
            string value = literal.Token.ValueText;

            if (!value.Contains("\\"))
            {
                return;
            }

            if (!PathHeuristics.IsPathLikeString(value))
            {
                return;
            }

            if (PathHeuristics.IsInReplaceBackslashContext(literal))
            {
                return;
            }

            if (PathHeuristics.IsInRegexContext(literal))
            {
                return;
            }

            if (PathHeuristics.IsInPathCombineFirstArgContext(literal))
            {
                return;
            }

            context.ReportDiagnostic(Diagnostic.Create(TerRules.PathBackslashInLiteral, literal.GetLocation()));
        }

        private static void AnalyzeInterpolatedString(SyntaxNodeAnalysisContext context)
        {
            var interpolated = (InterpolatedStringExpressionSyntax)context.Node;

            if (interpolated.Contents.Count > 0 &&
                interpolated.Contents[0] is InterpolatedStringTextSyntax firstText &&
                firstText.TextToken.ValueText.StartsWith("\\\\"))
            {
                return;
            }

            foreach (InterpolatedStringContentSyntax content in interpolated.Contents)
            {
                if (!(content is InterpolatedStringTextSyntax textPart))
                {
                    continue;
                }

                string text = textPart.TextToken.ValueText;
                if (!text.Contains("\\"))
                {
                    continue;
                }

                if (IsBackslashSeparatorInInterpolation(text))
                {
                    if (PathHeuristics.IsInRegexContext(interpolated))
                    {
                        return;
                    }

                    if (PathHeuristics.IsInReplaceBackslashContext(interpolated))
                    {
                        return;
                    }

                    context.ReportDiagnostic(Diagnostic.Create(TerRules.PathBackslashInConcatenation, interpolated.GetLocation()));
                    return;
                }
            }
        }

        private static void AnalyzeAddExpression(SyntaxNodeAnalysisContext context)
        {
            var binary = (BinaryExpressionSyntax)context.Node;

            if (binary.Parent is BinaryExpressionSyntax parentBinary &&
                parentBinary.IsKind(SyntaxKind.AddExpression))
            {
                return;
            }

            bool hasBackslashSeparator = false;
            bool hasPathLikeSegment = false;

            foreach (ExpressionSyntax part in FlattenConcatenation(binary))
            {
                if (!(part is LiteralExpressionSyntax literal))
                {
                    continue;
                }

                if (!literal.IsKind(SyntaxKind.StringLiteralExpression))
                {
                    continue;
                }

                string value = literal.Token.ValueText;

                if (value == "\\" || value == "\\\\")
                {
                    hasBackslashSeparator = true;
                }
                else if (value.Contains("\\") && PathHeuristics.IsPathLikeString(value))
                {
                    hasPathLikeSegment = true;
                }
            }

            if (!hasBackslashSeparator && !hasPathLikeSegment)
            {
                return;
            }

            if (hasBackslashSeparator)
            {
                if (PathHeuristics.IsInReplaceBackslashContext(binary))
                {
                    return;
                }

                if (PathHeuristics.IsInRegexContext(binary))
                {
                    return;
                }

                context.ReportDiagnostic(Diagnostic.Create(TerRules.PathBackslashInConcatenation, binary.GetLocation()));
            }
        }

        private static bool IsBackslashSeparatorInInterpolation(string text)
        {
            string trimmed = text.Trim();
            return trimmed == "\\" || trimmed.StartsWith("\\") || trimmed.EndsWith("\\");
        }

        private static System.Collections.Generic.IEnumerable<ExpressionSyntax> FlattenConcatenation(BinaryExpressionSyntax binary)
        {
            if (binary.Left is BinaryExpressionSyntax leftBinary && leftBinary.IsKind(SyntaxKind.AddExpression))
            {
                foreach (ExpressionSyntax part in FlattenConcatenation(leftBinary))
                {
                    yield return part;
                }
            }
            else
            {
                yield return binary.Left;
            }

            if (binary.Right is BinaryExpressionSyntax rightBinary && rightBinary.IsKind(SyntaxKind.AddExpression))
            {
                foreach (ExpressionSyntax part in FlattenConcatenation(rightBinary))
                {
                    yield return part;
                }
            }
            else
            {
                yield return binary.Right;
            }
        }
    }
}
