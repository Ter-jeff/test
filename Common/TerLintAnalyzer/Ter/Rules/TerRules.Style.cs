using Microsoft.CodeAnalysis;

namespace TerLintAnalyzer.Ter.Rules
{
    public static partial class TerRules
    {
        public static readonly DiagnosticDescriptor MemberOrderRule = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.StyleId + "01",
            title: "Member order is incorrect",
            messageFormat: "Member '{0}' should not appear after '{1}'",
            category: "Style",
            defaultSeverity: DiagnosticSeverity.Info,
            isEnabledByDefault: true);
    }
}
