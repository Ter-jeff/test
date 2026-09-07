using Microsoft.CodeAnalysis;

namespace TerLintAnalyzer.Ter.Rules
{
    public static partial class TerRules
    {
        public static readonly DiagnosticDescriptor RegexInlineRule = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.PerformanceId + "01",
            title: "Avoid inline Regex.IsMatch",
            messageFormat: "Inline Regex.IsMatch call should be replaced with a static readonly Regex field",
            category: "Performance",
            defaultSeverity: DiagnosticSeverity.Warning,
            isEnabledByDefault: true);

        public static readonly DiagnosticDescriptor PreferStringMethodsRule = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.PerformanceId + "02",
            title: "Prefer String Methods over Regex",
            messageFormat: "Regex pattern '{0}' can be replaced by 'string.{1}'",
            category: "Performance",
            defaultSeverity: DiagnosticSeverity.Warning,
            isEnabledByDefault: true);
    }
}
