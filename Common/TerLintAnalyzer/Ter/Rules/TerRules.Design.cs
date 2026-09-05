using Microsoft.CodeAnalysis;

namespace TerLintAnalyzer.Ter.Rules
{
    public static partial class TerRules
    {
        public static readonly DiagnosticDescriptor StaticSettableRule = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.DesignId + "01",
            title: "Settable Static Member Detected",
            messageFormat: "Static member '{0}' is settable",
            category: "Design",
            defaultSeverity: DiagnosticSeverity.Info,
            isEnabledByDefault: true);

        public static readonly DiagnosticDescriptor PathBackslashInLiteral = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.DesignId + "02",
            title: "Backslash path separator in string literal",
            messageFormat: "String literal contains a Windows-style path with backslash separators. Use Path.Combine() instead.",
            category: "Design",
            defaultSeverity: DiagnosticSeverity.Warning,
            isEnabledByDefault: true);

        public static readonly DiagnosticDescriptor PathBackslashInConcatenation = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.DesignId + "03",
            title: "String concatenation builds path with backslash",
            messageFormat: "String concatenation or interpolation builds a file path using backslash separators. Use Path.Combine() instead.",
            category: "Design",
            defaultSeverity: DiagnosticSeverity.Warning,
            isEnabledByDefault: true);
    }
}
