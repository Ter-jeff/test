using Microsoft.CodeAnalysis;

namespace TerLintAnalyzer.Ter.Rules
{
    public static partial class TerRules
    {
        public static readonly DiagnosticDescriptor MethodSpacingRule = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.LayoutId + "01",
            title: "Incorrect method spacing",
            messageFormat: "There should be exactly one blank line between method definitions",
            category: "Layout",
            defaultSeverity: DiagnosticSeverity.Warning,
            isEnabledByDefault: true);

        public static readonly DiagnosticDescriptor NoMethodWrappingRule = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.LayoutId + "02",
            title: "Method signature must be on a single line",
            messageFormat: "Method '{0}' is wrapped across multiple lines",
            category: "Layout",
            defaultSeverity: DiagnosticSeverity.Warning,
            isEnabledByDefault: true);

        public static readonly DiagnosticDescriptor GeneralNoWrappingRule = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.LayoutId + "03",
            title: "Declaration must be on a single line",
            messageFormat: "{0} '{1}' is wrapped across multiple lines",
            category: "Layout",
            defaultSeverity: DiagnosticSeverity.Info,
            isEnabledByDefault: true);

        public static readonly DiagnosticDescriptor ExcessiveSpacingRule = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.LayoutId + "04",
            title: "Excessive blank lines detected",
            messageFormat: "This file contains 2 or more consecutive empty lines",
            category: "Layout",
            defaultSeverity: DiagnosticSeverity.Info,
            isEnabledByDefault: true,
            description: "Avoid using multiple consecutive empty lines to keep code compact and readable.");
    }
}
