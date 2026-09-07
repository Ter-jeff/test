using Microsoft.CodeAnalysis;

namespace TerLintAnalyzer.Ter.Rules
{
    public static partial class TerRules
    {
        public static readonly DiagnosticDescriptor FilenameClassMatchRule = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.NamingId + "01",
            title: "Filename and Class name mismatch",
            messageFormat: "Class name '{0}' must exactly match the filename '{1}'",
            category: "Naming",
            defaultSeverity: DiagnosticSeverity.Warning,
            isEnabledByDefault: true);

        public static readonly DiagnosticDescriptor ArgumentNamingRule = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.NamingId + "02",
            title: "Incorrect argument naming",
            messageFormat: "Argument '{0}' should be named '{1}'",
            category: "Naming",
            defaultSeverity: DiagnosticSeverity.Info,
            isEnabledByDefault: true);

        public static readonly DiagnosticDescriptor EnumNamingRule = new DiagnosticDescriptor(
            id: TerRuleConstants.TerPrefix + TerRuleConstants.NamingId + "03",
            title: "Enum naming convention",
            messageFormat: "Enum '{0}' should start with the prefix 'Enum'",
            category: "Naming",
            defaultSeverity: DiagnosticSeverity.Info,
            isEnabledByDefault: true);
    }
}
