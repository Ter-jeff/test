using Microsoft.CodeAnalysis;

namespace TerLintAnalyzer
{
    /// <summary>
    /// Diagnostic descriptors for string comparison rules
    /// </summary>
    public static class StringRules
    {
        private const string RulePrefix = RuleConstants.RulePrefix + "3";

        /// <summary>
        /// Usage of case-sensitive string.Equals()
        /// </summary>
        public static readonly DiagnosticDescriptor StringEqualsCaseSensitive = new DiagnosticDescriptor(
            id: RulePrefix + "01",
            title: "Usage of case-sensitive string.Equals()",
            messageFormat: "Usage of 'string.Equals(x, y)' should be replaced by 'x == y'",
            category: RuleConstants.UsageWarning,
            defaultSeverity: DiagnosticSeverity.Error,
            isEnabledByDefault: true,
            description: "Should use 'x == y' instead of 'string.Equals(x, y)'."
        );

        /// <summary>
        /// Usage of case-insensitive string.Equals()
        /// </summary>
        public static readonly DiagnosticDescriptor StringEqualsCaseInsensitive = new DiagnosticDescriptor(
            id: RulePrefix + "02",
            title: "Usage of case-insensitive string.Equals()",
            messageFormat: "Usage of 'string.Equals(x, y, StringComparison.{0})' should be replaced by 'x.EqualsIgnoreCase(y)'",
            category: RuleConstants.UsageWarning,
            defaultSeverity: DiagnosticSeverity.Error,
            isEnabledByDefault: true,
            description: "Should use 'x.EqualsIgnoreCase(y)' instead of 'string.Equals(x, y, StringComparison)'."
        );

        /// <summary>
        /// Usage of case-sensitive .Equals()
        /// </summary>
        public static readonly DiagnosticDescriptor EqualsCaseSensitive = new DiagnosticDescriptor(
            id: RulePrefix + "03",
            title: "Usage of case-sensitive Equals()",
            messageFormat: "Usage of 'x.Equals(y)' should be replaced by 'x == y'",
            category: RuleConstants.UsageWarning,
            defaultSeverity: DiagnosticSeverity.Error,
            isEnabledByDefault: true,
            description: "Should use 'x == y' instead of 'x.Equals(y)'."
        );

        /// <summary>
        /// Usage of case-insensitive .Equals()
        /// </summary>
        public static readonly DiagnosticDescriptor EqualsCaseInsensitive = new DiagnosticDescriptor(
            id: RulePrefix + "04",
            title: "Usage of case-insensitive Equals()",
            messageFormat: "Usage of 'x.Equals(y, StringComparison.{0})' should be replaced by 'x.EqualsIgnoreCase(y)'",
            category: RuleConstants.UsageWarning,
            defaultSeverity: DiagnosticSeverity.Error,
            isEnabledByDefault: true,
            description: "Should use 'x.EqualsIgnoreCase(y)' instead of 'x.Equals(y, StringComparison)'."
        );

        /// <summary>
        /// Usage of ToUpper()/ToLower() in string comparison using Equals(), Contains(), StartsWith(), or EndsWith()
        /// </summary>
        public static readonly DiagnosticDescriptor ToUpperToLowerComparison = new DiagnosticDescriptor(
            id: RulePrefix + "05",
            title: "Usage of ToUpper()/ToLower() in string comparison",
            messageFormat: "Combined usage of {0}() with {1}() should be replaced by {1}IgnoreCase()",
            category: RuleConstants.UsageWarning,
            defaultSeverity: DiagnosticSeverity.Error,
            isEnabledByDefault: true,
            description: "Should use case-insensitive extension method instead of ToUpper()/ToLower() with Equals()/Contains()/StartsWith()/EndsWith()."
        );

        /// <summary>
        /// Usage of ToUpper()/ToLower() in string comparison using == or !=
        /// </summary>
        public static readonly DiagnosticDescriptor ToUpperToLowerEqualityOperator = new DiagnosticDescriptor(
            id: RulePrefix + "06",
            title: "Usage of ToUpper()/ToLower() in string equality comparison",
            messageFormat: "Combined usage of {0}() with {1} operator should be replaced by {2}EqualsIgnoreCase()",
            category: RuleConstants.UsageWarning,
            defaultSeverity: DiagnosticSeverity.Error,
            isEnabledByDefault: true,
            description: "Should use EqualsIgnoreCase() extension method instead of ToUpper/ToLower() with == or != operator."
        );

        /// <summary>
        /// Direct usage of StringComparer
        /// </summary>
        public static readonly DiagnosticDescriptor StringComparerUsage = new DiagnosticDescriptor(
            id: RulePrefix + "07",
            title: "Direct usage of StringComparer",
            messageFormat: "Direct usage of StringComparer should be replaced with StringExtensions.IgnoreCase",
            category: RuleConstants.UsageWarning,
            defaultSeverity: DiagnosticSeverity.Warning,
            isEnabledByDefault: true,
            description: "Should use StringExtensions.IgnoreCase instead of StringComparer."
        );

        /// <summary>
        /// Usage of manual null checks on strings instead of IsNullOrEmpty
        /// </summary>
        public static readonly DiagnosticDescriptor StringIsNullOrEmptyUsage = new DiagnosticDescriptor(
            id: RulePrefix + "08", // Adjust ID sequential number based on your existing rules
            title: "Use string.IsNullOrEmpty",
            messageFormat: "Expression '{0}' can be simplified to '{1}'",
            category: RuleConstants.UsageWarning,
            defaultSeverity: DiagnosticSeverity.Info,
            isEnabledByDefault: true,
            description: "Use string.IsNullOrEmpty(value) instead of explicit null checks."
        );
    }
}
