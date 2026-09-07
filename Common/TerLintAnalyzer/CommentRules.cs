using Microsoft.CodeAnalysis;

namespace TerLintAnalyzer
{
    /// <summary>
    /// Diagnostic descriptors for ATE analyzer rules with the prefix FujiLint
    /// </summary>
    public static class CommentRules
    {
        private const string RulePrefix = "FujiLint";

        /// <summary>
        /// Single line comment should go above the line it is commenting
        /// </summary>
        public static readonly DiagnosticDescriptor TrailingComment = new DiagnosticDescriptor(
            id: RulePrefix + "FujiLint001",
            title: "Single line comment is on same line as code",
            messageFormat: "Single line comment should go above the line it is commenting",
            category: "FujiStyle",
            defaultSeverity: DiagnosticSeverity.Info,
            isEnabledByDefault: true,
            description: "Comment should go above the line of code that it is commenting."
        );
    }
}
