using System.Collections.Immutable;
using System.IO;
using System.Linq;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.StaticSettable
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class StaticSettableAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(TerRules.StaticSettableRule);

        public override void Initialize(AnalysisContext analysisContext)
        {
            analysisContext.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            analysisContext.EnableConcurrentExecution();

            analysisContext.RegisterSyntaxNodeAction(AnalyzeStaticMember, SyntaxKind.PropertyDeclaration, SyntaxKind.FieldDeclaration);
        }

        private static void AnalyzeStaticMember(SyntaxNodeAnalysisContext context)
        {
            // Requirement: File path must contain "static"
            string filePath = context.Node.SyntaxTree.FilePath;
            string dir = Path.GetDirectoryName(filePath) ?? string.Empty;
            if (string.IsNullOrEmpty(filePath) || dir.ToLower().Contains("static"))
            {
                return;
            }

            if (context.Node is FieldDeclarationSyntax field)
            {
                // Added check: Must be Public
                bool isPublic = field.Modifiers.Any(SyntaxKind.PublicKeyword);
                bool isStatic = field.Modifiers.Any(SyntaxKind.StaticKeyword);
                bool isSettable = !field.Modifiers.Any(SyntaxKind.ReadOnlyKeyword) && !field.Modifiers.Any(SyntaxKind.ConstKeyword);

                if (isPublic && isStatic && isSettable)
                {
                    foreach (VariableDeclaratorSyntax variable in field.Declaration.Variables)
                    {
                        context.ReportDiagnostic(Diagnostic.Create(TerRules.StaticSettableRule, variable.GetLocation(), variable.Identifier.Text));
                    }
                }
            }
            else if (context.Node is PropertyDeclarationSyntax property)
            {
                // Added check: Must be Public
                bool isPublic = property.Modifiers.Any(SyntaxKind.PublicKeyword);
                bool isStatic = property.Modifiers.Any(SyntaxKind.StaticKeyword);
                AccessorDeclarationSyntax setAccessor = property.AccessorList?.Accessors.FirstOrDefault(a => a.IsKind(SyntaxKind.SetAccessorDeclaration));
                bool hasPubliclySettableSetter = setAccessor != null
                    && !setAccessor.Modifiers.Any(m => m.IsKind(SyntaxKind.PrivateKeyword) || m.IsKind(SyntaxKind.ProtectedKeyword) || m.IsKind(SyntaxKind.InternalKeyword));

                if (isPublic && isStatic && hasPubliclySettableSetter)
                {
                    context.ReportDiagnostic(Diagnostic.Create(TerRules.StaticSettableRule, property.Identifier.GetLocation(), property.Identifier.Text));
                }
            }
        }
    }
}
