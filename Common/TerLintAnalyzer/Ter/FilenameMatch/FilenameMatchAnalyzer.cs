using System.Collections.Immutable;
using System.IO;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.FilenameMatch
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class FilenameMatchAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(TerRules.FilenameClassMatchRule);

        public override void Initialize(AnalysisContext analysisContext)
        {
            analysisContext.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            analysisContext.EnableConcurrentExecution();

            analysisContext.RegisterSyntaxNodeAction(AnalyzeNaming, SyntaxKind.ClassDeclaration);
        }

        private static void AnalyzeNaming(SyntaxNodeAnalysisContext context)
        {
            var classDeclarationSyntax = (ClassDeclarationSyntax)context.Node;
            //check by first class only
            SyntaxNode syntaxNode = classDeclarationSyntax.Parent;
            if (!(syntaxNode is NamespaceDeclarationSyntax) && !(syntaxNode is CompilationUnitSyntax))
            {
                return;
            }
            SyntaxList<MemberDeclarationSyntax> members = syntaxNode is NamespaceDeclarationSyntax namespaceDeclarationSyntax ? namespaceDeclarationSyntax.Members : ((CompilationUnitSyntax)syntaxNode).Members;
            if (members.Count == 0 || members[0] != classDeclarationSyntax)
            {
                return;
            }
            string className = classDeclarationSyntax.Identifier.Text;
            string filePath = classDeclarationSyntax.SyntaxTree.FilePath;
            if (string.IsNullOrEmpty(filePath))
            {
                return;
            }

            string fileName = Path.GetFileNameWithoutExtension(filePath);
            // Ex: CSharpCodeFixVerifier`2+Test.cs
            if (fileName.Contains("`"))
            {
                return;
            }

            // Handle partial classes with dotted filenames (e.g., EpplusExtensions.Helpers.cs)
            bool isPartial = classDeclarationSyntax.Modifiers.Any(SyntaxKind.PartialKeyword);
            if (isPartial && fileName.Contains("."))
            {
                fileName = fileName.Split('.')[0];
            }

            if (className != fileName)
            {
                context.ReportDiagnostic(Diagnostic.Create(TerRules.FilenameClassMatchRule, classDeclarationSyntax.Identifier.GetLocation(), className, fileName));
            }
        }
    }
}
