using System.Collections.Immutable;
using System.Composition;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CodeActions;
using Microsoft.CodeAnalysis.CodeFixes;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Rename;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.MethodArgumentNaming
{
    [ExportCodeFixProvider(LanguageNames.CSharp, Name = nameof(MethodArgumentNamingCodeFixProvider)), Shared]
    public class MethodArgumentNamingCodeFixProvider : CodeFixProvider
    {
        public sealed override ImmutableArray<string> FixableDiagnosticIds => ImmutableArray.Create(TerRules.ArgumentNamingRule.Id);

        public sealed override FixAllProvider GetFixAllProvider() => WellKnownFixAllProviders.BatchFixer;

        public sealed override async Task RegisterCodeFixesAsync(CodeFixContext context)
        {
            SyntaxNode root = await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(false);
            Diagnostic diagnostic = context.Diagnostics.First();

            ParameterSyntax parameterSyntax = root.FindToken(diagnostic.Location.SourceSpan.Start)
                .Parent.AncestorsAndSelf().OfType<ParameterSyntax>().FirstOrDefault();

            if (parameterSyntax != null && diagnostic.Properties.TryGetValue("ExpectedName", out string expectedName))
            {
                context.RegisterCodeFix(CodeAction.Create(title: $"Rename parameter to '{expectedName}'", createChangedSolution: c => RenameInSolutionAsync(context.Document.Project.Solution, context.Document.Id, diagnostic, c), equivalenceKey: nameof(MethodArgumentNamingCodeFixProvider)),
                    diagnostic);
            }
        }

        private static async Task<Solution> RenameInSolutionAsync(Solution solution, DocumentId docId, Diagnostic diagnostic, CancellationToken ct)
        {
            Document document = solution.GetDocument(docId);
            SyntaxNode root = await document.GetSyntaxRootAsync(ct).ConfigureAwait(false);

            ParameterSyntax parameterSyntax = root.FindToken(diagnostic.Location.SourceSpan.Start).Parent.AncestorsAndSelf().OfType<ParameterSyntax>().FirstOrDefault();

            if (parameterSyntax != null && diagnostic.Properties.TryGetValue("ExpectedName", out string expectedName))
            {
                SemanticModel semanticModel = await document.GetSemanticModelAsync(ct).ConfigureAwait(false);
                ISymbol parameterSymbol = semanticModel.GetDeclaredSymbol(parameterSyntax, ct);

                if (parameterSymbol != null)
                {
                    // Renamer.RenameSymbolAsync handles updating all references across the solution
                    return await Renamer.RenameSymbolAsync(solution, parameterSymbol, new SymbolRenameOptions(), expectedName, ct).ConfigureAwait(false);
                }
            }

            return solution;
        }
    }
}
