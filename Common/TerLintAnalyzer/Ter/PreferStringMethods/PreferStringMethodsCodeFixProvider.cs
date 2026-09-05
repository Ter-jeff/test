using System.Collections.Generic;
using System.Collections.Immutable;
using System.Composition;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CodeActions;
using Microsoft.CodeAnalysis.CodeFixes;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.FindSymbols;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.PreferStringMethods
{
    [ExportCodeFixProvider(LanguageNames.CSharp, Name = nameof(PreferStringMethodsCodeFixProvider)), Shared]
    public class PreferStringMethodsCodeFixProvider : CodeFixProvider
    {
        public sealed override ImmutableArray<string> FixableDiagnosticIds => ImmutableArray.Create(TerRules.PreferStringMethodsRule.Id);

        public sealed override FixAllProvider GetFixAllProvider() => WellKnownFixAllProviders.BatchFixer;

        public sealed override async Task RegisterCodeFixesAsync(CodeFixContext context)
        {
            SyntaxNode root = await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(false);
            Diagnostic diagnostic = context.Diagnostics.First();
            SyntaxNode node = root.FindNode(diagnostic.Location.SourceSpan);

            ObjectCreationExpressionSyntax objectCreation = node.FirstAncestorOrSelf<ObjectCreationExpressionSyntax>();
            if (objectCreation == null)
            {
                return;
            }

            context.RegisterCodeFix(CodeAction.Create(title: "Simplify to string method", createChangedDocument: c => ReplaceWithStringMethodAsync(context.Document, objectCreation, c), equivalenceKey: "SimplifyToStringMethod"), diagnostic);
        }

        private async Task<Document> ReplaceWithStringMethodAsync(Document document, ObjectCreationExpressionSyntax objectCreation, CancellationToken cancellationToken)
        {
            SemanticModel semanticModel = await document.GetSemanticModelAsync(cancellationToken);
            SyntaxNode root = await document.GetSyntaxRootAsync(cancellationToken);

            // 1. Extract Pattern and Options
            ArgumentSyntax patternArgument = objectCreation.ArgumentList?.Arguments.FirstOrDefault();
            if (patternArgument == null)
            {
                return document;
            }

            Optional<object> constant = semanticModel.GetConstantValue(patternArgument.Expression);
            if (!constant.HasValue || !(constant.Value is string patternValue))
            {
                return document;
            }

            (string methodName, string cleanPattern) = PreferStringMethodsCodeFixProviderHelpers.AnalyzePattern(patternValue);

            // Correctly check for IgnoreCase using Semantic Model
            bool isIgnoreCase = PreferStringMethodsCodeFixProviderHelpers.CheckIgnoreCase(objectCreation, semanticModel, ref methodName);

            // 2. Case A: Immediate chained call -> new Regex(...).IsMatch(input)
            if (objectCreation.Parent is MemberAccessExpressionSyntax memberAccess &&
                memberAccess.Parent is InvocationExpressionSyntax invocation &&
                memberAccess.Name.Identifier.Text == "IsMatch")
            {
                return PreferStringMethodsCodeFixProviderHelpers.HandleIsMatchChainedCall(document, root, methodName, cleanPattern, isIgnoreCase, invocation);
            }

            // 3. Case B: Assignment (Local or Field)
            if (objectCreation.Parent is EqualsValueClauseSyntax equals && equals.Parent is VariableDeclaratorSyntax declarator)
            {
                ISymbol variableSymbol = semanticModel.GetDeclaredSymbol(declarator);
                if (variableSymbol == null)
                {
                    return document;
                }

                IEnumerable<ReferencedSymbol> references = await SymbolFinder.FindReferencesAsync(variableSymbol, document.Project.Solution, cancellationToken);
                var allLocations = references.SelectMany(r => r.Locations).ToList();

                // To keep it simple and safe: Only fix if there is exactly one usage
                if (allLocations.Count != 1)
                {
                    return document;
                }

                SyntaxNode usageNode = root.FindNode(allLocations.First().Location.SourceSpan);
                if (usageNode.Parent is MemberAccessExpressionSyntax varMA &&
                    varMA.Parent is InvocationExpressionSyntax varInv &&
                    varMA.Name.Identifier.Text == "IsMatch")
                {
                    return PreferStringMethodsCodeFixProviderHelpers.HandleIsMatch(document, root, methodName, cleanPattern, isIgnoreCase, declarator, varInv);
                }
            }

            return document;
        }
    }
}
