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
using Microsoft.CodeAnalysis.Editing;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.RegexInline
{
    [ExportCodeFixProvider(LanguageNames.CSharp, Name = nameof(RegexInlineCodeFixProvider)), Shared]
    public class RegexInlineCodeFixProvider : CodeFixProvider
    {
        public sealed override ImmutableArray<string> FixableDiagnosticIds => ImmutableArray.Create(TerRules.RegexInlineRule.Id);

        public sealed override FixAllProvider GetFixAllProvider() => FixAllProvider.Create(async (fixAllContext, document, diagnostics) =>
        {
            return await FixAllInDocumentAsync(document, diagnostics, fixAllContext.CancellationToken);
        });

        public sealed override async Task RegisterCodeFixesAsync(CodeFixContext context)
        {
            SyntaxNode root = await context.Document.GetSyntaxRootAsync(context.CancellationToken);
            InvocationExpressionSyntax invocation = root.FindNode(context.Span).FirstAncestorOrSelf<InvocationExpressionSyntax>();

            context.RegisterCodeFix(CodeAction.Create(title: "Extract to static Regex field", createChangedDocument: c => FixAllInDocumentAsync(context.Document, ImmutableArray.Create(context.Diagnostics[0]), c), equivalenceKey: nameof(RegexInlineCodeFixProvider)), context.Diagnostics[0]);
        }

        private async Task<Document> FixAllInDocumentAsync(Document document, IEnumerable<Diagnostic> diagnostics, CancellationToken ct)
        {
            // DocumentEditor is essential for making multiple sequential changes correctly
            DocumentEditor editor = await DocumentEditor.CreateAsync(document, ct);
            SyntaxNode root = await document.GetSyntaxRootAsync(ct);

            // Group by class to handle field insertion per class and avoid variable name clashes
            IEnumerable<IGrouping<ClassDeclarationSyntax, Diagnostic>> groups = diagnostics.GroupBy(d => root.FindNode(d.Location.SourceSpan).FirstAncestorOrSelf<ClassDeclarationSyntax>());

            foreach (IGrouping<ClassDeclarationSyntax, Diagnostic> group in groups)
            {
                ClassDeclarationSyntax classDecl = group.Key;
                if (classDecl == null)
                {
                    continue;
                }

                // Populate used names from existing fields to avoid name collisions
                var usedNames = new HashSet<string>(classDecl.Members
                    .OfType<FieldDeclarationSyntax>()
                    .SelectMany(f => f.Declaration.Variables)
                    .Select(v => v.Identifier.Text));

                var newFields = new List<FieldDeclarationSyntax>();

                // Sort diagnostics by span descending (bottom-to-top) to maintain node positions
                // This is a best-practice for manual tree mutations.
                IOrderedEnumerable<Diagnostic> sortedDiagnostics = group.OrderByDescending(d => d.Location.SourceSpan.Start);

                foreach (Diagnostic diagnostic in sortedDiagnostics)
                {
                    // Use the editor's tracking to find the exact node for this specific diagnostic
                    SyntaxNode node = editor.OriginalRoot.FindNode(diagnostic.Location.SourceSpan);
                    InvocationExpressionSyntax invocation = node.FirstAncestorOrSelf<InvocationExpressionSyntax>();

                    if (invocation == null)
                    {
                        continue;
                    }

                    // Generate unique name (e.g., _regex, _regex1, _regex2)
                    string fieldName = "_regex";
                    for (int i = 1; usedNames.Contains(fieldName); i++)
                    {
                        fieldName = "_regex" + i;
                    }

                    usedNames.Add(fieldName);

                    SeparatedSyntaxList<ArgumentSyntax> arguments = invocation.ArgumentList.Arguments;
                    if (arguments.Count < 2)
                    {
                        continue;
                    }

                    // Create the new static field
                    FieldDeclarationSyntax fieldDecl = RegexInlineCodeFixProviderHelpers.CreateRegexField(fieldName, arguments[1].Expression, arguments.Skip(2));
                    newFields.Add(fieldDecl);

                    // Create replacement: _regexX.IsMatch(input)
                    InvocationExpressionSyntax newCall = SyntaxFactory.InvocationExpression(
                        SyntaxFactory.MemberAccessExpression(
                            SyntaxKind.SimpleMemberAccessExpression,
                            SyntaxFactory.IdentifierName(fieldName),
                            SyntaxFactory.IdentifierName("IsMatch")))
                        .WithArgumentList(SyntaxFactory.ArgumentList(SyntaxFactory.SingletonSeparatedList(
                            SyntaxFactory.Argument(arguments[0].Expression))))
                        .WithTriviaFrom(invocation);

                    // editor.ReplaceNode handles the complexity of replacing identical-looking nodes
                    editor.ReplaceNode(invocation, newCall);
                }

                if (newFields.Any())
                {
                    // Insert all collected fields at the start of the class members list
                    editor.InsertMembers(classDecl, 0, newFields);
                }
            }

            return editor.GetChangedDocument();
        }
    }
}
