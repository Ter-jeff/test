using System.Collections.Immutable;
using System.Composition;
using System.Linq;
using System.Threading.Tasks;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CodeActions;
using Microsoft.CodeAnalysis.CodeFixes;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.StaticSettable
{
    [ExportCodeFixProvider(LanguageNames.CSharp, Name = nameof(StaticSettableCodeFixProvider)), Shared]
    public class StaticSettableCodeFixProvider : CodeFixProvider
    {
        public sealed override ImmutableArray<string> FixableDiagnosticIds => ImmutableArray.Create(TerRules.StaticSettableRule.Id);

        public sealed override FixAllProvider GetFixAllProvider() => WellKnownFixAllProviders.BatchFixer;

        public sealed override async Task RegisterCodeFixesAsync(CodeFixContext context)
        {
            SyntaxNode root = await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(false);
            Diagnostic diagnostic = context.Diagnostics.First();
            SyntaxNode node = root?.FindNode(diagnostic.Location.SourceSpan);

            if (node == null)
            {
                return;
            }

            VariableDeclaratorSyntax variable = node.FirstAncestorOrSelf<VariableDeclaratorSyntax>();
            if (variable?.Parent?.Parent is FieldDeclarationSyntax fieldDeclaration)
            {
                RegisterFieldFix(context, root, fieldDeclaration, variable, diagnostic);
                return;
            }

            PropertyDeclarationSyntax property = node.FirstAncestorOrSelf<PropertyDeclarationSyntax>();
            if (property != null)
            {
                RegisterPropertyFix(context, root, property, diagnostic);
            }
        }

        private static void RegisterFieldFix(CodeFixContext context, SyntaxNode root, FieldDeclarationSyntax fieldDeclaration, VariableDeclaratorSyntax variable, Diagnostic diagnostic)
        {
            // Multiple variables sharing one field declaration (e.g. "public static int A, B;")
            // would require splitting the declaration apart; out of scope for a mechanical fix.
            if (fieldDeclaration.Declaration.Variables.Count != 1)
            {
                return;
            }

            context.RegisterCodeFix(
                CodeAction.Create(
                    title: "Convert to auto-property with private setter",
                    createChangedDocument: _ => Task.FromResult(ConvertFieldToProperty(context.Document, root, fieldDeclaration, variable)),
                    equivalenceKey: "ConvertFieldToPrivateSetProperty"),
                diagnostic);
        }

        private static Document ConvertFieldToProperty(Document document, SyntaxNode root, FieldDeclarationSyntax fieldDeclaration, VariableDeclaratorSyntax variable)
        {
            string modifiers = string.Join(" ", fieldDeclaration.Modifiers.Select(m => m.Text));
            string type = fieldDeclaration.Declaration.Type.ToString();
            string name = variable.Identifier.Text;
            string initializer = variable.Initializer != null ? $" {variable.Initializer};" : string.Empty;

            string propertyText = $"{modifiers} {type} {name} {{ get; private set; }}{initializer}";

            var property = (PropertyDeclarationSyntax)SyntaxFactory.ParseMemberDeclaration(propertyText);
            property = property
                .WithAttributeLists(fieldDeclaration.AttributeLists)
                .WithTriviaFrom(fieldDeclaration);

            SyntaxNode newRoot = root.ReplaceNode(fieldDeclaration, property);
            return document.WithSyntaxRoot(newRoot);
        }

        private static void RegisterPropertyFix(CodeFixContext context, SyntaxNode root, PropertyDeclarationSyntax property, Diagnostic diagnostic)
        {
            AccessorDeclarationSyntax setAccessor = property.AccessorList?.Accessors.FirstOrDefault(a => a.IsKind(SyntaxKind.SetAccessorDeclaration));

            // Nothing mechanical to do if there's no plain (unrestricted) setter to lock down.
            if (setAccessor == null || setAccessor.Modifiers.Any(m => m.IsKind(SyntaxKind.PrivateKeyword) || m.IsKind(SyntaxKind.ProtectedKeyword) || m.IsKind(SyntaxKind.InternalKeyword)))
            {
                return;
            }

            context.RegisterCodeFix(
                CodeAction.Create(
                    title: "Make setter private",
                    createChangedDocument: _ => Task.FromResult(MakeSetterPrivate(context.Document, root, setAccessor)),
                    equivalenceKey: "MakePropertySetterPrivate"),
                diagnostic);
        }

        private static Document MakeSetterPrivate(Document document, SyntaxNode root, AccessorDeclarationSyntax setAccessor)
        {
            SyntaxToken privateKeyword = SyntaxFactory.Token(SyntaxKind.PrivateKeyword)
                .WithLeadingTrivia(setAccessor.GetLeadingTrivia())
                .WithTrailingTrivia(SyntaxFactory.Space);

            AccessorDeclarationSyntax newSetAccessor = setAccessor
                .WithModifiers(SyntaxFactory.TokenList(privateKeyword))
                .WithKeyword(setAccessor.Keyword.WithLeadingTrivia());

            SyntaxNode newRoot = root.ReplaceNode(setAccessor, newSetAccessor);
            return document.WithSyntaxRoot(newRoot);
        }
    }
}
