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

namespace TerLintAnalyzer.Ter.MemberOrder
{
    [ExportCodeFixProvider(LanguageNames.CSharp, Name = nameof(MemberOrderCodeFixProvider)), Shared]
    public class MemberOrderCodeFixProvider : CodeFixProvider
    {
        public sealed override ImmutableArray<string> FixableDiagnosticIds => ImmutableArray.Create(TerRules.MemberOrderRule.Id);

        public sealed override FixAllProvider GetFixAllProvider() => WellKnownFixAllProviders.BatchFixer;

        public sealed override async Task RegisterCodeFixesAsync(CodeFixContext context)
        {
            SyntaxNode root = await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(false);
            Diagnostic diagnostic = context.Diagnostics.First();
            SyntaxNode node = root?.FindNode(diagnostic.Location.SourceSpan);
            MemberDeclarationSyntax member = node?.FirstAncestorOrSelf<MemberDeclarationSyntax>();

            if (!(member?.Parent is TypeDeclarationSyntax typeDecl))
            {
                return;
            }

            context.RegisterCodeFix(
                CodeAction.Create(
                    title: "Move member to its correct position",
                    createChangedDocument: _ => Task.FromResult(MoveMember(context.Document, root, typeDecl, member)),
                    equivalenceKey: nameof(MemberOrderCodeFixProvider)),
                diagnostic);
        }

        private static Document MoveMember(Document document, SyntaxNode root, TypeDeclarationSyntax typeDecl, MemberDeclarationSyntax member)
        {
            int weight = MemberOrderAnalyzer.GetMemberWeight(member);

            SyntaxList<MemberDeclarationSyntax> withoutMember = SyntaxFactory.List(typeDecl.Members.Where(m => m != member));
            MemberDeclarationSyntax insertBefore = withoutMember.FirstOrDefault(m => MemberOrderAnalyzer.GetMemberWeight(m) > weight);
            int insertIndex = insertBefore == null ? withoutMember.Count : withoutMember.IndexOf(insertBefore);

            SyntaxList<MemberDeclarationSyntax> newMembers = withoutMember.Insert(insertIndex, member);
            TypeDeclarationSyntax newTypeDecl = typeDecl.WithMembers(newMembers);

            SyntaxNode newRoot = root.ReplaceNode(typeDecl, newTypeDecl);
            return document.WithSyntaxRoot(newRoot);
        }
    }
}
