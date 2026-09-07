using System.Collections.Immutable;
using System.Linq;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.MemberOrder
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class MemberOrderAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(TerRules.MemberOrderRule);

        public override void Initialize(AnalysisContext analysisContext)
        {
            analysisContext.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            analysisContext.EnableConcurrentExecution();

            analysisContext.RegisterSyntaxNodeAction(AnalyzeMemberOrder, SyntaxKind.ClassDeclaration, SyntaxKind.InterfaceDeclaration);
        }

        private static void AnalyzeMemberOrder(SyntaxNodeAnalysisContext context)
        {
            SyntaxList<MemberDeclarationSyntax> members;
            if (context.Node is ClassDeclarationSyntax classDecl)
            {
                members = classDecl.Members;
            }
            else if (context.Node is InterfaceDeclarationSyntax interfaceDecl)
            {
                members = interfaceDecl.Members;
            }
            else
            {
                return;
            }

            if (members.Count < 2)
            {
                return;
            }

            int lastWeight = -1;
            MemberDeclarationSyntax lastMember = null;

            foreach (MemberDeclarationSyntax member in members)
            {
                int currentWeight = GetMemberWeight(member);

                if (currentWeight < lastWeight)
                {
                    context.ReportDiagnostic(Diagnostic.Create(TerRules.MemberOrderRule, member.GetLocation(), GetMemberTypeName(member), GetMemberTypeName(lastMember)));
                }

                lastWeight = currentWeight;
                lastMember = member;
            }
        }

        private static bool IsRegex(SyntaxNode node)
        {
            // Detect [GeneratedRegex] attribute (C# 11+)
            if (node is MemberDeclarationSyntax member &&
                member.AttributeLists.Any(al => al.Attributes.Any(a => a.Name.ToString().Contains("GeneratedRegex"))))
            {
                return true;
            }

            // Detect Regex type in field declarations
            if (node is FieldDeclarationSyntax field)
            {
                return field.Declaration.Type.ToString().Contains("Regex");
            }

            return false;
        }

        internal static int GetMemberWeight(MemberDeclarationSyntax member)
        {
            if (member is ClassDeclarationSyntax || member is EnumDeclarationSyntax || member is StructDeclarationSyntax)
            {
                return 1;
            }

            if (member is FieldDeclarationSyntax field)
            {
                if (field.Modifiers.Any(SyntaxKind.ConstKeyword))
                {
                    return 2;
                }

                if (IsRegex(field))
                {
                    return 3;
                }

                return 4; // General Field
            }

            if (member is PropertyDeclarationSyntax || member is EventFieldDeclarationSyntax)
            {
                return 4;
            }

            if (member is MethodDeclarationSyntax method)
            {
                if (IsRegex(method))
                {
                    return 3;
                }

                return 6;
            }

            if (member is ConstructorDeclarationSyntax)
            {
                return 5;
            }

            return 7;
        }

        private static string GetMemberTypeName(MemberDeclarationSyntax member)
        {
            if (member is FieldDeclarationSyntax f && f.Modifiers.Any(SyntaxKind.ConstKeyword))
            {
                return "Constant";
            }

            if (IsRegex(member))
            {
                return "Regex";
            }

            if (member is FieldDeclarationSyntax)
            {
                return "Field";
            }

            if (member is PropertyDeclarationSyntax)
            {
                return "Property";
            }

            if (member is ConstructorDeclarationSyntax)
            {
                return "Constructor";
            }

            if (member is MethodDeclarationSyntax)
            {
                return "Method";
            }

            return member.ToString();
        }
    }
}
