using System.Collections.Immutable;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.Wrapping
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class WrappingAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(TerRules.NoMethodWrappingRule, TerRules.GeneralNoWrappingRule);

        public override void Initialize(AnalysisContext analysisContext)
        {
            analysisContext.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            analysisContext.EnableConcurrentExecution();

            analysisContext.RegisterSyntaxNodeAction(AnalyzeNode,
                SyntaxKind.MethodDeclaration,
                SyntaxKind.ClassDeclaration,
                SyntaxKind.PropertyDeclaration,
                SyntaxKind.NamespaceDeclaration,
                SyntaxKind.FileScopedNamespaceDeclaration,
                SyntaxKind.InvocationExpression);
        }

        private static void AnalyzeNode(SyntaxNodeAnalysisContext context)
        {
            SyntaxNode node = context.Node;
            Location identifierLoc;
            Location endLoc;
            string name;
            DiagnosticDescriptor rule;

            if (node is MethodDeclarationSyntax method)
            {
                // Method Logic (WARNING)
                rule = TerRules.NoMethodWrappingRule;
                name = method.Identifier.Text;
                identifierLoc = method.Identifier.GetLocation();
                endLoc = method.ParameterList.GetLocation();
                if (method.ExpressionBody != null)
                {
                    endLoc = method.ExpressionBody.GetLocation();
                }
            }
            else if (node is PropertyDeclarationSyntax prop)
            {
                // Property Logic (INFO)
                rule = TerRules.GeneralNoWrappingRule;
                name = prop.Identifier.Text;
                identifierLoc = prop.Identifier.GetLocation();
                endLoc = prop.Identifier.GetLocation(); // Check if identifier itself wraps
            }
            else if (node is InvocationExpressionSyntax invocation)
            {
                // 1. Skip if it contains lambdas (your previous requirement), or an argument is a
                // "new" expression with a brace initializer (e.g. Add(new Foo { ... })) - wrapping
                // an object/collection initializer across lines is normal, readable style and
                // isn't the kind of wrapping this rule is meant to flag.
                bool hasLambda = false;
                bool hasNewExpressionWithInitializer = false;
                foreach (ArgumentSyntax argument in invocation.ArgumentList.Arguments)
                {
                    if (argument.Expression is LambdaExpressionSyntax || argument.Expression is AnonymousFunctionExpressionSyntax)
                    {
                        hasLambda = true;
                        break;
                    }

                    if (argument.Expression is ObjectCreationExpressionSyntax objectCreation && objectCreation.Initializer != null)
                    {
                        hasNewExpressionWithInitializer = true;
                    }
                    else if (argument.Expression is ImplicitObjectCreationExpressionSyntax implicitObjectCreation && implicitObjectCreation.Initializer != null)
                    {
                        hasNewExpressionWithInitializer = true;
                    }
                }
                if (hasLambda || hasNewExpressionWithInitializer)
                {
                    return;
                }

                rule = TerRules.GeneralNoWrappingRule;

                // 2. Refined Location Logic:
                // Instead of using invocation.Expression (which includes the whole chain),
                // we only look at the 'Name' of the method being called right now.
                if (invocation.Expression is MemberAccessExpressionSyntax memberAccess)
                {
                    name = memberAccess.Name.Identifier.Text;
                    identifierLoc = memberAccess.Name.GetLocation();
                }
                else
                {
                    name = invocation.Expression.ToString();
                    identifierLoc = invocation.Expression.GetLocation();
                }

                endLoc = invocation.ArgumentList.GetLocation();
            }
            else if (node is ClassDeclarationSyntax cls)
            {
                // Class Logic (INFO)
                rule = TerRules.GeneralNoWrappingRule;
                name = cls.Identifier.Text;
                identifierLoc = cls.Identifier.GetLocation();
                endLoc = cls.Identifier.GetLocation();
            }
            else if (node is BaseNamespaceDeclarationSyntax ns)
            {
                // Namespace Logic (INFO)
                rule = TerRules.GeneralNoWrappingRule;
                name = ns.Name.ToString();
                identifierLoc = ns.Name.GetLocation();
                endLoc = ns.Name.GetLocation();
            }
            else
            {
                return;
            }

            int startLine = identifierLoc.GetLineSpan().StartLinePosition.Line;
            int endLine = endLoc.GetLineSpan().EndLinePosition.Line;

            if (startLine != endLine)
            {
                // For GeneralRule, we pass two args: TypeName and MemberName
                object[] args = rule == TerRules.NoMethodWrappingRule
                    ? new object[] { name }
                    : new object[] { node.Kind().ToString().Replace("Declaration", ""), name };

                context.ReportDiagnostic(Diagnostic.Create(rule, identifierLoc, args));
            }
        }
    }
}
