using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

using TerLintAnalyzer.Utility;

namespace TerLintAnalyzer.Ter.PreferStringMethods
{
    internal static class PreferStringMethodsCodeFixProviderHelpers
    {

        internal static (string method, string clean) AnalyzePattern(string pattern)
        {
            if (pattern.StartsWith("^") && pattern.EndsWith("$"))
            {
                return ("Equals", pattern.Substring(1, pattern.Length - 2));
            }

            if (pattern.StartsWith("^"))
            {
                return ("StartsWith", pattern.Substring(1));
            }

            if (pattern.EndsWith("$"))
            {
                return ("EndsWith", pattern.Substring(0, pattern.Length - 1));
            }

            return ("Contains", pattern);
        }

        internal static InvocationExpressionSyntax CreateNewInvocation(InvocationExpressionSyntax originalInvocation, string methodName, string cleanPattern)
        {
            // If it's instance call: regex.IsMatch(input) -> first arg is input
            // If it's static call: Regex.IsMatch(input, pattern) -> first arg is still input
            ArgumentSyntax inputArgument = originalInvocation.ArgumentList.Arguments.FirstOrDefault();
            if (inputArgument == null)
            {
                return originalInvocation;
            }

            return SyntaxFactory.InvocationExpression(SyntaxFactory.MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression, inputArgument.Expression, SyntaxFactory.IdentifierName(methodName)), SyntaxFactory.ArgumentList(SyntaxFactory.SingletonSeparatedList(SyntaxFactory.Argument(SyntaxFactory.LiteralExpression(SyntaxKind.StringLiteralExpression, SyntaxFactory.Literal(cleanPattern)))))).WithTriviaFrom(originalInvocation);
        }

        internal static bool CheckIgnoreCase(ObjectCreationExpressionSyntax objectCreation, SemanticModel semanticModel, ref string methodName)
        {
            bool isIgnoreCase = false;
            if (objectCreation.ArgumentList.Arguments.Count > 1)
            {
                ExpressionSyntax optionsArg = objectCreation.ArgumentList.Arguments[1].Expression;
                Optional<object> optionsValue = semanticModel.GetConstantValue(optionsArg);
                if (optionsValue.HasValue && optionsValue.Value is int intValue)
                {
                    // Cast the underlying int to the RegexOptions enum
                    var options = (System.Text.RegularExpressions.RegexOptions)intValue;

                    if (options.HasFlag(System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    {
                        methodName += "IgnoreCase";
                        isIgnoreCase = true;
                    }
                }
            }

            return isIgnoreCase;
        }

        internal static Document HandleIsMatch(Document document, SyntaxNode root, string methodName, string cleanPattern, bool isIgnoreCase, VariableDeclaratorSyntax declarator, InvocationExpressionSyntax varInv)
        {
            InvocationExpressionSyntax newInv = CreateNewInvocation(varInv, methodName, cleanPattern);

            // Identify if it's a Local Variable or a Field
            LocalDeclarationStatementSyntax localDecl = declarator.FirstAncestorOrSelf<LocalDeclarationStatementSyntax>();
            FieldDeclarationSyntax fieldDecl = declarator.FirstAncestorOrSelf<FieldDeclarationSyntax>();

            SyntaxNode nodeToRemove = (SyntaxNode)localDecl ?? fieldDecl;
            if (nodeToRemove == null)
            {
                return document;
            }

            SyntaxNode trackedRoot = root.TrackNodes(varInv, nodeToRemove);
            SyntaxNode newRoot = trackedRoot.ReplaceNode(trackedRoot.GetCurrentNode(varInv), newInv);

            // Safely remove the declaration (local or field)
            newRoot = newRoot.RemoveNode(newRoot.GetCurrentNode(nodeToRemove), SyntaxRemoveOptions.KeepExteriorTrivia);

            if (isIgnoreCase)
            {
                newRoot = Using.AddUsingDirective(newRoot, "CommonLib.Extension");
            }

            return document.WithSyntaxRoot(newRoot);
        }

        internal static Document HandleIsMatchChainedCall(Document document, SyntaxNode root, string methodName, string cleanPattern, bool isIgnoreCase, InvocationExpressionSyntax invocation)
        {
            InvocationExpressionSyntax newInvocation = CreateNewInvocation(invocation, methodName, cleanPattern);
            SyntaxNode newRoot = root.ReplaceNode(invocation, newInvocation);
            if (isIgnoreCase)
            {
                newRoot = Using.AddUsingDirective(newRoot, "CommonLib.Extension");
            }

            return document.WithSyntaxRoot(newRoot);
        }
    }
}
