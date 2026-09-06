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

using TerLintAnalyzer.Utility;

namespace TerLintAnalyzer
{
    [ExportCodeFixProvider(LanguageNames.CSharp, Name = nameof(StringCodeFixProvider)), Shared]
    public class StringCodeFixProvider : CodeFixProvider
    {
        public override ImmutableArray<string> FixableDiagnosticIds => ImmutableArray.Create(
            StringRules.StringEqualsCaseSensitive.Id,
            StringRules.StringEqualsCaseInsensitive.Id,
            StringRules.EqualsCaseSensitive.Id,
            StringRules.EqualsCaseInsensitive.Id,
            StringRules.ToUpperToLowerComparison.Id,
            StringRules.ToUpperToLowerEqualityOperator.Id,
            StringRules.StringComparerUsage.Id
        );

        public override FixAllProvider GetFixAllProvider() => WellKnownFixAllProviders.BatchFixer;

        public override async Task RegisterCodeFixesAsync(CodeFixContext context)
        {
            SyntaxNode root = await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(false);
            Diagnostic diagnostic = context.Diagnostics.First();
            SyntaxNode node = root.FindNode(diagnostic.Location.SourceSpan, getInnermostNodeForTie: true);

            // Handle Property References (StringComparer.OrdinalIgnoreCase)
            if (diagnostic.Id == StringRules.StringComparerUsage.Id)
            {
                context.RegisterCodeFix(CodeAction.Create(title: "Use StringExtensions.IgnoreCase", createChangedDocument: c => FixPropertyToIgnoreCaseAsync(context.Document, node, c), equivalenceKey: nameof(FixPropertyToIgnoreCaseAsync)), diagnostic);
            }
            // Handle Binary (== or !=)
            else if (node is BinaryExpressionSyntax binaryExpr)
            {
                context.RegisterCodeFix(CodeAction.Create(title: "Use EqualsIgnoreCase", createChangedDocument: c => FixEqualityToEqualsIgnoreCaseAsync(context.Document, binaryExpr, c), equivalenceKey: nameof(FixEqualityToEqualsIgnoreCaseAsync)), diagnostic);
            }
            // Handle Invocations (.Equals, .Contains, etc.)
            else if (node.AncestorsAndSelf().OfType<InvocationExpressionSyntax>().FirstOrDefault() is InvocationExpressionSyntax invocation)
            {
                context.RegisterCodeFix(CodeAction.Create(title: "Use IgnoreCase method", createChangedDocument: c => FixInvocationToIgnoreCaseAsync(context.Document, invocation, c), equivalenceKey: nameof(FixInvocationToIgnoreCaseAsync)), diagnostic);
            }
        }

        private async Task<Document> FixPropertyToIgnoreCaseAsync(Document document, SyntaxNode syntaxNode, CancellationToken cancellationToken)
        {
            // 1. Create the expression: StringExtensions.IgnoreCase
            MemberAccessExpressionSyntax expression = SyntaxFactory.MemberAccessExpression(
                SyntaxKind.SimpleMemberAccessExpression,
                SyntaxFactory.IdentifierName("StringExtensions"),
                SyntaxFactory.IdentifierName("IgnoreCase"));

            // 2. Wrap based on the node type using C# 7.3 compatible logic
            SyntaxNode replacementNode;
            if (syntaxNode is ArgumentSyntax)
            {
                replacementNode = SyntaxFactory.Argument(expression);
            }
            else
            {
                replacementNode = expression;
            }

            // 3. Maintain trivia and replace
            replacementNode = replacementNode.WithTriviaFrom(syntaxNode);

            SyntaxNode root = await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(false);
            SyntaxNode newRoot = root.ReplaceNode(syntaxNode, replacementNode);
            newRoot = Using.AddUsingDirective(newRoot, "CommonLib.Extension");
            return document.WithSyntaxRoot(newRoot);
        }

        private async Task<Document> FixEqualityToEqualsIgnoreCaseAsync(Document document, BinaryExpressionSyntax binaryExpr, CancellationToken cancellationToken)
        {
            ExpressionSyntax left = binaryExpr.Left;
            ExpressionSyntax right = binaryExpr.Right;
            ExpressionSyntax receiver;
            ExpressionSyntax comparisonValue;
            if (IsToUpperOrLowerCall(left, out ExpressionSyntax innerExpression))
            {
                receiver = innerExpression;
                comparisonValue = right;
            }
            else if (IsToUpperOrLowerCall(right, out innerExpression))
            {
                receiver = innerExpression;
                comparisonValue = left;
            }
            else
            {
                return document;
            }

            InvocationExpressionSyntax newInvocation = SyntaxFactory.InvocationExpression(SyntaxFactory.MemberAccessExpression(
            SyntaxKind.SimpleMemberAccessExpression, receiver, SyntaxFactory.IdentifierName("EqualsIgnoreCase")), SyntaxFactory.ArgumentList(SyntaxFactory.SingletonSeparatedList(SyntaxFactory.Argument(comparisonValue))));

            ExpressionSyntax finalExpression = binaryExpr.Kind() == SyntaxKind.NotEqualsExpression
                ? SyntaxFactory.PrefixUnaryExpression(SyntaxKind.LogicalNotExpression, newInvocation)
                : (ExpressionSyntax)newInvocation;

            SyntaxNode root = await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(false);

            SyntaxNode newRoot = root.ReplaceNode(binaryExpr, finalExpression.WithTriviaFrom(binaryExpr));

            newRoot = Using.AddUsingDirective(newRoot, "CommonLib.Extension");
            return document.WithSyntaxRoot(newRoot);
        }

        private bool IsToUpperOrLowerCall(ExpressionSyntax expression, out ExpressionSyntax inner)
        {
            inner = null;
            ExpressionSyntax current = expression;
            while (current is ParenthesizedExpressionSyntax parens)
            {
                current = parens.Expression;
            }

            if (current is InvocationExpressionSyntax invocation &&
                invocation.Expression is MemberAccessExpressionSyntax memberAccess &&
                (memberAccess.Name.Identifier.Text == "ToUpper" || memberAccess.Name.Identifier.Text == "ToLower"))
            {
                inner = memberAccess.Expression;
                return true;
            }
            return false;
        }

        private async Task<Document> FixInvocationToIgnoreCaseAsync(Document document, InvocationExpressionSyntax invocation, CancellationToken cancellationToken)
        {
            InvocationExpressionSyntax actualInvocation = invocation;

            // Handle the chaining case: s1.ToUpper().Contains(s2) -> actualInvocation becomes the .Contains call
            if (invocation.Expression is MemberAccessExpressionSyntax ma && (ma.Name.Identifier.Text == "ToUpper" || ma.Name.Identifier.Text == "ToLower"))
            {
                if (invocation.Parent is MemberAccessExpressionSyntax parentMemberAccess && parentMemberAccess.Parent is InvocationExpressionSyntax parentInvocation)
                {
                    actualInvocation = parentInvocation;
                }
            }

            SeparatedSyntaxList<ArgumentSyntax> args = actualInvocation.ArgumentList.Arguments;
            ExpressionSyntax receiver = null;
            ExpressionSyntax comparisonValue = null;
            string targetMethod = "EqualsIgnoreCase";
            bool isIgnoreCase = false;
            HandleMemberAccess(actualInvocation, args, ref receiver, ref comparisonValue, ref targetMethod, ref isIgnoreCase);

            // ... (After CASE 1, 2, 3 logic where actualInvocation is determined) ...

            if (receiver == null || comparisonValue == null)
            {
                return document;
            }

            // 1. Logic must check the parent of the ACTUAL invocation (the .Contains or .Equals)
            bool isNegated = actualInvocation.Parent is PrefixUnaryExpressionSyntax unary && unary.IsKind(SyntaxKind.LogicalNotExpression);

            // 2. We must replace the outer-most node (the ! or the full .Contains call)
            SyntaxNode nodeToReplace = isNegated ? actualInvocation.Parent : actualInvocation;

            ExpressionSyntax finalNode;
            if (targetMethod == "==")
            {
                SyntaxKind operatorKind = isNegated ? SyntaxKind.NotEqualsExpression : SyntaxKind.EqualsExpression;
                finalNode = SyntaxFactory.BinaryExpression(operatorKind, receiver, comparisonValue);
            }
            else
            {
                InvocationExpressionSyntax methodCall = SyntaxFactory.InvocationExpression(SyntaxFactory.MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression, receiver, SyntaxFactory.IdentifierName(targetMethod)), SyntaxFactory.ArgumentList(SyntaxFactory.SingletonSeparatedList(SyntaxFactory.Argument(comparisonValue))));

                finalNode = isNegated ? SyntaxFactory.PrefixUnaryExpression(SyntaxKind.LogicalNotExpression, methodCall) : (ExpressionSyntax)methodCall;
            }

            SyntaxNode root = await document.GetSyntaxRootAsync(cancellationToken);
            // 3. Replace actualInvocation (or its negation) with the new node
            root = root.ReplaceNode(nodeToReplace, finalNode.WithTriviaFrom(nodeToReplace));
            if (isIgnoreCase)
            {
                root = Using.AddUsingDirective(root, "CommonLib.Extension");
            }
            return document.WithSyntaxRoot(root);
        }

        private void HandleMemberAccess(InvocationExpressionSyntax actualInvocation, SeparatedSyntaxList<ArgumentSyntax> args, ref ExpressionSyntax receiver, ref ExpressionSyntax comparisonValue, ref string targetMethod, ref bool isIgnoreCase)
        {
            if (actualInvocation.Expression is MemberAccessExpressionSyntax memberAccess)
            {
                string methodName = memberAccess.Name.Identifier.Text;
                if ((methodName == "StartsWith" || methodName == "EndsWith") && args.Count >= 2)
                {
                    receiver = memberAccess.Expression;
                    comparisonValue = args.FirstOrDefault()?.Expression;
                    targetMethod = GetMethodName(args.Last().Expression.ToFullString().Contains("IgnoreCase"), methodName);
                    isIgnoreCase = true;
                }
                // CASE 1: Static string.Equals(s1, s2, ...)
                else if (methodName == "Equals" && args.Count >= 2)
                {
                    HandleEquals(args, out receiver, out comparisonValue, out targetMethod, out isIgnoreCase, memberAccess, methodName);
                }
                // CASE 2: ToUpper/ToLower chained with method (e.g., s1.ToUpper().Contains(s2))
                else if (IsToUpperOrLowerCall(memberAccess.Expression, out ExpressionSyntax innerReceiver))
                {
                    receiver = innerReceiver;
                    comparisonValue = args.FirstOrDefault()?.Expression;
                    // Map Contains -> ContainsIgnoreCase, etc.
                    targetMethod = GetMethodName(methodName.EndsWith("IgnoreCase"), methodName);
                    isIgnoreCase = true;
                }
                // CASE 3: Standard Instance Equals (s1.Equals(s2, StringComparison...))
                else
                {
                    isIgnoreCase = args.Last().Expression.ToFullString().Contains("IgnoreCase");
                    if (!isIgnoreCase && args.Count() == 1)
                    {
                        receiver = memberAccess.Expression;
                        comparisonValue = args[0].Expression;
                        targetMethod = "==";
                    }
                    else
                    {
                        receiver = memberAccess.Expression;
                        comparisonValue = args.FirstOrDefault()?.Expression;
                        targetMethod = methodName;
                    }
                }
            }
        }

        private void HandleEquals(SeparatedSyntaxList<ArgumentSyntax> args, out ExpressionSyntax receiver, out ExpressionSyntax comparisonValue, out string targetMethod, out bool isIgnoreCase, MemberAccessExpressionSyntax memberAccess, string methodName)
        {
            bool isStaticString = false;
            if (memberAccess.Expression is IdentifierNameSyntax id && id.Identifier.Text == "String")
            {
                isStaticString = true;
            }
            else if (memberAccess.Expression is PredefinedTypeSyntax predefined && predefined.Keyword.IsKind(SyntaxKind.StringKeyword))
            {
                isStaticString = true;
            }

            if (isStaticString)
            {
                isIgnoreCase = args.Last().Expression.ToFullString().Contains("IgnoreCase");
                if (!isIgnoreCase)
                {
                    receiver = args[0].Expression;
                    comparisonValue = args[1].Expression;
                    targetMethod = "==";
                }
                else
                {
                    receiver = args[0].Expression;
                    comparisonValue = args[1].Expression;
                    targetMethod = "EqualsIgnoreCase";
                }
            }
            else
            {
                receiver = memberAccess.Expression;
                comparisonValue = args.FirstOrDefault()?.Expression;
                isIgnoreCase = args.Last().Expression.ToFullString().Contains("IgnoreCase");
                targetMethod = GetMethodName(isIgnoreCase, methodName);
            }
        }

        private string GetMethodName(bool isIgnoreCase, string methodName)
        {
            if (methodName.EndsWith("Equals") || methodName.EndsWith("StartsWith") || methodName.EndsWith("EndsWith"))
            {
                return isIgnoreCase ? methodName + "IgnoreCase" : methodName;
            }

            return methodName + "IgnoreCase";
        }
    }
}
