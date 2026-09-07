using System;
using System.Collections.Immutable;
using System.Linq;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Operations;

namespace TerLintAnalyzer
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class StringAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(
            StringRules.StringEqualsCaseSensitive,
            StringRules.StringEqualsCaseInsensitive,
            StringRules.EqualsCaseSensitive,
            StringRules.EqualsCaseInsensitive,
            StringRules.ToUpperToLowerComparison,
            StringRules.ToUpperToLowerEqualityOperator,
            StringRules.StringComparerUsage,
            StringRules.StringIsNullOrEmptyUsage
        );

        public override void Initialize(AnalysisContext analysisContext)
        {
            analysisContext.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            analysisContext.EnableConcurrentExecution();

            analysisContext.RegisterOperationAction(AnalyzeOperation,
                OperationKind.Invocation,
                OperationKind.PropertyReference,
                OperationKind.BinaryOperator);
        }

        private static void AnalyzeOperation(OperationAnalysisContext context)
        {
            switch (context.Operation)
            {
                case IInvocationOperation invocation when invocation.TargetMethod.ContainingType.SpecialType == SpecialType.System_String:
                    AnalyzeStringOperation(context, invocation);
                    break;
                case IPropertyReferenceOperation property:
                    AnalyzePropertyReference(context, property);
                    break;
                case IBinaryOperation binaryOp:
                    AnalyzeBinaryOperation(context, binaryOp);
                    break;
            }
        }

        private static void AnalyzeBinaryOperation(OperationAnalysisContext context, IBinaryOperation binaryOp)
        {
            // Restructured to strictly look for equality (==) expressions only
            if (binaryOp.OperatorKind == BinaryOperatorKind.Equals)
            {
                IOperation left = UnwrapImplicitConversion(binaryOp.LeftOperand);
                IOperation right = UnwrapImplicitConversion(binaryOp.RightOperand);

                bool isLeftNull = left is ILiteralOperation leftLit && leftLit.ConstantValue.HasValue && leftLit.ConstantValue.Value == null;
                bool isRightNull = right is ILiteralOperation rightLit && rightLit.ConstantValue.HasValue && rightLit.ConstantValue.Value == null;

                if (!isLeftNull && left.Syntax.IsKind(Microsoft.CodeAnalysis.CSharp.SyntaxKind.NullLiteralExpression))
                {
                    isLeftNull = true;
                }

                if (!isRightNull && right.Syntax.IsKind(Microsoft.CodeAnalysis.CSharp.SyntaxKind.NullLiteralExpression))
                {
                    isRightNull = true;
                }

                if (!isLeftNull && !isRightNull)
                {
                    return;
                }

                IOperation valueOperand = isLeftNull ? right : left;
                ITypeSymbol type = valueOperand.Type;

                if (type == null)
                {
                    if (valueOperand is IPropertyReferenceOperation propRef)
                    {
                        type = propRef.Property.Type;
                    }
                    else if (valueOperand is IFieldReferenceOperation fieldRef)
                    {
                        type = fieldRef.Field.Type;
                    }
                }

                if (type?.SpecialType == SpecialType.System_String || type?.ToDisplayString() == "string")
                {
                    string variableName = valueOperand.Syntax.ToString();

                    // Simplified: Since we only handle BinaryOperatorKind.Equals, the suggestion is always positive
                    string suggestion = $"string.IsNullOrEmpty({variableName})";

                    context.ReportDiagnostic(Diagnostic.Create(
                        StringRules.StringIsNullOrEmptyUsage,
                        binaryOp.Syntax.GetLocation(),
                        variableName,
                        suggestion));
                }
            }
        }

        private static IOperation UnwrapImplicitConversion(IOperation operation)
        {
            // Recursively strip away conversion wrappers to see the underlying node
            while (operation is IConversionOperation conversion && conversion.IsImplicit)
            {
                operation = conversion.Operand;
            }
            return operation;
        }

        private static void AnalyzePropertyReference(OperationAnalysisContext context, IPropertyReferenceOperation operation)
        {
            IPropertySymbol property = operation.Property;
            if (property.IsStatic && property.ContainingType.Name == "StringComparer")
            {
                context.ReportDiagnostic(Diagnostic.Create(StringRules.StringComparerUsage, operation.Syntax.GetLocation()));
            }
        }

        private static void AnalyzeStringOperation(OperationAnalysisContext context, IInvocationOperation invocation)
        {
            IMethodSymbol method = invocation.TargetMethod;
            bool isEquals = method.Name == "Equals";
            bool isCheckMethod = method.Name == "StartsWith" || method.Name == "EndsWith";

            if (isEquals || isCheckMethod)
            {
                bool isStatic = method.IsStatic;
                IOperation stringComparisonArg = invocation.Arguments.FirstOrDefault(x => x.Parameter?.Type.Name == "StringComparison")?.Value;

                string argValue = null;
                if (stringComparisonArg?.ConstantValue.HasValue == true)
                {
                    argValue = Enum.GetName(typeof(StringComparison), stringComparisonArg.ConstantValue.Value);
                }

                if (argValue != null && argValue.Contains("IgnoreCase"))
                {
                    context.ReportDiagnostic(Diagnostic.Create(isStatic ? StringRules.StringEqualsCaseInsensitive : StringRules.EqualsCaseInsensitive, invocation.Syntax.GetLocation(), argValue));
                }
                else if (isEquals)
                {
                    context.ReportDiagnostic(Diagnostic.Create(isStatic ? StringRules.StringEqualsCaseSensitive : StringRules.EqualsCaseSensitive, invocation.Syntax.GetLocation()));
                }
            }

            HandleToUpperToLower(context, invocation, method);
        }

        private static void HandleToUpperToLower(OperationAnalysisContext context, IInvocationOperation invocation, IMethodSymbol method)
        {
            if (method.Name == "ToUpper" || method.Name == "ToLower")
            {
                IOperation current = invocation.Parent;

                while (current is IInvocationOperation parentInvoke && !IsStringComparison(parentInvoke.TargetMethod))
                {
                    current = parentInvoke.Parent;
                }

                if (current is IInvocationOperation targetInvoke && IsStringComparison(targetInvoke.TargetMethod))
                {
                    context.ReportDiagnostic(Diagnostic.Create(StringRules.ToUpperToLowerComparison, invocation.Syntax.GetLocation(), method.Name, targetInvoke.TargetMethod.Name));
                }
                else if (current is IBinaryOperation binaryOp && (binaryOp.OperatorKind == BinaryOperatorKind.Equals || binaryOp.OperatorKind == BinaryOperatorKind.NotEquals))
                {
                    bool isEquals = binaryOp.OperatorKind == BinaryOperatorKind.Equals;
                    context.ReportDiagnostic(Diagnostic.Create(StringRules.ToUpperToLowerEqualityOperator, binaryOp.Syntax.GetLocation(), method.Name, isEquals ? "==" : "!=", isEquals ? "" : "!"));
                }
            }
        }

        private static bool IsStringComparison(IMethodSymbol method)
        {
            if (method.ContainingType.SpecialType != SpecialType.System_String)
            {
                return false;
            }

            switch (method.Name)
            {
                case "Equals":
                case "Contains":
                case "StartsWith":
                case "EndsWith":
                case "IndexOf":
                case "LastIndexOf":
                    return true;
                default:
                    return false;
            }
        }
    }
}
