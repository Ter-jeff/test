using System.Collections.Immutable;
using System.Linq;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;

using TerLintAnalyzer.Ter.Rules;

namespace TerLintAnalyzer.Ter.MethodArgumentNaming
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class MethodArgumentNamingAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(TerRules.ArgumentNamingRule);

        public override void Initialize(AnalysisContext analysisContext)
        {
            analysisContext.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            analysisContext.EnableConcurrentExecution();

            analysisContext.RegisterSymbolAction(AnalyzeMethodSymbol, SymbolKind.Method);
        }

        private static void AnalyzeMethodSymbol(SymbolAnalysisContext symbolAnalysisContext)
        {
            var methodSymbol = (IMethodSymbol)symbolAnalysisContext.Symbol;

            if (methodSymbol.AssociatedSymbol is IPropertySymbol)
            {
                return;
            }

            var parameterGroups = methodSymbol.Parameters
                .GroupBy(p => p.Type, SymbolEqualityComparer.Default)
                .ToList();

            foreach (IGrouping<ISymbol, IParameterSymbol> group in parameterGroups)
            {
                if (group.Count() > 1)
                {
                    continue;
                }

                IParameterSymbol parameterSymbol = group.First();
                string expectedName = GetExpectedName(parameterSymbol.Type);

                if (!string.IsNullOrEmpty(expectedName) && parameterSymbol.Name != expectedName)
                {
                    ImmutableDictionary<string, string> properties = ImmutableDictionary<string, string>.Empty.Add("ExpectedName", expectedName);
                    symbolAnalysisContext.ReportDiagnostic(Diagnostic.Create(TerRules.ArgumentNamingRule, parameterSymbol.Locations[0], properties, parameterSymbol.Name, expectedName));
                }
            }
        }

        private static string GetExpectedName(ITypeSymbol typeSymbol)
        {
            if (typeSymbol.SpecialType != SpecialType.None)
            {
                return string.Empty;
            }

            if (typeSymbol is IArrayTypeSymbol arrayType)
            {
                ITypeSymbol elementType = arrayType.ElementType;
                if (IsInvalidOrComplex(elementType))
                {
                    return string.Empty;
                }
                return ToCamelCase(elementType.Name) + "Array";
            }

            if (typeSymbol is INamedTypeSymbol namedType)
            {
                if (namedType.IsGenericType)
                {
                    // Check if it implements or is IEnumerable<T>
                    bool isCollection = namedType.OriginalDefinition.SpecialType == SpecialType.System_Collections_Generic_IEnumerable_T ||
                                       namedType.AllInterfaces.Any(i => i.OriginalDefinition.SpecialType == SpecialType.System_Collections_Generic_IEnumerable_T);

                    if (isCollection && namedType.TypeArguments.Length == 1)
                    {
                        ITypeSymbol elementType = namedType.TypeArguments[0];
                        if (IsInvalidOrComplex(elementType))
                        {
                            return string.Empty;
                        }

                        // Route List<T> or explicit IEnumerable<T> to the "s" rule
                        bool isListOrEnumerable = namedType.MetadataName == "List`1" || namedType.MetadataName == "IEnumerable`1" || namedType.OriginalDefinition.SpecialType == SpecialType.System_Collections_Generic_IEnumerable_T;

                        if (isListOrEnumerable)
                        {
                            return ToCamelCase(elementType.Name) + "s"; // e.g., employees
                        }

                        return ToCamelCase(elementType.Name) + "Collect"; // e.g., employeeCollect
                    }
                }

                if (namedType.IsGenericType && namedType.TypeArguments.Length == 1)
                {
                    ITypeSymbol t = namedType.TypeArguments[0];
                    if (IsInvalidOrComplex(t))
                    {
                        return string.Empty;
                    }
                    return ToCamelCase(t.Name) + "s";
                }

                if (namedType.IsGenericType)
                {
                    return string.Empty;
                }
            }

            return ToCamelCase(typeSymbol.Name);
        }

        // Helper to deduplicate your validation logic
        private static bool IsInvalidOrComplex(ITypeSymbol symbol)
        {
            return symbol == null ||
                   symbol.SpecialType != SpecialType.None ||
                   (symbol is INamedTypeSymbol inner && inner.IsGenericType) ||
                   symbol is IArrayTypeSymbol;
        }


        private static string ToCamelCase(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return name;
            }

            if (name.Length > 1 && name[0] == 'I' && char.IsUpper(name[1]))
            {
                name = name.Substring(1);
            }
            return char.ToLower(name[0]) + name.Substring(1);
        }
    }
}
