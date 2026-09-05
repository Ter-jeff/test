using System.Text.RegularExpressions;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace TerLintAnalyzer.Ter.PathSeparator
{
    internal static class PathHeuristics
    {
        private static readonly Regex _pathSegmentPattern = new Regex(
            @"[A-Za-z0-9_\-. ]{2,}\\[A-Za-z0-9_\-. ]{2,}",
            RegexOptions.Compiled);

        private static readonly Regex _driveLetterPattern = new Regex(
            @"^[A-Za-z]:\\",
            RegexOptions.Compiled);

        private static readonly Regex _relativePathPattern = new Regex(
            @"^\.\.?\\",
            RegexOptions.Compiled);

        private static readonly string[] _regexIndicators =
        {
            @"\d", @"\w", @"\s", @"\b", @"\D", @"\W", @"\S", @"\B",
            "(?", "[^", "]+", "]*", ")?", "|", "{0,", "+$", "*$",
            @"\:"
        };

        internal static bool IsPathLikeString(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            if (value == "\\")
            {
                return false;
            }

            if (value.StartsWith("\\\\"))
            {
                return false;
            }

            if (LooksLikeRegex(value))
            {
                return false;
            }

            if (IsRegistryKeyPath(value))
            {
                return false;
            }

            if (_driveLetterPattern.IsMatch(value))
            {
                return true;
            }

            if (_relativePathPattern.IsMatch(value))
            {
                return true;
            }

            if (_pathSegmentPattern.IsMatch(value))
            {
                return true;
            }

            return false;
        }

        internal static bool IsInReplaceBackslashContext(SyntaxNode node)
        {
            if (!(node.Parent is ArgumentSyntax argument))
            {
                return false;
            }

            if (!(argument.Parent is ArgumentListSyntax argList))
            {
                return false;
            }

            if (!(argList.Parent is InvocationExpressionSyntax invocation))
            {
                return false;
            }

            if (!(invocation.Expression is MemberAccessExpressionSyntax memberAccess))
            {
                return false;
            }

            if (memberAccess.Name.Identifier.Text != "Replace")
            {
                return false;
            }

            SeparatedSyntaxList<ArgumentSyntax> args = argList.Arguments;
            if (args.Count != 2)
            {
                return false;
            }

            string firstArgText = args[0].Expression.ToString();
            return firstArgText == "\"/\"" || firstArgText == "@\"/\"" || firstArgText == "'/'";
        }

        internal static bool IsInRegexContext(SyntaxNode node)
        {
            SyntaxNode current = node;
            for (int i = 0; i < 5 && current != null; i++)
            {
                if (current is InvocationExpressionSyntax invocation)
                {
                    string methodName = GetMethodName(invocation);
                    if (methodName != null && IsRegexMethodName(methodName))
                    {
                        return true;
                    }
                }
                else if (current is ObjectCreationExpressionSyntax creation)
                {
                    string typeName = creation.Type.ToString();
                    if (typeName == "Regex" || typeName.EndsWith(".Regex"))
                    {
                        return true;
                    }
                }

                current = current.Parent;
            }

            return false;
        }

        internal static bool IsInPathCombineFirstArgContext(SyntaxNode node)
        {
            if (!(node.Parent is ArgumentSyntax argument))
            {
                return false;
            }

            if (!(argument.Parent is ArgumentListSyntax argList))
            {
                return false;
            }

            if (argList.Arguments.IndexOf(argument) != 0)
            {
                return false;
            }

            if (!(argList.Parent is InvocationExpressionSyntax invocation))
            {
                return false;
            }

            if (!(invocation.Expression is MemberAccessExpressionSyntax memberAccess))
            {
                return false;
            }

            return memberAccess.Name.Identifier.Text == "Combine" &&
                   (memberAccess.Expression.ToString() == "Path" ||
                    memberAccess.Expression.ToString().EndsWith(".Path"));
        }

        private static bool IsRegistryKeyPath(string value)
        {
            return value.StartsWith("SOFTWARE\\") ||
                   value.StartsWith("SYSTEM\\") ||
                   value.StartsWith("HARDWARE\\") ||
                   value.StartsWith("HKEY_");
        }

        internal static bool LooksLikeRegex(string value)
        {
            int matches = 0;
            foreach (string indicator in _regexIndicators)
            {
                if (value.Contains(indicator))
                {
                    matches++;
                }
            }

            // Require at least 2 distinct indicators: a single match is too often a coincidental
            // substring of a real path (e.g. "\Basic" contains the "\B" word-boundary escape),
            // while genuine regex patterns reliably contain multiple special constructs.
            return matches >= 2;
        }

        private static string GetMethodName(InvocationExpressionSyntax invocation)
        {
            switch (invocation.Expression)
            {
                case MemberAccessExpressionSyntax memberAccess:
                    return memberAccess.Name.Identifier.Text;
                case IdentifierNameSyntax identifier:
                    return identifier.Identifier.Text;
                default:
                    return null;
            }
        }

        private static bool IsRegexMethodName(string name)
        {
            switch (name)
            {
                case "Match":
                case "IsMatch":
                case "Replace":
                case "Split":
                case "Matches":
                    return true;
                default:
                    return false;
            }
        }
    }
}
