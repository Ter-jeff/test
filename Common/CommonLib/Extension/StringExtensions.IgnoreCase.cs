#pragma warning disable
using System;
using System.Text.RegularExpressions;

namespace CommonLib.Extension
{
    public static partial class StringExtensions
    {
        /// <summary>
        /// Project-wide constant for use with case-insensitive operations
        /// </summary>
        private const StringComparison IgnoreCaseComparison = StringComparison.OrdinalIgnoreCase;

        /// <summary>
        /// Project-wide getter for use with case-insensitive operations
        /// </summary>
        public static StringComparer IgnoreCase => StringComparer.OrdinalIgnoreCase;

        /// <summary>Determines whether this string and a specified <see cref="T:System.String" /> object have the same value, ignoring case.</summary>
        /// <param name="source">Source string.</param>
        /// <param name="value">The string to compare to this instance.</param>
        /// <returns>
        /// <see langword="true" /> if the value of the <paramref name="value" /> parameter is the same as the value of this instance; otherwise, <see langword="false" />. If <paramref name="value" /> is <see langword="null" />, the method returns <see langword="false" />.
        /// </returns>
        public static bool EqualsIgnoreCase(this string source, string value)
        {
            if (source == null)
            {
                if (value == null)
                {
                    return true;
                }
                return false;
            }
            return string.Equals(source, value, IgnoreCaseComparison);
        }

        public static int IndexOfIgnoreCase(this string source, string value)
        {
            if (source == null)
            {
                return -1;
            }

            return source.IndexOf(value, IgnoreCaseComparison);
        }

        public static bool ContainsIgnoreCase(this string source, string value)
        {
            if (source == null)
            {
                return value == null;
            }

            return source.IndexOf(value, IgnoreCaseComparison) >= 0;
        }

        public static bool StartsWithIgnoreCase(this string source, string value)
        {
            if (source == null)
            {
                return value == null;
            }

            if (value == null)
            {
                return false;
            }

            return source.StartsWith(value, IgnoreCaseComparison);
        }

        public static bool EndsWithIgnoreCase(this string source, string value)
        {
            if (source == null)
            {
                return value == null;
            }

            if (value == null)
            {
                return false;
            }

            return source.EndsWith(value, IgnoreCaseComparison);
        }

        public static string ReplaceStartsWith(this string input, string pattern, string replacement)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(pattern))
            {
                return input;
            }

            if (input.StartsWith(pattern, IgnoreCaseComparison))
            {
                return replacement + input.Substring(pattern.Length);
            }

            return input;
        }

        public static string ReplaceEndsWith(this string input, string pattern, string replacement)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(pattern))
            {
                return input;
            }

            if (input.EndsWith(pattern, IgnoreCaseComparison))
            {
                return input.Substring(0, input.Length - pattern.Length) + replacement;
            }

            return input;
        }

        public static string[] SplitStartsWith(this string input, string delimiter)
        {
            string[] result;
            if (input.StartsWith(delimiter, IgnoreCaseComparison))
            {
                result = new[] { "", input.Substring(delimiter.Length) };
            }
            else
            {
                result = new[] { input };
            }
            return result;
        }
    }
}
