using System.Collections.Generic;
using System.Text.RegularExpressions;

using CommonLib.Extension;

namespace TestPlanLib.Parser
{
    public class KeyValueParser
    {
        protected readonly Regex SeparatorRegex;
        protected readonly string RegexPattern = @"(?<key>[^=:,;\s]+)\s*[=:]\s*(?<value>(?:""[^""]*"")|[^,;\s]+)|(?<key>\w+)";

        public KeyValueParser(string regexPattern)
        {
            if (regexPattern != null)
            {
                RegexPattern = regexPattern;
            }

            SeparatorRegex = SetSeparatorRegex();
        }

        private Regex SetSeparatorRegex()
        {
            return new Regex(RegexPattern, RegexOptions.Compiled);
        }

        public virtual Dictionary<string, string> ParseKeyValueToDictionary(string input, out string newValue)
        {
            newValue = input;
            if (string.IsNullOrEmpty(input))
            {
                return new Dictionary<string, string>(StringExtensions.IgnoreCase);
            }

            var result = new Dictionary<string, string>(StringExtensions.IgnoreCase);

            foreach (Match match in SeparatorRegex.Matches(input))
            {
                string key = match.Groups["key"].Value.Trim();
                string value = match.Groups["value"].Success ? match.Groups["value"].Value.Trim() : key;

                if (result.TryGetValue(key, out string? value1))
                {
                    result[key] = $"{value1};{value}";
                }
                else
                {
                    result[key] = value;
                }
            }
            return result;
        }
    }
}
