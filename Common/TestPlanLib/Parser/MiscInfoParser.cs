using System.Collections.Generic;
using System.Text.RegularExpressions;

using CommonLib.Extension;

namespace TestPlanLib.Parser
{
    public class MiscInfoParser : KeyValueParser
    {
        private static readonly HashSet<string> _miscInfoRelayLists = new(StringExtensions.IgnoreCase) { "RELAYON", "RELAYOFF", "RELAY_ON", "RELAY_OFF" };
        public MiscInfoParser() : base(@"(?<key>[^:;\s]+)(?:\s*[:]\s*(?<value>(?:""[^""]*"")|[^;\s]*))?")
        {
        }

        public override Dictionary<string, string> ParseKeyValueToDictionary(string input, out string newValue)
        {
            newValue = "";
            if (string.IsNullOrEmpty(input))
            {
                return new Dictionary<string, string>(StringExtensions.IgnoreCase);
            }

            var result = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            foreach (Match match in SeparatorRegex.Matches(input))
            {
                string key = match.Groups["key"].Value.Trim();
                string value = match.Groups["value"].Success ? match.Groups["value"].Value.Trim() : "";
                if (_miscInfoRelayLists.Contains(value))
                {
                    (key, value) = (value, key);
                }

                if (result.TryGetValue(key, out string? value1))
                {
                    result[key] = $"{value1};{value.Trim('\"').Trim()}";
                }
                else
                {
                    result[key] = value.Trim('\"').Trim();
                }

                if (match.Groups["value"].Success)
                {
                    newValue += $"{key}:{value};";
                }
                else
                {
                    newValue += $"{key};";
                }
            }
            return result;
        }
    }
}
