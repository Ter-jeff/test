using System.Collections.Generic;
using System.IO;

namespace ProjectConfigLib.ProjectConfig
{
    public class ProjectConfigRow
    {
        public string Group = "";
        public string Name = "";
        public string Value = "";
    }

    public static class ProjectConfigIniReader
    {
        public static List<ProjectConfigRow> ReadFile(string filePath)
        {
            IEnumerable<string> lines = File.ReadLines(filePath);
            return ReadProjectConfigLines(lines);
        }

        private static List<ProjectConfigRow> ReadProjectConfigLines(IEnumerable<string> lines)
        {
            List<ProjectConfigRow> rows = [];
            string currentGroup = "";

            foreach (string rawLine in lines)
            {
                if (string.IsNullOrEmpty(rawLine))
                {
                    continue;
                }

                string line = rawLine.Trim();

                if (line.Length == 0)
                {
                    continue;
                }

                if (line.StartsWith(';') ||
                    line.StartsWith('#'))
                {
                    continue;
                }

                if (line.StartsWith('[') &&
                    line.EndsWith(']'))
                {
                    currentGroup = line[1..^1].Trim();
                    continue;
                }

                int eqIndex = line.IndexOf('=');
                if (eqIndex < 0)
                {
                    continue;
                }

                string name = line[..eqIndex].Trim();
                string value = line[(eqIndex + 1)..].Trim();

                ProjectConfigRow row = new ProjectConfigRow
                {
                    Group = currentGroup,
                    Name = name,
                    Value = value
                };

                rows.Add(row);
            }

            return rows;
        }
    }

}
