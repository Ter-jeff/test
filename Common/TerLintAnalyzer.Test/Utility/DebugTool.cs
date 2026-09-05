using System;
using System.Text;

namespace TerLintAnalyzer.Test.Utility
{
    internal static class DebugTool
    {
        public static string GetText(string source, int startLine, int startCol, int endLine, int endCol)
        {
            // Split into lines using all common newline types
            string[] lines = source.Split(["\r\n", "\r", "\n"], StringSplitOptions.None);

            // Convert 1-based Roslyn coordinates to 0-based array indices
            int startLineIdx = startLine - 1;
            int endLineIdx = endLine - 1;
            int startColIdx = startCol - 1;
            int endColIdx = endCol - 1;

            if (startLineIdx == endLineIdx)
            {
                // Single line extraction
                return lines[startLineIdx][startColIdx..endColIdx];
            }

            // Multi-line extraction
            var result = new StringBuilder();
            for (int i = startLineIdx; i <= endLineIdx; i++)
            {
                if (i == startLineIdx)
                {
                    result.Append(lines[i][startColIdx..]);
                }
                else if (i == endLineIdx)
                {
                    result.Append(lines[i][..endColIdx]);
                }
                else
                {
                    result.Append(lines[i]);
                }

                if (i != endLineIdx)
                {
                    result.AppendLine();
                }
            }
            return result.ToString();
        }
    }
}
