using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;

using Automation.Static;

using CommonLib.Static;

namespace Automation.GenerateIgxl.PostAction.VersionTracing
{
    /// <summary>
    /// Table where each row represents a key-value pair of some metadata
    /// </summary>
    public class ExecInfoTable
    {
        private readonly string _workingDir;
        private readonly Dictionary<string, string> _dllMetadata;
        private readonly List<string[]> _table;

        private ExecInfoTable()
        {
            // Assume that this is within the documents repo
            _workingDir = FolderStructure.DirReference;

            _table = new List<string[]>();

            // The other git metadata is stored in this assembly
            Assembly thisAssembly = Assembly.GetAssembly(GetType());
            _dllMetadata = GetAssemblyMetadata(thisAssembly);

            AddAutogenMetadata();
            AddDocumentMetadata();
        }

        public static List<string[]> CreateExecInfoTable()
        {
            var execInfoTable = new ExecInfoTable();
            return execInfoTable._table;
        }

        private void AddRow(string key, string value, bool skipIfEmpty = true)
        {
            if (skipIfEmpty && string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            _table.Add(new[] { key, value });
        }

        #region Autogen Metadata

        /// <summary>
        /// Get all the <see cref="AssemblyMetadataAttribute"/> embedded in the given DLL
        /// </summary>
        private static Dictionary<string, string> GetAssemblyMetadata(Assembly assembly)
        {
            // Cannot use LINQ ToDictionary because it is not supported by IGXL red button (production) runs
            Dictionary<string, string> metadata = new Dictionary<string, string>();
            foreach (AssemblyMetadataAttribute attribute in assembly.GetCustomAttributes<AssemblyMetadataAttribute>())
            {
                if (!metadata.ContainsKey(attribute.Key))
                {
                    metadata.Add(attribute.Key, attribute.Value);
                }
            }
            return metadata;
        }

        private string GetMetadata(params string[] keys)
        {
            foreach (string key in keys)
            {
                if (_dllMetadata.TryGetValue(key, out string value)
                    && !string.IsNullOrWhiteSpace(value)
                    && !value.Contains("command not found")
                    && !value.Contains("not recognized as an internal or external command"))
                {
                    return value.Trim();
                }
            }

            return null;
        }

        private static string GetCurrentTimestamp()
        {
            DateTime now = TimeContext.Now;
            string dateString = now.ToString("yyyy-MM-dd HH:mm:ss");
            string timeZoneOffset = now.ToString("zzz").Replace(":", "");

            TimeZoneInfo localTimeZone = TimeZoneInfo.Local;
            string timeZoneName = now.IsDaylightSavingTime() ? localTimeZone.DaylightName : localTimeZone.StandardName;

            return $"{dateString} {timeZoneOffset} ({timeZoneName})";
        }

        private void AddAutogenMetadata()
        {
            // The version string is stored in AutogenCommandLine
            Assembly cliAssembly = Assembly.GetEntryAssembly() ?? GetType().Assembly;
            AssemblyName cliAssemblyName = cliAssembly.GetName();
            FileVersionInfo cliVersionInfo = FileVersionInfo.GetVersionInfo(cliAssembly.Location);
            string cliVersionString = !string.IsNullOrEmpty(cliVersionInfo.ProductVersion)
                ? cliVersionInfo.ProductVersion : cliAssemblyName.Version.ToString();
            AddRow("Autogen Version", cliVersionString);
            AddRow("Autogen Assembly", cliAssemblyName.FullName);

            AddRow("Autogen Build Host", GetMetadata("BuildHost"));
            AddRow("Autogen Build Date", GetMetadata("BuildDate"));
            AddRow("Autogen Build Configuration", GetMetadata("BuildConfiguration"));
            AddRow("Autogen Build .NET SDK", GetMetadata("DotnetSdk"));

            AddRow("Autogen Runtime .NET", RuntimeInformation.FrameworkDescription);
            AddRow("Autogen Runtime Host", Environment.MachineName);
            AddRow("Autogen Runtime Date", GetCurrentTimestamp());

            AddRow("Autogen Git Branch", GetMetadata("JenkinsMergeRequestBranch", "JenkinsGitBranch", "GitLocalBranch"));
            AddRow("Autogen Git Commit Hash", GetMetadata("JenkinsMergeRequestHash", "JenkinsGitBranchHash", "GitCommitHash"));
            AddRow("Autogen Git Tag", GetMetadata("GitCommitTag"));
            AddRow("Autogen Git Author Date", GetMetadata("GitAuthorDate"));
            AddRow("Autogen Git Commit Date", GetMetadata("GitCommitterDate"));

            AddRow("Autogen Pipeline URL", GetMetadata("JenkinsBuildUrl"));
            AddRow("Autogen Pipeline Build Number", GetMetadata("JenkinsBuildNumber"));
        }

        #endregion

        #region Documents Metadata

        private string RunGitCommand(string args)
        {
            // Assume that Autogen is being run from within the documents repo
            var processStartInfo = new ProcessStartInfo
            {
                FileName = "git",
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WorkingDirectory = _workingDir
            };
            using (var process = new Process())
            {
                process.StartInfo = processStartInfo;
                process.Start();

                // Only look at stdout, not stderr
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                return output.Trim();
            }
        }

        private static string GetEnvVar(params string[] envVarNames)
        {
            foreach (string envVarName in envVarNames)
            {
                string value = Environment.GetEnvironmentVariable(envVarName);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }
            return null;
        }

        private void AddDocumentMetadata()
        {
            AddRow("Document Git Branch", GetEnvVar("CHANGE_BRANCH", "BRANCH_NAME")
                ?? RunGitCommand("rev-parse --abbrev-ref HEAD"));
            AddRow("Document Git Commit Hash", GetEnvVar("GITLAB_OA_LAST_COMMIT_ID", "GITLAB_CHECKOUT_SHA")
                ?? RunGitCommand("describe --long --always --dirty --abbrev=40"));
            AddRow("Document Git Author Date", RunGitCommand("show -s --format=%ai HEAD"));
            AddRow("Document Git Commit Date", RunGitCommand("show -s --format=%ci HEAD"));

            AddRow("Document Pipeline URL", GetEnvVar("BUILD_URL"));
            AddRow("Document Pipeline Build Number", GetEnvVar("BUILD_NUMBER"));
        }

        #endregion
    }
}
