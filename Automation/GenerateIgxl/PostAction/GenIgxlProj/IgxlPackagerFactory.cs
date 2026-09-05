using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Automation.GenerateIgxl.PostAction.GenIgxlProj
{
    public static class IgxlPackagerFactory
    {
        private const string PackagerExeName = "Automation.IgxlPackaging.exe";
        private const string PackagerSubfolder = "IgxlPackager";
        private const string PackagerPathEnvVar = "IGXL_PACKAGER_PATH";
        private const string MonoPathEnvVar = "MONO_PATH";

        /// <summary>
        /// Returns a packager that can build .igxlProj files for this run.
        /// Falls back to <see cref="NoOpIgxlPackager"/> if the side-car is missing,
        /// (on non-Windows) Mono cannot be located, or <paramref name="skipIgLink"/>
        /// is true (i.e. the user passed <c>--skip-iglink</c>).
        /// </summary>
        public static IIgxlPackager Create(bool skipIgLink = false)
        {
            if (skipIgLink)
            {
                return new NoOpIgxlPackager("--skip-iglink was specified");
            }

            string exePath = LocateSidecarExe();
            if (exePath == null)
            {
                return new NoOpIgxlPackager($"{PackagerExeName} not found next to entry assembly or via {PackagerPathEnvVar}");
            }

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                return new IgxlPackagerProcessLauncher(exePath, monoExe: null);
            }

            string monoExe = LocateMono();
            if (monoExe == null)
            {
                return new NoOpIgxlPackager("mono not found in PATH; install Mono or set MONO_PATH to run IGXL packaging on this OS");
            }

            return new IgxlPackagerProcessLauncher(exePath, monoExe);
        }

        internal static string LocateSidecarExe()
        {
            string explicitPath = Environment.GetEnvironmentVariable(PackagerPathEnvVar);
            if (!string.IsNullOrEmpty(explicitPath) && File.Exists(explicitPath))
            {
                return explicitPath;
            }

            foreach (string baseDir in CandidateBaseDirectories())
            {
                if (string.IsNullOrEmpty(baseDir))
                {
                    continue;
                }
                // Prefer the isolated subfolder (the build target drops it there
                // so the net48 DLLs can't collide with the .NET 8 host's).
                string subfolderCandidate = Path.Combine(baseDir, PackagerSubfolder, PackagerExeName);
                if (File.Exists(subfolderCandidate))
                {
                    return subfolderCandidate;
                }
                string flatCandidate = Path.Combine(baseDir, PackagerExeName);
                if (File.Exists(flatCandidate))
                {
                    return flatCandidate;
                }
            }

            return null;
        }

        internal static IEnumerable<string> CandidateBaseDirectories()
        {
            yield return Path.GetDirectoryName(Assembly.GetEntryAssembly()?.Location ?? string.Empty);
            yield return Path.GetDirectoryName(typeof(IgxlPackagerFactory).Assembly.Location ?? string.Empty);
            yield return AppContext.BaseDirectory;
        }

        internal static string LocateMono()
        {
            string explicitMono = Environment.GetEnvironmentVariable(MonoPathEnvVar);
            if (!string.IsNullOrEmpty(explicitMono) && File.Exists(explicitMono))
            {
                return explicitMono;
            }

            string pathEnv = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            foreach (string dir in pathEnv.Split(Path.PathSeparator))
            {
                if (string.IsNullOrEmpty(dir))
                {
                    continue;
                }
                string candidate = Path.Combine(dir, "mono");
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
            return null;
        }
    }
}
