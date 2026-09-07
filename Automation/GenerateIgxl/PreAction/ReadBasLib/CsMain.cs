using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Loader;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.PreAction.ReadBasLib
{
    public class CsMain
    {
        private string _libFolderPath = "";
        private string _libFolderName = "";
        private string _outputFolderPath = "";
        private string _csProjPath = "";
        private string _dllFolderPath = "";
        private string _tempDllFolderPath = "";

#nullable enable
        public (List<Function> functions, ReferenceSheet? referenceSheet) WorkFlow(string outputPath, string csLibraryPath)
#nullable restore
        {
            _dllFolderPath = Path.Combine(outputPath, "central_library_cs", "bin");
            _tempDllFolderPath = Path.Combine(Environment.CurrentDirectory, "TempDlls");
            if (File.Exists(csLibraryPath))
            {
                if (!(string.Equals(".csproj", Path.GetExtension(csLibraryPath)) || string.Equals(".sln", Path.GetExtension(csLibraryPath))))
                {
                    return (new List<Function>(), null);
                }
                _libFolderPath = Path.GetDirectoryName(csLibraryPath);
                _libFolderName = Path.GetFileName(_libFolderPath);
                if (_libFolderName != null)
                {
                    _outputFolderPath = Path.Combine(outputPath, _libFolderName);
                }

                _csProjPath = Path.Combine(_outputFolderPath, Path.GetFileName(csLibraryPath));
                return GetFromSln();
            }

            if (Directory.Exists(csLibraryPath))
            {
                return GetFromDll(csLibraryPath);
            }

            return (new List<Function>(), null);
        }

#nullable enable
        private (List<Function> functions, ReferenceSheet? referenceSheet) GetFromSln()
#nullable restore
        {
            if (Directory.Exists(_libFolderPath))
            {
                CopyDirectory(_libFolderPath, _outputFolderPath, true);
            }

            try
            {
                ReadBasLibUtils.InitialParamMapping();
                BuildCsLibrary(_csProjPath, _dllFolderPath);
                CopyDirectory(_dllFolderPath, _tempDllFolderPath, true, true);
                return ReadLocalLib(_tempDllFolderPath);
            }
            catch (Exception e)
            {
                throw new Exception("Read bas lib failed! " + e.StackTrace);
            }
        }

#nullable enable
        private (List<Function> functions, ReferenceSheet? referenceSheet) GetFromDll(string csLibraryPath)
#nullable restore
        {
            if (Directory.Exists(csLibraryPath))
            {
                CopyDirectory(csLibraryPath, _dllFolderPath, true);
                CopyDirectory(_dllFolderPath, _tempDllFolderPath, true, true);
            }

            try
            {
                ReadBasLibUtils.InitialParamMapping();
                return ReadLocalLib(csLibraryPath);
            }
            catch (Exception e)
            {
                throw new Exception("Read bas lib failed! " + e.StackTrace);
            }
        }

        internal void CopyDirectory(string sourceDir, string destinationDir, bool recursive = false, bool ignorePdb = false)
        {
            // Get information about the source directory
            var dir = new DirectoryInfo(sourceDir);

            // Check if the source directory exists
            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException($"Source directory not found: {dir.FullName}");
            }

            // Cache directories before we start copying
            DirectoryInfo[] dirs = dir.GetDirectories();

            if (Directory.Exists(destinationDir))
            {
                Directory.Delete(destinationDir, true);
            }

            // Create the destination directory
            Directory.CreateDirectory(destinationDir);

            // Get the files in the source directory and copy to the destination directory
            foreach (FileInfo file in dir.GetFiles())
            {
                if (Path.GetExtension(file.Name).ToLower() == ".pdb" && ignorePdb)
                {
                    continue;
                }

                string targetFilePath = Path.Combine(destinationDir, file.Name);
                file.CopyTo(targetFilePath);
            }

            // If recursive and copying subdirectories, recursively call this method
            if (recursive)
            {
                foreach (DirectoryInfo subDir in dirs)
                {
                    string newDestinationDir = Path.Combine(destinationDir, subDir.Name);
                    CopyDirectory(subDir.FullName, newDestinationDir, true);
                }
            }
        }

#nullable enable
        public static (List<Function> functions, ReferenceSheet? referenceSheet) ReadLocalLib(string dllFolder)
#nullable restore
        {
            if (!Directory.Exists(dllFolder))
            {
                throw new Exception("Read C# lib failed, the C# directory not exist!");
            }

            var functionNames = new Dictionary<string, string>();
            var necessaryDlls = new List<string> { "CoreTestLibrary.dll", "Teradyne.TestAPI.dll", "PlatformCommonTestLibrary.dll", "PlatformInterfacesLibrary.dll", "PowerBinning.dll" };

            var rows = new List<string>();
            var functionBaseList = new List<Function>();

            // 1. Create a collectible AssemblyLoadContext to allow unloading later
            var alc = new AssemblyLoadContext("TempContext", isCollectible: true);

            try
            {
                string[] dllFiles = Directory.GetFiles(dllFolder, "*.dll", SearchOption.AllDirectories);
                Array.Sort(dllFiles, StringComparer.OrdinalIgnoreCase);

                // 2. Pre-load ALL target directory DLLs into the context first to resolve dependencies
                var loadedAssemblies = new Dictionary<string, Assembly>(StringComparer.OrdinalIgnoreCase);
                foreach (string dllFile in dllFiles)
                {
                    try
                    {
                        Assembly assembly = alc.LoadFromAssemblyPath(Path.GetFullPath(dllFile));
                        string fileName = Path.GetFileName(dllFile);
                        loadedAssemblies[fileName] = assembly;
                    }
                    catch
                    {
                        // Log or ignore unreadable native/corrupt DLLs
                    }
                }

                // 3. Process the loaded assemblies matching your business logic
                foreach (KeyValuePair<string, Assembly> kvp in loadedAssemblies)
                {
                    string fileName = kvp.Key;
                    Assembly assembly = kvp.Value;

                    if (!necessaryDlls.Exists(x => fileName.StartsWith(x, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        continue;
                    }

                    string referencePath = @".\central_library_cs\bin\" + fileName;
                    rows.Add(referencePath);

                    if (!fileName.StartsWith("coreTestLibrary", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    try
                    {
                        foreach (Type type in assembly.GetTypes())
                        {
                            foreach (MethodInfo method in type.GetMethods())
                            {
                                var attributes = method.CustomAttributes.ToList();
                                bool isTestMethod = attributes.Any(x => x.AttributeType.FullName == "Teradyne.Igxl.Interfaces.Public.TestMethodAttribute");
                                bool isInterposeMethod = attributes.Any(x => x.AttributeType.FullName == "Teradyne.Igxl.Interfaces.Public.InterposeMethod");

                                if (!isTestMethod && !isInterposeMethod)
                                {
                                    continue;
                                }

                                string functionName = method.Name;
                                string nameSpace = $"{method.DeclaringType?.FullName}";

                                if (!functionNames.ContainsKey(functionName))
                                {
                                    functionNames.Add(functionName, nameSpace);
                                }
                                else
                                {
                                    var nameSpaceList = new List<string> { functionNames[functionName], nameSpace };
                                    if (nameSpaceList.Exists(x => x.Contains("IgxlWrapper")))
                                    {
                                        functionNames[functionName] = nameSpaceList.Find(x => x.Contains("IgxlWrapper"));
                                    }

                                    ErrorReportManager.AddError(PreActionErrorType.E_DuplicateLibrary_01, "", 1, 0,
                                        new string[] { functionName, nameSpace, functionNames[functionName] },
                                        new ErrorInfo() { Comments = new List<string>() { nameSpace, functionNames[functionName] } });
                                }

                                var newVbt = new Function(functionName);
                                ParameterInfo[] parameters = method.GetParameters();

                                newVbt.Parameters = string.Join(",", parameters.Select(x => x.Name));
                                newVbt.ParameterDefaults = string.Join(",", parameters.Select(x => x.DefaultValue?.ToString() ?? ""));
                                newVbt.Type = ".NET";
                                newVbt.FileName = assembly.Location;
                                newVbt.NameSpace = nameSpace;
                                newVbt.IsInterPose = isInterposeMethod;

                                var allPatternParam = parameters.Where(x => x.ParameterType.FullName == "Teradyne.Igxl.Interfaces.Public.Pattern").ToList();
                                foreach (ParameterInfo arg in allPatternParam)
                                {
                                    int paramIndex = Array.IndexOf(parameters, arg);
                                    newVbt.PatternDic.Add(arg.Name, paramIndex);
                                }

                                functionBaseList.Add(newVbt);
                            }
                        }
                    }
                    catch
                    {
                        // Ignored per original logic for unreadable types
                    }
                }
            }
            finally
            {
                // 4. Safely unload the entire reflection load context from memory
                alc.Unload();
                Console.WriteLine("AssemblyLoadContext unloaded.");
            }

            var referenceSheet = new ReferenceSheet("References");
            rows.ForEach(x => referenceSheet.Rows.Add(new ReferenceRow { FilePath = x }));

            return (functionBaseList, referenceSheet);
        }

        private void BuildCsLibrary(string csProjPath, string dllFolderPath)
        {
            if (!File.Exists(csProjPath))
            {
                throw new Exception($"Can't find .csproj file in \"{csProjPath}\" !");
            }

            string visualStudioPath = GetVisualStudioInstallationPath();

            if (!string.IsNullOrEmpty(visualStudioPath))
            {
                string msBuildPath = Path.Combine(visualStudioPath, "MSBuild", "Current", "Bin", "MSBuild.exe");

                BuildProject(msBuildPath, csProjPath, dllFolderPath);

                Response.Report("Build C# library completed");
            }
            else
            {
                throw new Exception("Visual Studio not found. Please make sure it is installed.");
            }
        }

        private string GetVisualStudioInstallationPath()
        {
            string vsWherePath = Path.Combine("C:\\", "Program Files (x86)", "Microsoft Visual Studio", "Installer", "vswhere.exe");
            ProcessStartInfo processStartInfo = new ProcessStartInfo
            {
                FileName = vsWherePath,
                Arguments = "-latest -property installationPath",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = new Process())
            {
                process.StartInfo = processStartInfo;
                process.Start();

                string output = process.StandardOutput.ReadToEnd().Trim();

                process.WaitForExit();

                return output;
            }
        }

        private void BuildProject(string msBuildPath, string csProjPath, string dllFolderPath)
        {
            if (File.Exists(msBuildPath))
            {
                string msBuildArguments = $"{csProjPath} /t:Rebuild /p:Configuration=Debug;OutDir={dllFolderPath}";

                ProcessStartInfo processStartInfo = new ProcessStartInfo
                {
                    FileName = msBuildPath,
                    Arguments = msBuildArguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = new Process())
                {
                    process.StartInfo = processStartInfo;
                    process.Start();

                    process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    process.WaitForExit();

                    if (!string.IsNullOrEmpty(error))
                    {
                        throw new Exception("MSBuild Error: \n" + error);
                    }
                }
            }
            else
            {
                throw new Exception("MSBuild.exe not found. Please make sure Visual Studio is installed.");
            }
        }
    }
}
