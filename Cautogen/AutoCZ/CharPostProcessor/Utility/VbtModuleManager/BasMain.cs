using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Loader;
using System.Text.RegularExpressions;

using Cautogen.Utility;

using LogLib.Utility;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility.VbtModuleManager
{
    public class BasMain
    {
        /* property */
        public static VbtFunctionLib VbtFunctionLib = new VbtFunctionLib();

        public static VbtFunction FuntionalChar
        {
            get { return VbtFunctionLib.GetFunctionByName(VbtFunctionLib.FunctionalCharName); }
        }

        public static VbtFunction FuntionalUpdated
        {
            get { return VbtFunctionLib.GetFunctionByName(VbtFunctionLib.FunctionalName); }
        }

        /* Developing for C-Autogen for C#*/
        /* methods */
        public static void Parse(string basLibraryPath, bool generateCSharp = false)
        {
            VbtFunctionLib = new VbtFunctionLib();

            if (generateCSharp)
            {
                //var solutionDllPath = Path.Combine(basLibraryPath, "bin");
                string solutionDllPath = basLibraryPath;

                LogHelper.Info("Parsing all C# Library... ");
                VbtFunctionLib.AddVbtFunctionRange(ReadLocalLib(solutionDllPath));
            }
            else
            {
                LogHelper.Info("Parsing all vbt bas files... ");

                // reset previous result

                // check input
                if (!Directory.Exists(basLibraryPath))
                {
                    throw new Exception($"Read bas lib failed, the bas directory {basLibraryPath} not exist!");
                }

                // parese .bas files
                foreach (string basFile in FileUtility.GetFiles(basLibraryPath, "VBT_*.bas"))
                {
                    _ReadBasFile(basFile);
                }
            }

        }

        private static void _ReadBasFile(string fileName)
        {
            try
            {
                var fInfo = new FileInfo(fileName);
                var sr = new StreamReader(fInfo.FullName);
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (!Regex.IsMatch(line, @"^\s*Public\sFunction"))
                    {
                        continue;
                    }

                    string functionName = Regex.Match(line, @"Public\sFunction\s(?<func>\w+)\(").Groups["func"].ToString();
                    string paramStr = "";

                    if (Regex.IsMatch(line, @"\(.*\)"))
                    {
                        paramStr = Regex.Match(line, @"\((?<str>.*)\)").Groups["str"].ToString();
                    }
                    else
                    {
                        paramStr = Regex.Match(line, @"\((?<str>.*)").Groups["str"].ToString();
                        while ((line = sr.ReadLine()) != null && !Regex.IsMatch(line, @".*\)"))
                        {
                            if (!Regex.IsMatch(line, @"\s*\'"))
                            {
                                paramStr += line;
                            }
                        }
                        if (line != null)
                        {
                            paramStr += Regex.Match(line, @"(?<str>.*)\)").Groups["str"].ToString();
                        }
                    }
                    var paramterList =
                        Regex.Matches(paramStr, @"(?<param>\w+)\sAs\s", RegexOptions.IgnoreCase)
                            .Cast<Match>()
                            .Select(a => a.Groups["param"].ToString()).ToList();

                    VbtFunctionLib.AddVbtFunction(new VbtFunction(functionName)
                    {
                        FileName = fileName,
                        Parameters = string.Join(",", paramterList.Where(a => a.ToLower() != "step_").ToList())
                    });
                }
                sr.Close();
            }
            catch (Exception e)
            {
                throw new Exception("Read bas lib failed! " + e.Message);
            }
        }

        public static List<VbtFunction> ReadLocalLib(string dllFolder)
        {
            if (!Directory.Exists(dllFolder))
            {
                throw new DirectoryNotFoundException($"Read C# lib failed, the C# directory does not exist: {dllFolder}");
            }

            Dictionary<string, string> functionNames = new Dictionary<string, string>();
            List<string> excludeDlls = new List<string> { "Igxl.Interfaces.Public.dll", "IGXLFakes.dll" };
            var rows = new List<string>();
            var functionBaseList = new List<VbtFunction>();
            var loadedAssemblies = new List<(Assembly Assembly, string FilePath)>();

            var loadContext = new AssemblyLoadContext("TempContext", isCollectible: true);

            // FIX 1: Automatically resolve missing sub-dependencies from the same folder
            loadContext.Resolving += (context, assemblyName) =>
            {
                string expectedPath = Path.Combine(dllFolder, $"{assemblyName.Name}.dll");
                if (File.Exists(expectedPath))
                {
                    return context.LoadFromAssemblyPath(expectedPath);
                }
                return null; // Let the runtime fallback to system directories
            };

            try
            {
                string[] dllFiles = Directory.GetFiles(dllFolder, "*.dll", SearchOption.AllDirectories);

                // Phase 1: Load all valid DLLs
                foreach (string dllFile in dllFiles)
                {
                    string fileName = Path.GetFileName(dllFile);
                    if (excludeDlls.Any(x => string.Equals(x, fileName, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    try
                    {
                        Assembly assembly = loadContext.LoadFromAssemblyPath(dllFile);
                        AssemblyName[] references = assembly.GetReferencedAssemblies();

                        bool hasCorrectReferences = references.Any(x =>
                            x.Name.StartsWith("Teradyne", StringComparison.OrdinalIgnoreCase) ||
                            x.Name.StartsWith("Igxl", StringComparison.OrdinalIgnoreCase));

                        if (hasCorrectReferences)
                        {
                            loadedAssemblies.Add((assembly, dllFile));
                            string referencePath = @".\bin\" + fileName;
                            rows.Add(referencePath);
                        }
                    }
                    catch (Exception)
                    {
                        continue; // Skips non-managed or corrupted assemblies safely
                    }
                }

                // Phase 2: Extract methods safely
                foreach (var item in loadedAssemblies)
                {
                    IEnumerable<Type> types;
                    try
                    {
                        types = item.Assembly.GetTypes();
                    }
                    catch (ReflectionTypeLoadException ex)
                    {
                        // FIX 2: Extract whatever valid types DID load, and skip the broken ones!
                        types = ex.Types.Where(t => t != null);

                        // Optional Debugging: Log what specific DLL dependency was missing
                        foreach (var loaderEx in ex.LoaderExceptions)
                        {
                            Console.WriteLine($"Missing Dependency in {item.FilePath}: {loaderEx?.Message}");
                        }
                    }

                    foreach (Type type in types)
                    {
                        try
                        {
                            foreach (MethodInfo method in type.GetMethods())
                            {
                                bool hasAttribute = method.CustomAttributes.Any(x =>
                                    x.AttributeType.FullName == "Teradyne.Igxl.Interfaces.Public.TestMethodAttribute" &&
                                    x.AttributeType.Name == "TestMethodAttribute");

                                if (!hasAttribute)
                                    continue;

                                string functionName = method.Name;
                                string nameSpace = $"{method.DeclaringType?.FullName}";

                                if (!functionNames.TryGetValue(functionName, out string existingNamespace))
                                {
                                    functionNames.Add(functionName, nameSpace);
                                }
                                else
                                {
                                    string outString = $"Duplicate C# function: \"{functionName}\" in {nameSpace} != {existingNamespace}";
                                }

                                var newVbt = new VbtFunction(functionName)
                                {
                                    Parameters = string.Join(",", method.GetParameters().Select(x => x.Name)),
                                    Type = ".NET",
                                    FileName = item.FilePath,
                                    NameSpace = nameSpace
                                };

                                functionBaseList.Add(newVbt);
                            }
                        }
                        catch (Exception)
                        {
                            // Guard against reflection issues on specific problematic types
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Read C# lib failed during assembly processing.", ex);
            }
            finally
            {
                loadContext.Unload();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return functionBaseList;
        }
    }
}
