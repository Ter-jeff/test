using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlSheets;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.PreAction.ReadBasLib
{
    public class BasMain
    {
        private Dictionary<string, string> _functionNames;
        private static readonly Regex _regex = new Regex(@"^\s*(Public\s)?Function.*\(");
        private static readonly Regex _regex2 = new Regex(@"(Public\s)?Function\s(?<func>\w+)\(");
        private static readonly Regex _regex3 = new Regex(@"\(.*\)");
        private static readonly Regex _regex4 = new Regex(@"\((?<str>.*)\)");
        private static readonly Regex _regex5 = new Regex(@"\((?<str>.*)");
        private static readonly Regex _regex6 = new Regex(@".*\)");
        private static readonly Regex _regex7 = new Regex(@"\s*\'");
        private static readonly Regex _regex8 = new Regex(@"(?<str>.*)\)");

        public (List<Function> functions, ReferenceSheet reference) WorkFlow(string outputPath)
        {
            try
            {
                ReferenceSheet reference = null;
                var functions = new List<Function>();
                _functionNames = new Dictionary<string, string>();
                if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
                {
                    var csMain = new CsMain();
                    (List<Function> functions, ReferenceSheet referenceSheet) result = csMain.WorkFlow(FolderStructure.DirIgLink, LocalSpecs.CsLibraryFolder);
                    reference = result.referenceSheet;
                    functions.AddRange(result.functions);
                }

                if (string.IsNullOrEmpty(LocalSpecs.BasLibraryFolder))
                {
                    return (functions, reference);
                }

                string libPath = LocalSpecs.BasLibraryFolder;
                var excludeFolders = new List<string> { ".svn", "PMIC", "Wireless" };

                if (LocalSpecs.Options.Device == EnumDevice.AP || LocalSpecs.Options.Device == EnumDevice.RF || LocalSpecs.Options.Device == EnumDevice.LCD)
                {
                    CopyFolderAndSubFolder(libPath, outputPath, excludeFolders);
                }
                else if (LocalSpecs.Options.Device == EnumDevice.RF)
                {
                    if (Directory.Exists(Path.Combine(libPath, "Wireless")))
                    {
                        CopyFolderAndSubFolder(Path.Combine(libPath, "Wireless"), Path.Combine(outputPath, "Wireless"));
                    }
                    else
                    {
                        CopyFolderAndSubFolder(libPath, outputPath);
                    }
                }

                ReadBasLibUtils.InitialParamMapping();
                functions.AddRange(ReadLocalLib(outputPath));
                return (functions, reference);
            }
            catch (Exception e)
            {
                throw new Exception("Read bas lib failed! " + e.StackTrace);
            }
        }

        public List<Function> ReadLib(string csLibraryFolder, string basLibraryFolder)
        {
            try
            {
                var functions = new List<Function>();
                if (!string.IsNullOrEmpty(csLibraryFolder))
                {
                    var csMain = new CsMain();
                    (List<Function> functions, ReferenceSheet referenceSheet) result = csMain.WorkFlow("", csLibraryFolder);
                    functions.AddRange(result.functions);
                }

                if (!string.IsNullOrEmpty(basLibraryFolder))
                {
                    ReadBasLibUtils.InitialParamMapping();
                    functions.AddRange(ReadLocalLib(basLibraryFolder));
                }

                return functions;
            }
            catch (Exception e)
            {
                throw new Exception("Read bas lib failed! " + e.StackTrace);
            }
        }

        public List<Function> ReadLocalLib(string dirLib)
        {
            var vbtFunctionBaseList = new List<Function>();
            if (Directory.Exists(dirLib))
            {
                string[] fileList = Directory.GetFiles(dirLib, "VBT_*");
                foreach (string basFile in fileList)
                {
                    string extension = Path.GetExtension(basFile);
                    if (string.Equals(extension, ".bas", StringComparison.OrdinalIgnoreCase))
                    {
                        vbtFunctionBaseList.AddRange(ReadBasFile(basFile));
                    }
                }

                foreach (string sub in Directory.GetDirectories(dirLib))
                {
                    string[] subList;
                    if (sub.IndexOf("Wireless", StringComparison.OrdinalIgnoreCase) > 0 ||
                        sub.IndexOf("PMIC", StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        subList = Directory.GetFiles(sub, "VBT_*", SearchOption.AllDirectories);
                    }
                    else
                    {
                        subList = Directory.GetFiles(sub, "VBT_*");
                    }

                    foreach (string basFile in subList)
                    {
                        string extension = Path.GetExtension(basFile);
                        if (extension != null && string.Equals(extension, ".bas", StringComparison.OrdinalIgnoreCase))
                        {
                            vbtFunctionBaseList.AddRange(ReadBasFile(basFile));
                        }
                    }
                }
            }
            else
            {
                throw new Exception("Read bas lib failed, the bas directory not exist!");
            }
            return vbtFunctionBaseList;
        }

        private void CopyFolderAndSubFolder(string sourcePath, string targetPath, List<string> excludeFolders = null)
        {
            if (!Directory.Exists(targetPath))
            {
                Directory.CreateDirectory(targetPath);
            }

            FileCopy(sourcePath, targetPath);
            foreach (string sub1 in Directory.GetDirectories(sourcePath))
            {
                if (LocalSpecs.Options.Device == EnumDevice.RF)
                {
                    if (Path.GetFileName(sub1).Equals("efuse", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!sub1.ContainsIgnoreCase("wireless"))
                        {
                            continue;
                        }
                    }
                }
                if (excludeFolders != null && excludeFolders.Exists(x => x.Equals(Path.GetFileName(sub1), StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                string newFolder1 = Path.Combine(targetPath, Path.GetFileName(sub1));
                if (!Directory.Exists(newFolder1))
                {
                    Directory.CreateDirectory(newFolder1);
                }

                FileCopy(sub1, newFolder1);
            }
        }

        private void FileCopy(string sourcePath, string targetPath)
        {
            foreach (string file in Directory.GetFiles(sourcePath))
            {
                string extension = Path.GetExtension(file);
                if (extension.ToLower() == ".bas" ||
                    extension.ToLower() == ".cls" ||
                    extension.ToLower() == ".frm" ||
                    extension.ToLower() == ".frx")
                {

                    if (LocalSpecs.Options.Device == EnumDevice.RF || LocalSpecs.Options.Device == EnumDevice.LCD)
                    {
                        if (!Path.GetFileNameWithoutExtension(file).EndsWith("_AP", StringComparison.OrdinalIgnoreCase))
                        {
                            File.Copy(file, Path.Combine(targetPath, Path.GetFileName(file)), true);
                            if (LocalSpecs.Options.Device == EnumDevice.LCD)
                            {
                                AddComment(Path.Combine(targetPath, Path.GetFileName(file)));
                            }
                        }
                    }
                    else
                    {
                        File.Copy(file, Path.Combine(targetPath, Path.GetFileName(file)), true);
                    }
                }
            }
        }

        private List<Function> ReadBasFile(string fileName)
        {
            var vbtFunctionBaseList = new List<Function>();
            var fInfo = new FileInfo(fileName);
            var sr = new StreamReader(fInfo.FullName);
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                if (!_regex.IsMatch(line))
                {
                    continue;
                }

                string functionName = _regex2.Match(line).Groups["func"].ToString();
                string paramStr;
                line = line.TrimEnd('_');
                if (_regex3.IsMatch(line))
                {
                    paramStr = _regex4.Match(line).Groups["str"].ToString();
                }
                else
                {
                    paramStr = _regex5.Match(line).Groups["str"].ToString();
                    while ((line = sr.ReadLine()) != null && !_regex6.IsMatch(line))
                    {
                        line = line.TrimEnd('_');
                        if (!_regex7.IsMatch(line))
                        {
                            paramStr += line;
                        }
                    }
                    if (line != null)
                    {
                        paramStr += _regex8.Match(line).Groups["str"].ToString();
                    }
                }
                List<Parameter> parameters = GetParameters(paramStr);
                if (_functionNames != null)
                {
                    if (!_functionNames.ContainsKey(functionName))
                    {
                        _functionNames.Add(functionName, fileName);
                    }
                    else
                    {
                        ErrorReportManager.AddError(PreActionErrorType.E_DuplicateLibrary_01, "", 1, 0,
                            new string[] { functionName, fileName, _functionNames[functionName] },
                            new ErrorInfo() { Comments = new List<string>() { fileName, _functionNames[functionName] } });
                    }
                }

                var newVbt = new Function(functionName) { FileName = fileName };

                for (int a = 0; a < parameters.Count; a++)
                {
                    if (string.Equals(parameters[a].Name, "step_", StringComparison.OrdinalIgnoreCase))
                    {
                        parameters.RemoveAt(a);
                        break;
                    }
                }

                newVbt.Parameters = string.Join(",", parameters.Select(x => x.Name));
                newVbt.ParameterDefaults = string.Join(",", parameters.Select(x => x.Default));
                vbtFunctionBaseList.Add(newVbt);
            }
            sr.Close();
            return vbtFunctionBaseList;
        }

        private List<Parameter> GetParameters(string paramStr)
        {
            if (string.IsNullOrEmpty(paramStr))
            {
                return new List<Parameter>();
            }

            var parameters = new List<Parameter>();
            foreach (string str in paramStr.Split(','))
            {
                var parameter = new Parameter();
                string parameterName = Regex.Match(str, @"(?<param>\w+)\sAs\s", RegexOptions.IgnoreCase).Groups["param"].ToString();
                string type = Regex.Match(paramStr, @"(?<param>\w+)\sAs\s(?<type>[^,]*)", RegexOptions.IgnoreCase).Groups["type"].ToString();
                parameter.Name = parameterName;
                if (type.Contains("="))
                {
                    parameter.Type = type.Replace(" ", "").Split('=')[0].Replace("\"", "");
                    parameter.Default = type.Replace(" ", "").Split('=')[1].Replace("\"", "");
                }
                else
                {
                    parameter.Type = type;
                    parameter.Default = "";
                }

                parameters.Add(parameter);
            }

            return parameters;
        }

        public void AddComment(string filePath)
        {
            if (File.Exists(filePath))
            {
                List<string> lines = ReadBasContent(filePath, null);
                File.WriteAllLines(filePath, lines);
            }
        }

        public static List<string> ReadBasContent(string fileName, List<Proc> procs)
        {
            var list = new List<string>();
            int cnt = 0;
            using (var sw = new StreamReader(fileName))
            {
                string line;
                while ((line = sw.ReadLine()) != null)
                {
                    cnt++;
                    bool isSkip = false;
                    if (procs != null)
                    {
                        foreach (Proc skipLine in procs)
                        {
                            if (cnt >= skipLine.Start && cnt < skipLine.End)
                            {
                                isSkip = true;
                                break;
                            }
                        }
                    }

                    if (!isSkip)
                    {
                        list.Add(line);
                    }
                }
            }
            return list;
        }
    }

    public class Parameter
    {
        public string Name;
        public string Default;
        public string Type;
    }

    public class Proc
    {
        public string Name;
        public int Start;
        public int End;
        public string Type;
        public List<Parameter> Parameters;
    }
}
