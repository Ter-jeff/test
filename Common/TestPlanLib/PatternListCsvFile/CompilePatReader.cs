using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace TestPlanLib.PatternListCsvFile
{
    public partial class CompilePatReader
    {
        [GeneratedRegex(".*Product.*Compilation.*", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        private string _compileFile = "";
        private readonly Dictionary<string, CompileItem> _compileList = [];

        public Dictionary<string, CompileItem> ReadCompileFile(string filePath)
        {
            _compileFile = filePath;

            int idxProduct = 0;
            int idxVersion = 0;
            int idxTpCategory = 0;
            int idxAtpName = 0;
            int idxOpCode = 0;
            int idxScanMode = 0;
            int idxHalt = 0;
            int idxCompilation = 0;
            int idxMd5 = 0;
            int idxHLv = 0;
            int idxScanSetupTSet = 0;
            var reader = new StreamReader(File.OpenRead(_compileFile));
            bool dataStart = false;
            while (!reader.EndOfStream)
            {
                string? line = reader.ReadLine();
                if (line != null)
                {
                    string[] values = line.Split(',');
                    if (dataStart)
                    {
                        if (values[idxAtpName].Length > 10) //key
                        {
                            var compItem = new CompileItem
                            {
                                Product = values[idxProduct],
                                Version = values[idxVersion],
                                TpCategory = values[idxTpCategory],
                                AtpName = values[idxAtpName],
                                OpCode = values[idxOpCode],
                                ScanMode = values[idxScanMode],
                                Halt = values[idxHalt],
                                Compilation = values[idxCompilation],
                                Md5 = values[idxMd5],
                                HLv = values[idxHLv],
                                ScanSetupTSet = values[idxScanSetupTSet]
                            };
                            _compileList.TryAdd(values[idxAtpName], compItem);
                        }
                    }
                    if (_regex.IsMatch(line)) //the 1st line
                    {
                        for (int item = 0; item < values.Length; item++)
                        {
                            if (values[item] == "Product")
                            {
                                idxProduct = item;
                            }

                            if (values[item] == "Version")
                            {
                                idxVersion = item;
                            }

                            if (values[item] == "T/P Category")
                            {
                                idxTpCategory = item;
                            }

                            if (values[item] == "AtpName")
                            {
                                idxAtpName = item;
                            }

                            if (values[item] == "OpCode")
                            {
                                idxOpCode = item;
                            }

                            if (values[item] == "ScanMode")
                            {
                                idxScanMode = item;
                            }

                            if (values[item] == "Halt")
                            {
                                idxHalt = item;
                            }

                            if (values[item] == "Compilation")
                            {
                                idxCompilation = item;
                            }

                            if (values[item] == "MD5")
                            {
                                idxMd5 = item;
                            }

                            if (values[item] == "HLV")
                            {
                                idxHLv = item;
                            }

                            if (values[item] == "ScanSetupTSet")
                            {
                                idxScanSetupTSet = item;
                            }
                        }
                        dataStart = true;
                    }
                }
            }
            reader.Close();
            return _compileList;
        }

        public Dictionary<string, string> GetLatestPatDict()
        {
            var lGenericPatternGroup = new Dictionary<string, List<string>>();
            foreach (KeyValuePair<string, CompileItem> compiledItem in _compileList)
            {
                var patGeneric = new PatternNameInfo(compiledItem.Key);

                if (!lGenericPatternGroup.TryGetValue(patGeneric.GenericName, out List<string>? value))
                {
                    value = [];
                    lGenericPatternGroup.Add(patGeneric.GenericName, value);
                }

                value.Add(patGeneric.PatternVersion + "_" + patGeneric.SiliconVersion + "_" + patGeneric.TimeStamp);
            }

            var lLatestPatDict = new Dictionary<string, string>();
            foreach (KeyValuePair<string, List<string>> patGroup in lGenericPatternGroup)
            {
                List<string> allVersions = patGroup.Value;
                string latestVersion = "";
                foreach (string patternVersion in allVersions)
                {
                    if (latestVersion.Length == 0)
                    {
                        latestVersion = patternVersion;
                        continue;
                    }
                    latestVersion = CompareVersion(latestVersion, patternVersion);
                }
                lLatestPatDict.Add(patGroup.Key, latestVersion);
            }
            return lLatestPatDict;
        }

        private static string CompareVersion(string version1, string version2)
        {
            // Version1:  "1_A0_1510070021"  =>1
            // Version2:  "2_A0_1510070021"  =>2
            // 2 > 1 return Version 2

            int i = Convert.ToInt16(version1.Split('_')[0]);
            int j = Convert.ToInt16(version2.Split('_')[0]);
            if (i > j)
            {
                return version1;
            }

            return version2;
        }
    }
}
