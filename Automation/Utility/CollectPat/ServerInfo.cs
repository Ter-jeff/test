using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Automation.Utility.CollectPat
{
    public class SubrPatInfo
    {
        public List<string> Subroutine { get; set; } = new List<string>();
        public string VmVector { get; set; }
        public string SubroutineCnt { get; set; }
    }

    public static class ServerInfo
    {
        public static Dictionary<string, SubrPatInfo> ReadHardIpInfoAll(string hardipInfoFile)
        {
            var patternInfoAll = new Dictionary<string, SubrPatInfo>();

            if (!File.Exists(hardipInfoFile))
            {
                return patternInfoAll;
            }

            string line;
            string genericPatName = "";
            string version = "";
            var subrPatInfo = new SubrPatInfo();
            var file = new StreamReader(hardipInfoFile);
            while ((line = file.ReadLine()) != null)
            {
                if (line.StartsWith("GenericPat:=", StringComparison.CurrentCultureIgnoreCase))
                {
                    genericPatName = line.Substring(line.IndexOf(":=", StringComparison.Ordinal) + 2);
                }
                else if (line.StartsWith("Version:=", StringComparison.CurrentCultureIgnoreCase))
                {
                    version = line.Substring(line.IndexOf(":=", StringComparison.Ordinal) + 2);
                }
                else if (line.StartsWith("<HardIP_Info_Token>", StringComparison.CurrentCultureIgnoreCase))
                {
                    // new Pat
                    string name = genericPatName + "_" + version;
                    if (!patternInfoAll.ContainsKey(name.ToUpper()))
                    {
                        patternInfoAll.Add(name.ToUpper(), subrPatInfo);
                    }

                    // clean up variables
                    genericPatName = "";
                    version = "";
                    subrPatInfo = new SubrPatInfo();

                }
                else if (line != "")
                {
                    if (line.StartsWith("Subr:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        string[] subList = line.Split(':')[1].Trim().Split(',');
                        foreach (string sub in subList)
                        {
                            subrPatInfo.Subroutine.Add(sub);
                        }
                    }

                    if (line.StartsWith("VM_Vector:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        // "VM_Vector: A-Z | a-z | 0-9 | _
                        var m1 = new Regex(@"(VM_Vector:)((\s)+)(?<VM_vector>(([A-Z]|[a-z]|[0-9]|[_])+))");
                        Match match = m1.Match(line);
                        if (match.Success)
                        {
                            subrPatInfo.VmVector = match.Groups["VM_vector"].Value;
                        }
                    }
                }
                if (line.StartsWith("Call_Subrs_Cnt_Diff", StringComparison.CurrentCultureIgnoreCase))
                {
                    string arr = line.Split(':')[1];
                    subrPatInfo.SubroutineCnt = arr;
                }
            }
            file.Close();
            return patternInfoAll;
        }
    }
}
