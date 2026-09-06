using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Basic.Business.GenPatSet.Business;

namespace Automation.Reader
{
    public class PatInfoCmdReader
    {
        public void Read(ref PatternResult patternResult, List<string> list)
        {
            foreach (string line in list)
            {
                string context = line.Trim('\r').Trim();

                if (context == "")
                {
                    continue;
                }

                if (Regex.IsMatch(line, "Number of pattern vectors:", RegexOptions.IgnoreCase))
                {
                    patternResult.PatternVectorCount = Convert.ToInt32(context.Split(':')[1]);
                }
                else if (Regex.IsMatch(context, "Module names:", RegexOptions.IgnoreCase))
                {
                    patternResult.ModuleNameList =
                        context.Split(':')[1].Split(',')
                            .Select(x => x.Trim())
                            .Where(x => !string.IsNullOrEmpty(x))
                            .ToList();
                    patternResult.VmVectorName = string.Join(",", patternResult.ModuleNameList);
                }
                else if (Regex.IsMatch(context, "Number of SVM Chan Locs", RegexOptions.IgnoreCase))
                {
                    patternResult.NumOfSvm = Convert.ToInt32(context.Split(':')[1]);
                }
                else if (Regex.IsMatch(context, "Number of LVM Chan Locs", RegexOptions.IgnoreCase))
                {
                    patternResult.NumOfLvm = Convert.ToInt32(context.Split(':')[1]);
                    break;
                }
            }

            int lineIndex = list.FindIndex(x => Regex.IsMatch(x, "opcode_mode", RegexOptions.IgnoreCase));
            if (lineIndex != -1)
            {
                patternResult.OpcodeMode =
                    Regex.Match(list[lineIndex], @".*opcode_mode\s+switch\s+STRING\s+(?<opcode>\w+)$",
                        RegexOptions.IgnoreCase).Groups["opcode"].ToString();
                int index = list[lineIndex].LastIndexOf("STRING");
                patternResult.OpcodeMode = list[lineIndex].Substring(index + 6).Trim();
                if (patternResult.OpcodeMode.Equals("SINGLE", StringComparison.CurrentCultureIgnoreCase))
                {
                    patternResult.LicenseCount = GetNoRemainderBy128(patternResult.NumOfLvm * 4);
                }
                else if (patternResult.OpcodeMode.Equals("DUAL", StringComparison.CurrentCultureIgnoreCase))
                {
                    patternResult.LicenseCount = GetNoRemainderBy128(patternResult.NumOfLvm * 2);
                }
                else if (patternResult.OpcodeMode.Equals("QUAD", StringComparison.CurrentCultureIgnoreCase))
                {
                    patternResult.LicenseCount = GetNoRemainderBy128(patternResult.NumOfLvm);
                }
            }
        }

        private int GetNoRemainderBy128(int count)
        {
            int remainder = count % 128;
            return remainder == 0 ? count : ((count / 128) + 1) * 128;
        }
    }
}
