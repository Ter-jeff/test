using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using TestPlanLib.BinCut;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Static;

namespace TestPlanLib.Utility
{
    public static class BinCutInstanceRowUtility
    {
        public static bool IsConditionAllSame(BinCutInstanceRow sourceRow, BinCutInstanceRow targetRow, bool isIncludeInstance = true)
        {
            if (sourceRow.FlowName != targetRow.FlowName)
            {
                return false;
            }

            if (isIncludeInstance)
            {
                if (sourceRow.Instance != targetRow.Instance)
                {
                    return false;
                }
            }

            if (sourceRow.EnableAndDevice != targetRow.EnableAndDevice)
            {
                return false;
            }

            if (sourceRow.SubFlow != targetRow.SubFlow)
            {
                return false;
            }

            if (sourceRow.EnableFlow != targetRow.EnableFlow)
            {
                return false;
            }

            string jobTestStage1 = sourceRow.JobTestStage;
            string jobTestStage2 = targetRow.JobTestStage;
            if (jobTestStage1.EqualsIgnoreCase("All"))
            {
                jobTestStage1 = "";
            }

            if (jobTestStage2.EqualsIgnoreCase("All"))
            {
                jobTestStage2 = "";
            }

            if (jobTestStage1 != jobTestStage2)
            {
                return false;
            }

            if (sourceRow.SiteVar != targetRow.SiteVar)
            {
                return false;
            }

            if (sourceRow.FailFlag != targetRow.FailFlag)
            {
                return false;
            }

            if (sourceRow.BinOutStage != targetRow.BinOutStage)
            {
                return false;
            }

            if (sourceRow.DCcategory != targetRow.DCcategory)
            {
                return false;
            }

            if (sourceRow.TimeSet != targetRow.TimeSet)
            {
                return false;
            }

            return sourceRow.ShiftSpeed == targetRow.ShiftSpeed;
        }

        public static string GetDomain(BinCutInstanceRow binCutInstanceRow)
        {
            Domain domain = GetDomainByFlowName(binCutInstanceRow.FlowName);
            if (!string.IsNullOrEmpty(domain.Name))
            {
                return domain.Name;
            }

            if (string.IsNullOrEmpty(domain.Name))
            {
                domain = GetDomainByPattern(binCutInstanceRow.PatternList);
            }

            return domain.Name;
        }

        public static List<string> GetEnableJobs(BinCutInstanceRow binCutInstanceRow, List<string> totalJobs)
        {
            List<string> jobs = GetJobs(binCutInstanceRow.JobTestStage);
            if (totalJobs != null)
            {
                jobs = [.. jobs.Intersect(totalJobs, StringExtensions.IgnoreCase)];
            }

            return jobs;
        }

        public static List<string> GetJobs(string jobTestStage)
        {
            var jobs = new List<string>();
            foreach (string item in jobTestStage.Split(','))
            {
                if (!item.EqualsIgnoreCase("All"))
                {
                    jobs.Add(item);
                }
            }
            return jobs.ConvertAll(x => x.Trim());
        }

        public static string GetMode(string pattern)
        {
            string performanceMode = "";
            string[] subStrings = pattern.Split('_');
            if (subStrings.Length > 9 && Reg._regex4.IsMatch(subStrings[9]))
            {
                performanceMode = subStrings[9];
            }
            return performanceMode;
        }

        public static string GetTypeByFlowNameOrDcCategory(string name)
        {
            string[] type = name.Split(['_', ' '], StringSplitOptions.RemoveEmptyEntries);
            var finalType = type.Where(x => Reg._regex2.IsMatch(x)).ToList();
            return finalType.Count != 0 ? finalType[0] : "UnknowType";
        }

        public static string GetBlockByFlowName(string flowName)
        {
            string block;
            if (flowName.Contains("TD", StringComparison.OrdinalIgnoreCase) || flowName.Contains("SA", StringComparison.OrdinalIgnoreCase) || flowName.Contains("CHAIN", StringComparison.OrdinalIgnoreCase) || flowName.Contains("SSBBIST", StringComparison.OrdinalIgnoreCase) || flowName.Contains("SBST", StringComparison.OrdinalIgnoreCase))
            {
                block = "Td";
            }
            else if (flowName.Contains("BIST", StringComparison.OrdinalIgnoreCase) || flowName.Contains("BIRA", StringComparison.OrdinalIgnoreCase) || flowName.Contains("BIR", StringComparison.OrdinalIgnoreCase) || flowName.Contains("MBIST", StringComparison.OrdinalIgnoreCase))
            {
                block = "Mbist";
            }
            else if (flowName.Contains("ELB", StringComparison.OrdinalIgnoreCase) || flowName.Contains("ILB", StringComparison.OrdinalIgnoreCase))
            {
                block = "DDR";
            }
            else if (flowName.Contains("CPM", StringComparison.OrdinalIgnoreCase))
            {
                block = "Td";
            }
            else if (flowName.Contains("SCAN", StringComparison.OrdinalIgnoreCase))
            {
                block = "Scan";
            }
            else if (flowName.Contains("HTOL", StringComparison.OrdinalIgnoreCase))
            {
                block = "Htol";
            }
            else
            {
                block = "Td";
            }
            return block;
        }

        public static Domain GetDomainByFlowName(string flowName)
        {
            var domain = new Domain { Name = "" };
            if (flowName.Contains("cpu", StringComparison.OrdinalIgnoreCase))
            {
                domain.Name = "Cpu";
            }
            else if (flowName.Contains("gpu", StringComparison.OrdinalIgnoreCase))
            {
                domain.Name = "Gpu";
            }
            else if (flowName.Contains("gfx", StringComparison.OrdinalIgnoreCase))
            {
                domain.Name = "Gfx";
            }
            else if (flowName.Contains("soc", StringComparison.OrdinalIgnoreCase))
            {
                domain.Name = "Soc";
            }
            else if (flowName.Contains("spi", StringComparison.OrdinalIgnoreCase))
            {
                domain.Name = "Spi";
            }
            return domain;
        }

        public static bool IsBist(string flowName)
        {
            List<string> arr = [.. flowName.Split([' ', '#', ':'], StringSplitOptions.RemoveEmptyEntries)];
            var syntaxtList = new List<string> { "BIST", "BIR", "BIRA", "MBIST", "SRT", "SFT", "WUS", "SP3", "24N" };
            foreach (string syntax in syntaxtList)
            {
                if (arr.Count != 0)
                {
                    if (arr.Exists(x => x.EqualsIgnoreCase(syntax)))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public static Domain GetDomainByPattern(List<string> patternList)
        {
            string domain = "";
            //Organization : 'A:HARD_IP,C:CPU,L:GFX,P:HARD_IP,S:SOC,V:HARD_IP'
            //private static readonly Regex RgxOrg = new Regex(@"A|C|L|P|S|V|H", RegexOptions.IgnoreCase | RegexOptions.Compiled); //DP for Dummy Pattern
            foreach (string pattern in patternList)
            {
                List<string> arr = [.. pattern.Split('_')];
                if (arr.Count > 2)
                {
                    if (arr[2].EqualsIgnoreCase("A"))
                    {
                        domain = "HardIP";
                    }
                    else if (arr[2].EqualsIgnoreCase("P"))
                    {
                        domain = "HardIP";
                    }
                    else if (arr[2].EqualsIgnoreCase("V"))
                    {
                        domain = "HardIP";
                    }
                    else if (arr[2].EqualsIgnoreCase("H"))
                    {
                        domain = "HardIP";
                    }
                    else if (arr[2].EqualsIgnoreCase("C"))
                    {
                        domain = "Cpu";
                    }
                    else if (arr[2].EqualsIgnoreCase("L"))
                    {
                        domain = "Gfx";
                    }
                    else if (arr[2].EqualsIgnoreCase("S"))
                    {
                        domain = "Soc";
                    }
                }
                if (!string.IsNullOrEmpty(domain))
                {
                    return new Domain { Name = domain };
                }
            }
            return new Domain { Name = "" };
        }
    }
}
