using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Reader.ScghFile.ProCharPatternSet.Business;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    public class PatSetSheetGenerator
    {
        public HardIpInputData HardIpInputData { get; }

        public PatSetSheetGenerator(HardIpInputData hardIpInputData)
        {
            HardIpInputData = hardIpInputData;
        }

        public PatSetSheet GenPatSet(Dictionary<string, HardIpSheet> planDic, ScghData scghData, string sheetName = "PatSets_HardIP")
        {
            //const string sheetName = @"PatSets_HardIP";
            PatSetSheet existPatsetSheet = TestProgram.IgxlWorkBk.GetPatSetsSheet(sheetName, FolderStructure.DirPatSetsAll);
            PatSetSheet newSheet = existPatsetSheet.Rows.Any() ? existPatsetSheet : new PatSetSheet(sheetName);
            HardIpPatSetConstructor hardIpPatset = new HardIpPatSetConstructor(HardIpInputData);

            var pllFreqSet = new List<HardIpPattern>();
            try
            {
                foreach (KeyValuePair<string, HardIpSheet> keyValuePair in planDic)
                {
                    hardIpPatset.SheetName = keyValuePair.Key;
                    List<HardIpPattern> hardIpPatterns = keyValuePair.Value.Rows;
                    foreach (HardIpPattern pattern in hardIpPatterns)
                    {
                        string blockName = CommonGenerator.GetBlockName(pattern);
                        if (string.IsNullOrEmpty(pattern.BlockType))
                        {
                            #region General HardIP Rule

                            if (pattern.Pattern.IsMultiTimeDomain())
                            {
                                foreach (List<string> pats in pattern.Pattern.PatternSetList)
                                {
                                    if (pats.Count > 1)
                                    {
                                        string newPatSetName = hardIpPatset.GenPatSetName(
                                            scghData, pats, pattern.MiscInfo, blockName, newSheet);
                                        pattern.Pattern.InstancePatternName.Add(newPatSetName);
                                        pattern.Pattern.InstancePayloadName.Add(newPatSetName);
                                    }
                                    else
                                    {
                                        pattern.Pattern.InstancePayloadName.Add(pats[0]);
                                    }
                                }
                            }
                            else if (pattern.Pattern.IsMultiple())
                            {
                                pattern.Pattern.InstancePatternName = new List<string>();
                                var list = pattern.Pattern.PatternSetList.SelectMany(p => p).ToList();
                                if (list.Count > 1)
                                {
                                    string newPatSetName = hardIpPatset.GenPatSetName(
                                        scghData, list, pattern.MiscInfo, blockName, newSheet);
                                    pattern.Pattern.InstancePatternName.Add(newPatSetName);
                                    pattern.Pattern.InstancePayloadName.Add(newPatSetName);
                                }
                                else
                                {
                                    pattern.Pattern.InstancePayloadName.Add(list[0]);
                                }
                            }

                            #endregion
                        }
                        else
                        {
                            #region non HardIP block

                            pllFreqSet.Add(pattern);
                            if (pattern.MeasPins.Count > 0)
                            {
                                string regDSelSrm = "_[D]*SRA*M";
                                HardIpPattern burstPattern =
                                    pllFreqSet.FirstOrDefault(p =>
                                        Regex.IsMatch(p.Pattern.GetLastPayload(), regDSelSrm,
                                            RegexOptions.IgnoreCase)) ?? pllFreqSet[0];

                                burstPattern.Pattern.InstancePatternName = new List<string>();

                                PatSet newPatSet = new PatSet
                                {
                                    PatSetName = GenPllPatSetName(pllFreqSet, burstPattern)
                                };
                                burstPattern.Pattern.InstancePatternName.Add(newPatSet.PatSetName);
                                burstPattern.Pattern.InstancePayloadName.Add(newPatSet.PatSetName);
                                burstPattern.TestName = newPatSet.PatSetName;
                                foreach (string pat in pllFreqSet.Select(p => p.Pattern.GetLastPayload()))
                                {
                                    if (pllFreqSet.Last().Pattern.GetLastPayload() == pat)
                                    {
                                        continue;
                                    }

                                    var patSet = new PatSetRow { Burst = "No", File = pat };
                                    newPatSet.AddRow(patSet);
                                }

                                AddPatSet(newSheet, newPatSet);
                                pllFreqSet.Clear();
                            }

                            #endregion
                        }

                    }
                }

                if (hardIpPatset.MissingPatSets.Count > 0)
                {
                    var patSet = new PatSetRow { Burst = "", File = "" };
                    PatSet newPatSet = new PatSet { PatSetName = "" };
                    newPatSet.AddRow(patSet);
                    newPatSet.AddRow(patSet);
                    AddPatSet(newSheet, newPatSet);
                    foreach (PatSet patset in hardIpPatset.MissingPatSets)
                    {
                        AddPatSet(newSheet, patset);
                    }
                }
            }
            catch (Exception exception)
            {
                Response.Report("Writing HardIP PatSet file failed " + exception.Message, EnumMessageLevel.Error, 95);
            }

            return newSheet;
        }

        internal void AddPatSet(PatSetSheet newSheet, PatSet newPatSet)
        {
            if (newSheet.IsExistTheSamePatSet(newPatSet, out _))
            {
                return;
            }

            if (newPatSet.PatSetRows.Count == 1 &&
                newPatSet.PatSetRows[0].File.Equals(newPatSet.PatSetName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            newSheet.AddRow(newPatSet);
        }

        internal string CheckBistOrBira(string pattern)
        {
            if (pattern.Contains("BIR"))
            {
                return "Bira";
            }

            return "Bist";
        }

        private string GetModuleName(string pattern, string module = "")
        {
            string[] sgmts = pattern.Split('_');
            string result = module;
            switch (sgmts[2].ToUpper())
            {
                case "L":
                    result = "Gfx";
                    break;
                case "C":
                    result = "Cpu";
                    break;
                case "S":
                    result = "Soc";
                    break;
            }

            return result;
        }

        internal string GenPllPatSetName(List<HardIpPattern> patterns, HardIpPattern burstPattern)
        {
            string module = GetModuleName(patterns.Last().Pattern.GetLastPayload());
            string block = GetBlock(patterns.Last().Pattern.GetLastPayload());
            foreach (HardIpPattern pattern in patterns)
            {
                if (string.IsNullOrEmpty(module))
                {
                    module = GetModuleName(pattern.Pattern.GetLastPayload());
                }

                if (string.IsNullOrEmpty(block))
                {
                    block = GetBlock(pattern.Pattern.GetLastPayload());
                }
            }
            if (string.IsNullOrEmpty(module))
            {
                module = "Cpu";
            }

            if (string.IsNullOrEmpty(block))
            {
                block = "Sa";
            }

            var nameList = new List<string> { module + block, burstPattern.Pattern.GetLastPayload() };
            if (burstPattern.DupIndex > 0)
            {
                nameList.Add(burstPattern.DupIndex.ToString());
            }

            return string.Join("_", nameList);
        }

        private string GetBlock(string pattern)
        {
            ScghData scgh = ScghStatic.ScghData;
            string[] sgmts = pattern.ToUpper().Split('_');
            if (scgh.GetScanPatterns.Exists(p => p.PayloadValue.Equals(pattern, StringComparison.OrdinalIgnoreCase)))
            {
                return "Sa";
            }

            if (scgh.GetBistPatterns.Exists(p => p.PayloadValue.Equals(pattern, StringComparison.OrdinalIgnoreCase)))
            {
                return "mbist";
            }

            if (Regex.IsMatch(sgmts[4], "sc|ch", RegexOptions.IgnoreCase))
            {
                return GetScanType(pattern);
            }

            if (Regex.IsMatch(sgmts[4], "bi", RegexOptions.IgnoreCase))
            {
                return CheckBistOrBira(pattern);
            }

            return "";
        }

        internal string GetScanType(string pattern)
        {
            string[] sgmts = pattern.Split('_');
            if (sgmts.Length <= 7)
            {
                return "";
            }

            string result = "";
            if (sgmts[4].Equals("sc", StringComparison.OrdinalIgnoreCase) &&
                sgmts[6].Equals("tdf", StringComparison.OrdinalIgnoreCase))
            {
                result = "Td";
            }

            if (sgmts[4].Equals("ch", StringComparison.OrdinalIgnoreCase) &&
                sgmts[6].Equals("tdf", StringComparison.OrdinalIgnoreCase))
            {
                result = "TdChain";
            }

            if ((sgmts[4].Equals("sc", StringComparison.OrdinalIgnoreCase) &&
                 sgmts[6].Equals("saa", StringComparison.OrdinalIgnoreCase)) ||
                sgmts[6].Equals("bdf", StringComparison.OrdinalIgnoreCase))
            {
                result = "Sa";
            }

            if (sgmts[4].Equals("ch", StringComparison.OrdinalIgnoreCase) &&
                sgmts[6].Equals("saa", StringComparison.OrdinalIgnoreCase))
            {
                result = "SaChain";
            }

            return result;
        }
    }
}
