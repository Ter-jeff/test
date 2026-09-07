using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.DividerManager;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Utility;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz
{
    public class FreqPllBlockInsGenerator : BlockInstanceGenerator
    {
        public FreqPllBlockInsGenerator(HardIpInputData hardIpInputData, string sheetName, HardIpSheet hardIpSheet)
            : base(hardIpInputData, sheetName, hardIpSheet)
        {
            InstanceRowGenerator = new FreqInsRowGenerator(hardIpInputData, hardIpSheet, sheetName);
            ((FreqInsRowGenerator)InstanceRowGenerator).BlockScbi = Block;
        }

        private string _module = "";
        public string Block = "";

        public override List<InstanceSheet> GenBlockInsRows()
        {
            return GenBlockInsRows(HardIpConstData.LabelNv);
        }

        public List<InstanceSheet> GenBlockInsRows(string voltage)
        {
            List<InstanceSheet> instanceSheetList = new List<InstanceSheet>();
            List<HardIpPattern> patLstToGen = DividerMain.DivideInstancePattern(HardIpInputData, HardIpSheet.Rows).ToList();

            try
            {
                instanceSheetList.AddRange(GenBlockInsRowsByBlock(patLstToGen, voltage));
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
            return instanceSheetList;
        }

        private List<InstanceSheet> GenBlockInsRowsByBlock(List<HardIpPattern> patternList, string labelVoltage)
        {
            List<InstanceSheet> instanceSheetList = new List<InstanceSheet>();
            string blockName = CommonGenerator.GetHardipSheetName(SheetName).ToUpper();
            var block = new InstanceSheet("TestInst_" + blockName, SheetName);

            var instances = new List<InstanceRow>();
            var patterns = new List<HardIpPattern>();
            foreach (HardIpPattern pattern in patternList)
            {
                try
                {
                    InstanceRowGenerator.LabelVoltage = labelVoltage;
                    InstanceRowGenerator.Pat = pattern;
                    patterns.Add(pattern);
                    List<InstanceRow> insRowList = InstanceRowGenerator.GenInsRows();
                    foreach (InstanceRow insRow in insRowList)
                    {
                        if (!(Regex.IsMatch(insRow.TestName, "INSREMOV_", RegexOptions.IgnoreCase) &&
                              labelVoltage != "NV"))
                        {
                            insRow.TestName = insRow.TestName.Replace("INSREMOV_", "");
                        }
                        instances.Add(insRow);

                        try
                        {
                            //apply all AC DC levels into set
                            //_ApplySetACCategory(instances);
                            ApplySetDcCategory(instances, patterns.Where(p => !p.Pattern.GetLastPayload().Contains(":")).ToList());
                            ApplySetLevels(instances);
                        }
                        catch (Exception ex)
                        {
                            ErrorMessageBox.Show(string.Format(ex.ToString()));
                        }

                        int patSetIndex = patterns.Select(p => p.Pattern.GetLastPayload()).ToList().FindIndex(p => Regex.IsMatch(p, "_[D]*SRA*M", RegexOptions.IgnoreCase));
                        if (patSetIndex == -1)
                        {
                            patSetIndex = 0;
                        }

                        block.AddRow(instances[patSetIndex]);
                        if (instances.Count > 1)
                        {
                            block.AddRow(insRow);
                        }

                        patterns.Clear();
                        instances.Clear();
                        _module = "";
                    }
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }

            if (block.Rows.Count != 0)
            {
                instanceSheetList.Add(block);
            }

            return instanceSheetList;
        }

        private void ApplySetDcCategory(List<InstanceRow> instanceList, List<HardIpPattern> allPatterns)
        {
            string property = instanceList.Select(p => p.DcCategory).FirstOrDefault(p => !string.IsNullOrEmpty(p));
            if (allPatterns.Count == 0)
            {
                return;
            }

            var patterns = allPatterns.Select(p => p.Pattern.GetLastPayload()).ToList();
            if (string.IsNullOrEmpty(property))
            {
                string pMode = "";
                string payloadType = GetScanType(patterns.Last(p => !p.Contains(":")));
                _module = GetModuleName(patterns.Last(p => !p.Contains(":")));
                BlockType block = GetBlock(patterns.Last(p => !p.Contains(":")));
                foreach (string pattern in patterns)
                {
                    if (string.IsNullOrEmpty(pMode))
                    {
                        pMode = PerformanceModeSingleton.Instance().FindPerformanceMode(pattern);
                    }

                    if (block == BlockType.HardIp)
                    {
                        block = GetBlock(pattern);
                    }

                }
                property = SummaryDcCategory(patterns, block, pMode, payloadType);
            }
            instanceList.ForEach(p => p.DcCategory = property);
        }

        private void ApplySetLevels(List<InstanceRow> instanceList)
        {
            string property = instanceList.Select(p => p.PinLevels).FirstOrDefault(p => !string.IsNullOrEmpty(p));
            instanceList.ForEach(p => p.PinLevels = property);
        }

        private string SummaryDcCategory(List<string> patterns, BlockType block, string pMode, string payloadType)
        {
            string dc = "";
            if (block == BlockType.Scan)
            {
                if (payloadType == "")
                {
                    payloadType = "Sa";
                }

                if (pMode != "")
                {
                    dc = MultiTestSettingSheetsSingleton.Instance().FindScanCategoryName(payloadType, _module, pMode, patterns, out EnumMessageLevel _, out string _);
                }
                if (dc == "")
                {
                    List<string> defaultList = payloadType.Equals("Sa", StringComparison.OrdinalIgnoreCase) ? new List<string> { "Sa", "Td" } : new List<string> { "Td", "Sa" };
                    foreach (string item in defaultList)
                    {
                        if (dc == "")
                        {
                            dc = MultiTestSettingSheetsSingleton.Instance().FindScanCategoryName(item, _module,
                                    pMode, patterns, out EnumMessageLevel _, out string _);
                        }
                    }
                }

            }
            else if (block == BlockType.Mbist)
            {
                dc = ExceptionListSingleton.Instance().GetDcCategoryByInstance(patterns.Last());
                if (string.IsNullOrEmpty(dc))
                {
                    dc = MultiTestSettingSheetsSingleton.Instance().FindMbistCatgeoryName(_module, CheckBistOrBira(patterns.Last()), pMode, patterns, out EnumMessageLevel _, out string _, _module);
                }
            }

            return dc;
        }

        private string CheckBistOrBira(string pattern)
        {
            if (pattern.Contains("BIR"))
            {
                return "Bira";
            }

            return "Bist";
        }

        private string GetModuleName(string pattern)
        {
            string[] arr = pattern.Split('_');
            if (arr.Length <= 5)
            {
                return "";
            }

            string result = "cpu";
            switch (arr[2].ToUpper())
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

        private BlockType GetBlock(string pattern)
        {
            if (pattern.Split('_').Length < 5)
            {
                return BlockType.HardIp;
            }

            ScghData scghData = ScghStatic.ScghData;
            string[] arr = pattern.ToUpper().Split('_');
            if (scghData.GetScanPatterns.Exists(p => p.PayloadValue.Equals(pattern, StringComparison.OrdinalIgnoreCase)))
            {
                return BlockType.Scan;
            }

            if (scghData.GetBistPatterns.Exists(p => p.PayloadValue.Equals(pattern, StringComparison.OrdinalIgnoreCase)))
            {
                return BlockType.Mbist;
            }

            if (Regex.IsMatch(arr[4], "sc|ch", RegexOptions.IgnoreCase))
            {
                return BlockType.Scan;
            }

            if (Regex.IsMatch(arr[4], "bi", RegexOptions.IgnoreCase))
            {
                return BlockType.Mbist;
            }

            return BlockType.HardIp;
        }

        private string GetScanType(string pattern)
        {
            string[] arr = pattern.Split('_');
            if (arr.Length <= 7)
            {
                return "";
            }

            string result = "";
            if (arr[4].Equals("sc", StringComparison.OrdinalIgnoreCase) &&
                arr[6].Equals("tdf", StringComparison.OrdinalIgnoreCase))
            {
                result = "Td";
            }

            if (arr[4].Equals("ch", StringComparison.OrdinalIgnoreCase) &&
                arr[6].Equals("tdf", StringComparison.OrdinalIgnoreCase))
            {
                result = "TdChain";
            }

            if ((arr[4].Equals("sc", StringComparison.OrdinalIgnoreCase) &&
                 arr[6].Equals("saa", StringComparison.OrdinalIgnoreCase)) ||
                arr[6].Equals("bdf", StringComparison.OrdinalIgnoreCase))
            {
                result = "Sa";
            }

            if (arr[4].Equals("ch", StringComparison.OrdinalIgnoreCase) &&
                arr[6].Equals("saa", StringComparison.OrdinalIgnoreCase))
            {
                result = "SaChain";
            }

            return result;
        }
    }
}
