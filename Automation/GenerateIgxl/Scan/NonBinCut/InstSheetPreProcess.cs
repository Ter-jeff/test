using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.Scan.Harvest;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Singleton;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.IgxlBase;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Utility;

namespace Automation.GenerateIgxl.Scan.NonBinCut
{
    public class InstSheetPreProcess
    {
        private readonly ScanConfig _scanConfig;
        protected static readonly Regex Regex = new Regex(PerformanceModeSingleton.RegContainPerformanceModeByPattern, RegexOptions.IgnoreCase);

        public InstSheetPreProcess(ScanConfig scanConfig)
        {
            _scanConfig = scanConfig;
        }

        public virtual BinCutFinalInstanceRows InitialInstance(List<BinCutInstanceSheet> binCutInstanceSheets, BinCutInstanceNamingSheet binCutInstanceNamingSheet)
        {
            var rows = new BinCutFinalInstanceRows();
            foreach (BinCutInstanceSheet binCutInstanceSheet in binCutInstanceSheets)
            {
                foreach (BinCutInstanceRow binCutInstanceRow in binCutInstanceSheet.Rows)
                {
                    try
                    {
                        //PatSet Name = DomainBlock_Mode_skipKey_init_pattern
                        string block = GetBlock(binCutInstanceRow);
                        string domain = BinCutInstanceRowUtility.GetDomain(binCutInstanceRow);
                        string mode = GetMode(binCutInstanceRow);
                        string initName = binCutInstanceNamingSheet != null ? binCutInstanceNamingSheet.GetInitName(binCutInstanceRow, mode) : "";
                        string payloadName = GetPayloadName(binCutInstanceRow);
                        string instance = binCutInstanceRow.Instance;
                        string patSetName = GetPatSetName(binCutInstanceRow, domain, block, mode, initName, payloadName, instance);
                        string initPatSetName = $"{patSetName}_INIT";
                        BinCutFinalInstanceRow row = GetBinCutFinalInstanceRow(domain, block, mode, initName, payloadName, patSetName, binCutInstanceRow, initPatSetName);
                        if (patSetName.Length > 140)
                        {
                            binCutInstanceSheet.AddError(ScanErrorType.E_InstanceName_01, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, 1, "Instance name too long(over 150), please modify instance name by PatSetName(Orange)!!!");
                        }

                        rows.Add(row);
                    }
                    catch (Exception)
                    {
                        //pass
                    }
                }
            }
            return rows;
        }

        internal virtual string GetPatSetName(BinCutInstanceRow binCutInstanceRow, string domain, string block, string mode,
            string initName, string payloadName, string instance)
        {
            string lastName = !string.IsNullOrEmpty(binCutInstanceRow.PatSetNameOrange) ? Combination.CombineByUnderLine(new List<string> { binCutInstanceRow.PatSetNameOrange }) : Combination.CombineByUnderLine(new List<string> { initName, payloadName });

            if (string.IsNullOrEmpty(binCutInstanceRow.PatSetNameOrange) && !string.IsNullOrEmpty(instance))
            {
                lastName += "_" + instance;
            }

            string patSetName = !string.IsNullOrEmpty(binCutInstanceRow.SplitKey) ? Combination.CombineByUnderLine(new List<string> { domain + block, mode, binCutInstanceRow.SplitKey, lastName }) : Combination.CombineByUnderLine(new List<string> { domain + block, mode, lastName });
            return patSetName;
        }

        protected virtual BinCutFinalInstanceRow GetBinCutFinalInstanceRow(string domain, string block, string mode, string initName, string payloadName, string patSetName, BinCutInstanceRow binCutInstanceRow, string initPatSetName = null)
        {
            var row = new BinCutFinalInstanceRow
            {
                Domain = domain,
                Block = block,
                ModeByFlowName = mode,
                InitName = initName,
                PayloadName = payloadName,
                BinCutInstanceRow = binCutInstanceRow
            };
            row.InitList.AddRange(binCutInstanceRow.InitList);
            row.PayloadList.AddRange(binCutInstanceRow.PayloadList);
            row.InitPatSetNameNew = initPatSetName;
            row.InitPatSetNameNewTemp = initPatSetName;
            row.PatSetName = patSetName;
            row.PatSetNameTemp = patSetName;
            var initModes = binCutInstanceRow.InitList.Where(x => x.IsInit()).Select(BinCutInstanceRowUtility.GetMode).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList();
            string finalMode = !initModes.Any() ? "" : initModes.Last();
            row.PerformanceMode = string.IsNullOrEmpty(finalMode) ? mode : finalMode;
            row.PatternList.AddRange(binCutInstanceRow.PatternList);
            row.JudgeBurst();
            return row;
        }

        internal virtual string GetPayloadName(BinCutInstanceRow binCutInstanceRow)
        {
            string payloadName = binCutInstanceRow.PayloadList.Any() ? PatSet.GetPatSetName(binCutInstanceRow.PayloadList) : PatSet.GetPatSetName(binCutInstanceRow.PatternList, false);
            return payloadName;
        }

        public virtual string GetBlock(BinCutInstanceRow binCutInstanceRow)
        {
            string block = "";
            if (binCutInstanceRow.PayloadList.Any())
            {
                block = new ScanNonBinCutInstanceMain(_scanConfig).GetPayloadType(binCutInstanceRow.PayloadList.First());
            }
            else if (binCutInstanceRow.InitList.Any())
            {
                block = new ScanNonBinCutInstanceMain(_scanConfig).GetPayloadType(binCutInstanceRow.InitList.First());
            }

            if (string.IsNullOrEmpty(block))
            {
                block = BinCutInstanceRowUtility.GetBlockByFlowName(binCutInstanceRow.FlowName);
            }

            return block;
        }

        public virtual string GetMode(BinCutInstanceRow binCutInstanceRow)
        {
            string[] parts = binCutInstanceRow.FlowName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return parts.FirstOrDefault(Regex.IsMatch) ?? string.Empty;
        }
    }
}
