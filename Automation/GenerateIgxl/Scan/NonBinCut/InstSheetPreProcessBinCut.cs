using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.Reader.ConfigFile.NamingRule.Base;

using CommonLib.Enums;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Utility;

namespace Automation.GenerateIgxl.Scan.NonBinCut
{
    public class InstSheetPreProcessBinCut : InstSheetPreProcess
    {
        public InstSheetPreProcessBinCut(ScanConfig scanConfig) : base(scanConfig)
        {
        }

        public override BinCutFinalInstanceRows InitialInstance(List<BinCutInstanceSheet> bincutInstanceSheets, BinCutInstanceNamingSheet binCutInstanceNamingSheet)
        {
            var instDataRows = new BinCutFinalInstanceRows();
            foreach (BinCutInstanceSheet binCutInstanceSheet in bincutInstanceSheets)
            {
                foreach (BinCutInstanceRow binCutInstanceRow in binCutInstanceSheet.Rows)
                {
                    if (binCutInstanceRow.PayloadList.Count == 0 && binCutInstanceRow.Type == BincutInstanceType.Pattern)
                    {
                        continue;
                    }

                    string block = GetBlock(binCutInstanceRow);
                    string domain = BinCutInstanceRowUtility.GetDomain(binCutInstanceRow);
                    string mode = GetMode(binCutInstanceRow);
                    string initName = binCutInstanceNamingSheet != null ? binCutInstanceNamingSheet.GetInitName(binCutInstanceRow, mode) : "";
                    string payloadName = domain + block + "_" + GetPayloadNameNew(binCutInstanceRow, block);
                    string initPatSetName = "";
                    if (block == "Mbist")
                    {
                        initPatSetName = domain + block + "_" + GetInitNameNew(binCutInstanceRow);
                    }
                    string patSetName = !string.IsNullOrEmpty(binCutInstanceRow.PatSetNameOrange) ? binCutInstanceRow.PatSetNameOrange : payloadName;
                    if (binCutInstanceRow.Type == BincutInstanceType.Rtos && binCutInstanceRow.PayloadList.Any() &&
                        !string.IsNullOrEmpty(binCutInstanceRow.PayloadList.First()))
                    {
                        patSetName = binCutInstanceRow.PayloadList.First();
                    }

                    BinCutFinalInstanceRow row = GetBinCutFinalInstanceRow(domain, block, mode, initName, payloadName, patSetName, binCutInstanceRow, initPatSetName);
                    instDataRows.Add(row);
                    if (patSetName.Length > 140)
                    {
                        binCutInstanceSheet.AddError(BinCutErrorType.E_InstanceName_01, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, 1, "Instance name too long(over 150), please modify instance name by PatSetName(Orange)!!!");
                    }
                }
            }
            return instDataRows;
        }

        internal override string GetPayloadName(BinCutInstanceRow binCutInstanceRow)
        {
            var patterns = binCutInstanceRow.PayloadList.Where(x => x.GetPatternType() != EnumPatternType.Unknown && x.GetPatternType() != EnumPatternType.Init).ToList();
            string payloadName = PatSet.GetNewPatSetNameWithX(patterns);
            return payloadName;
        }

        internal string GetPayloadNameNew(BinCutInstanceRow binCutInstanceRow, string block)
        {
            string payloadName;
            if (block == "Mbist")
            {
                var patterns = binCutInstanceRow.PayloadList.Where(x => x.GetPatternType() != EnumPatternType.Unknown && x.GetPatternType() != EnumPatternType.Init).ToList();

                if (patterns.Count == 1)
                {
                    payloadName = PatSet.GetNewPatSetNameWithX(patterns) + "_SINGLE";
                }
                else
                {
                    payloadName = PatSet.GetNewPatSetNameWithX(patterns);
                }
            }
            else
            {
                var patterns = binCutInstanceRow.PayloadList.Where(x => x.GetPatternType() != EnumPatternType.Unknown).ToList();
                payloadName = PatSet.GetNewPatSetNameWithX(patterns);
            }
            return payloadName;
        }

        protected string GetInitNameNew(BinCutInstanceRow binCutInstanceRow)
        {
            var patterns = binCutInstanceRow.InitList.Where(x => x.GetPatternType() != EnumPatternType.Unknown && x.GetPatternType() != EnumPatternType.Payload).ToList();
            string initName = PatSet.GetNewPatSetNameWithX(patterns);
            return initName;

        }

        public override string GetBlock(BinCutInstanceRow binCutInstanceRow)
        {
            string block = "";
            if (binCutInstanceRow.PayloadList.Any())
            {
                block = GetBlockByPattern(binCutInstanceRow.PayloadList.First());
            }
            else if (binCutInstanceRow.InitList.Any())
            {
                block = GetBlockByPattern(binCutInstanceRow.InitList.First());
            }

            if (string.IsNullOrEmpty(block) && (binCutInstanceRow.FlowName.Split(' ').Length > 1 || binCutInstanceRow.FlowName.Split('_').Length > 1))
            {
                block = BinCutInstanceRowUtility.GetBlockByFlowName(binCutInstanceRow.FlowName);
            }

            return block;
        }

        internal string GetBlockByPattern(string pattern)
        {
            string[] split = pattern.ToUpper().Split('_');
            if (split.Contains("SC"))
            {
                return "Td";
            }

            if (split.Contains("BI"))
            {
                return "Mbist";
            }

            return "";
        }

        public override string GetMode(BinCutInstanceRow binCutInstanceRow)
        {
            string[] parts = binCutInstanceRow.FlowName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return parts.FirstOrDefault(Regex.IsMatch) ?? parts.FirstOrDefault() ?? string.Empty;
        }
    }
}
