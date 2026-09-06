using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Utility.HardIP;

namespace Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess
{
    internal class TestPlanSheetPatPreprocess
    {
        private readonly TestPlanSheet _planSheet;

        public TestPlanSheetPatPreprocess(TestPlanSheet planSheet)
        {
            _planSheet = planSheet;
        }

        /// <summary>
        /// 1.if [SubBlock]+[pattern name]+[Instance Substring] duplicated, add "_[Index]" to differentiate them
        /// </summary>
        public void UpdateSheetPattern()
        {
            var instInProductList = _planSheet.PatternRows.Where(x => x.ForceCondition.IsShmooInProdInst | !x.ForceCondition.IsShmooInForce && !x.Pattern.IsMultiTimeDomain()).ToList();
            GetPatternDupIndex(instInProductList);
            var mtdInstInProductList = _planSheet.PatternRows.Where(x => x.Pattern.IsMultiTimeDomain()).ToList();
            GetPatternDupIndex(mtdInstInProductList);
            var instInCharList = _planSheet.PatternRows.Where(x => x.ForceCondition.IsShmooInCharInst).ToList();
            GetPatternDupIndex(instInCharList);
        }

        /// <summary>
        /// if pattern name duplicated, add "_[Index]" to differentiate them
        /// </summary>
        /// <param name="planList"></param>
        private void GetPatternDupIndex(List<PatternRow> planList)
        {
            var patItems = new Dictionary<string, List<PatternRow>>();
            foreach (PatternRow row in planList)
            {
                string patternName = string.Empty;

                Match matchOpcodeInPatt = HardIpConstData.RegOpcodeInPatt.Match(row.Pattern.RealPatternName);
                if (matchOpcodeInPatt.Success)
                {
                    patternName = matchOpcodeInPatt.Groups["Parameter"].ToString();
                }
                else
                {
                    string blockName = CommonGenerator.GetBlockNameFromSheetName(row.SheetName);
                    string subBlock = CommonGenerator.GetSubBlockName(row.Pattern.GetLastPayload(), row.MiscInfo, blockName).ToLower();
                    string miscSubstring = HardIpService.GetInstNameSubStr(row.MiscInfo).ToLower();
                    bool cz = row.ForceCondition.IsCz2InstName;
                    patternName = subBlock + "_" + row.Pattern.GetLastPayload().ToLower() + "_" + miscSubstring + "_" + cz;
                }

                if (!patItems.ContainsKey(patternName))
                {
                    patItems.Add(patternName, new List<PatternRow>());
                }
                patItems[patternName].Add(row);
            }

            foreach (KeyValuePair<string, List<PatternRow>> patItem in patItems)
            {
                if (patItem.Value.Count == 1 ||
                    patItem.Value.Select(x => x.RowNum).Distinct().Count() == 1)
                {
                    continue;
                }

                int i = 1;

                foreach (IGrouping<int, PatternRow> grouping in patItem.Value.GroupBy(x => x.RowNum))
                {
                    foreach (PatternRow item in grouping)
                    {
                        item.DupIndex = i;
                    }

                    i++;
                }
            }
        }
    }
}
