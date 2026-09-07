using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlSheets;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz
{
    public class InstanceGenerator
    {
        protected readonly HardIpInputData HardIpInputData;

        public InstanceGenerator(HardIpInputData hardIpInputData)
        {
            HardIpInputData = hardIpInputData;
        }

        public List<InstanceSheet> GenInst(Dictionary<string, HardIpSheet> planDic)
        {
            var instSheets = new List<InstanceSheet>();
            foreach (string sheetName in planDic.Keys)
            {
                BlockInstanceGenerator blockInstanceGenerator = HardipInstanceFactory(planDic, sheetName);
                instSheets.AddRange(blockInstanceGenerator.GenBlockInsRows());
            }

            instSheets = MergeInstSheet(instSheets.ToList());
            instSheets = DivideInstanceSheet(instSheets);
            return instSheets;
        }

        private BlockInstanceGenerator HardipInstanceFactory(Dictionary<string, HardIpSheet> planDic, string sheetName)
        {
            BlockInstanceGenerator blockInstanceGenerator;
            if (Regex.IsMatch(sheetName, NeededSheets.HardIpPllMeas, RegexOptions.IgnoreCase))
            {
                blockInstanceGenerator = new FreqPllBlockInsGenerator(HardIpInputData, sheetName, planDic[sheetName]);
            }
            else if (SearchInfo.IsHardipIdsSheet(sheetName))
            {
                blockInstanceGenerator = new IdsBlockInsGenerator(HardIpInputData, sheetName, planDic[sheetName]);
            }
            else if (SearchInfo.IsHardipRtosSheet(sheetName))
            {
                blockInstanceGenerator = new RtosBlockInsGenerator(HardIpInputData, sheetName, planDic[sheetName]);
            }
            else
            {
                blockInstanceGenerator = new HardIpBlockInsGenerator(HardIpInputData, sheetName, planDic[sheetName]);
            }

            return blockInstanceGenerator;
        }

        private List<InstanceSheet> MergeInstSheet(List<InstanceSheet> instSheets)
        {
            var dict = new Dictionary<string, InstanceSheet>();
            foreach (InstanceSheet sheet in instSheets.Where(s => s.Rows.Count > 0))
            {
                if (dict.TryGetValue(sheet.Name, out InstanceSheet existing))
                {
                    existing.Rows.AddRange(sheet.Rows);
                }
                else
                {
                    dict[sheet.Name] = sheet;
                }
            }

            foreach (InstanceSheet sheet in dict.Values)
            {
                sheet.Rows = new InstanceRows(sheet.Rows.OrderBy(row => row.RowNum));
            }

            return dict.Values.ToList();

        }

        private List<InstanceSheet> DivideInstanceSheet(List<InstanceSheet> instSheets)
        {
            List<InstanceSheet> instanceSheets = new List<InstanceSheet>();
            foreach (InstanceSheet sheet in instSheets)
            {
                if (sheet.Rows.Count > 10000)
                {
                    int loopCnt = sheet.Rows.Count / 10000;
                    int lastRowCount = sheet.Rows.Count % 10000;

                    for (int i = 0; i < loopCnt; i++)
                    {
                        var subInstSheet = new InstanceSheet($"{sheet.Name}_{i + 1}");
                        subInstSheet.AddRows(sheet.Rows.GetRange(i * 10000, 10000));
                        instanceSheets.Add(subInstSheet);
                    }
                    var finalInstSheet = new InstanceSheet($"{sheet.Name}_{loopCnt + 1}");
                    finalInstSheet.AddRows(sheet.Rows.GetRange(loopCnt * 10000, lastRowCount));
                    instanceSheets.Add(finalInstSheet);
                }
                else
                {
                    instanceSheets.Add(sheet);
                }
            }
            return instanceSheets;
        }
    }
}
