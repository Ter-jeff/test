using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using LogLib.Static;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class DuplicateStoreNameChecker
    {
        private class DupItem
        {
            public readonly string Sheet = "";
            public readonly int Row;

            public DupItem(string sheet, int row)
            {
                Sheet = sheet;
                Row = row;
            }
        }

        public void CheckStoreName(List<HardIpPattern> patterns)
        {
            var allblkmultipats = new HashSet<string>();
            //Block<StoreName , <SheetName, Pattern, Row>>
            var mutidupstoname = new List<Dictionary<string, List<DupItem>>>();
            var singpatterns = new List<HardIpPattern>();
            foreach (HardIpPattern pat in patterns)
            {
                if (new[] { "+", "#" }.Any(pat.Pattern.RealPatternName.Contains))
                {
                    allblkmultipats.Add(pat.Pattern.RealPatternName);
                }
                else
                {
                    singpatterns.Add(pat);
                }
            }

            foreach (string mutpat in allblkmultipats)
            {
                var dupstoname = new Dictionary<string, List<DupItem>>();
                var pats = new List<HardIpPattern>();
                pats = singpatterns.FindAll(singpat => !new[] { "+", "#" }.Any(singpat.Pattern.RealPatternName.Contains) && mutpat.Contains(singpat.Pattern.RealPatternName));
                foreach (HardIpPattern pat in pats)
                {
                    string sheet = pat.SheetName;
                    var cusstrs = pat.MeasPins.Select(pin => pin.CusStr).ToList();
                    var hashcusstrs = new HashSet<string>();
                    for (int i = 0; i < cusstrs.Count; i++)
                    {
                        if (cusstrs[i] == "" || !hashcusstrs.Add(cusstrs[i]))
                        {
                            continue;
                        }

                        int row = pat.RowNum + 1 + i;

                        if (dupstoname.ContainsKey(cusstrs[i]))
                        {
                            dupstoname[cusstrs[i]].Add(new DupItem(sheet, row));
                        }
                        else
                        {
                            dupstoname.Add(cusstrs[i], new List<DupItem> { new DupItem(sheet, row) });
                        }
                    }
                }
                if (dupstoname.Count > 0)
                {
                    mutidupstoname.Add(dupstoname);
                }
            }

            //StoreName , <SheetName, Pattern, Row>
            var singdupstoname = new Dictionary<string, List<DupItem>>();
            foreach (HardIpPattern pat in singpatterns)
            {
                string sheet = pat.SheetName;
                var cusstrs = pat.MeasPins.Select(pin => pin.CusStr).ToList();
                var hashcusstrs = new HashSet<string>();
                for (int i = 0; i < cusstrs.Count; i++)
                {
                    if (cusstrs[i] == "" || !hashcusstrs.Add(cusstrs[i]))
                    {
                        continue;
                    }

                    int row = pat.RowNum + 1 + i;

                    if (singdupstoname.ContainsKey(cusstrs[i]))
                    {
                        singdupstoname[cusstrs[i]].Add(new DupItem(sheet, row));
                    }
                    else
                    {
                        singdupstoname.Add(cusstrs[i], new List<DupItem> { new DupItem(sheet, row) });
                    }
                }
            }

            bool isDuplicate = false;
            foreach (Dictionary<string, List<DupItem>> dupstoname in mutidupstoname)
            {
                foreach (KeyValuePair<string, List<DupItem>> storename in dupstoname)
                {
                    if (storename.Value.Count < 2)
                    {
                        continue;
                    }

                    foreach (DupItem dupitem in storename.Value)
                    {
                        isDuplicate = true;
                        string msg = $"We found duplicate Store name \"{storename.Key}\" in HardIP multiple pattern.";
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_DuplicateStoreName_01,
                            dupitem.Sheet,
                            dupitem.Row,
                            9,
                            [storename.Key]
                        );
                    }
                }
            }

            foreach (KeyValuePair<string, List<DupItem>> storename in singdupstoname)
            {
                if (storename.Value.Count < 2)
                {
                    continue;
                }

                foreach (DupItem dupitem in storename.Value)
                {
                    isDuplicate = true;
                    ErrorReportManager.AddError(
                        HardIpErrorType.I_DuplicateStoreName_01,
                        dupitem.Sheet,
                        dupitem.Row,
                        9,
                        [storename.Key]
                    );
                }
            }

            if (isDuplicate)
            {
                Response.Report("We found duplicate Store name in HardIP patterns, please check.", EnumMessageLevel.Warning, 0);
            }
        }
    }
}
