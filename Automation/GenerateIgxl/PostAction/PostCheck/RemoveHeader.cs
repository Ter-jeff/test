using System;
using System.Collections.Generic;

using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.PostCheck
{
    public class RemoveHeader
    {
        public void WorkFlow()
        {
            var removes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, InstanceSheet> sheet in TestProgram.IgxlWorkBk.InsSheets)
            {
                for (int i = sheet.Value.Rows.Count - 1; i >= 0; i--)
                {
                    InstanceRow row = sheet.Value.Rows[i];
                    if (row.VbtName.Equals("Print_Header", StringComparison.CurrentCultureIgnoreCase) || row.VbtName.Equals("Print_Footer", StringComparison.CurrentCultureIgnoreCase))
                    {
                        removes.Add(row.TestName);
                        sheet.Value.Rows.RemoveAt(i);
                    }
                }
            }

            foreach (KeyValuePair<string, SubFlowSheet> sheet in TestProgram.IgxlWorkBk.SubFlowSheets)
            {
                string name = sheet.Value.Name.StartsWithIgnoreCase("Flow_") ? sheet.Value.Name.Substring(5) : sheet.Value.Name;
                for (int i = sheet.Value.Rows.Count - 1; i >= 0; i--)
                {
                    FlowRow row = sheet.Value.Rows[i];
                    if (row.Parameter.Equals(name + "_Header_1", StringComparison.CurrentCultureIgnoreCase) || row.Parameter.Equals(name + "_Footer_1", StringComparison.CurrentCultureIgnoreCase))
                    {
                        sheet.Value.Rows.RemoveAt(i);
                    }
                    else if (row.Parameter.Equals(sheet.Value.Name + "_Header_1", StringComparison.CurrentCultureIgnoreCase) || row.Parameter.Equals(sheet.Value.Name + "_Footer_1", StringComparison.CurrentCultureIgnoreCase))
                    {
                        sheet.Value.Rows.RemoveAt(i);
                    }
                    else if (removes.Contains(row.Parameter))
                    {
                        sheet.Value.Rows.RemoveAt(i);
                    }
                }
            }

            foreach (KeyValuePair<string, InstanceSheet> sheet in TestProgram.IgxlWorkBk.InsSheets)
            {
                foreach (InstanceRow row in sheet.Value.Rows)
                {
                    row.SheetName = sheet.Value.Name;
                }
            }
        }
    }
}
