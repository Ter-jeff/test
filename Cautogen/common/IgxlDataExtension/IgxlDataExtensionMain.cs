using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

namespace Cautogen.common.IgxlDataExtension
{
    public static class IgxlDataExtensionMain
    {
        public static void AddInstanceSheet(this JobListSheet joblistSheet, string name)
        {
            foreach (JobRow row in joblistSheet.Rows)
            {
                if (string.IsNullOrEmpty(row.TestInstances))
                {
                    row.TestInstances = name;
                }
                else
                {
                    row.TestInstances += "," + name;
                }
            }
        }

        public static void AddCharSheet(this JobListSheet joblistSheet, string name)
        {
            foreach (JobRow row in joblistSheet.Rows)
            {
                if (string.IsNullOrEmpty(row.Characterization))
                {
                    row.Characterization = name;
                }
                else
                {
                    row.Characterization += "," + name;
                }
            }
        }

        public static string AcCatalogContainsTimeSet(this AcSpecSheet acSpecSheet, ComTimeSetBasicSheet comTimeSetBasicSheet, string category = "")
        {
            IEnumerable<string> cats = string.IsNullOrEmpty(category) ?
                acSpecSheet.Rows.SelectMany(x => x.CategoryList.Select(y => y.Name)).Distinct()
                : new List<string> { category };
            foreach (string cat in cats)
            {
                bool flag = true;
                var items = acSpecSheet.Rows.SelectMany(x => x.CategoryList.Where(y => y.Name == cat)).ToList();

                foreach (ComTimeSetBasicSheet.TSetEqnVarMap item in comTimeSetBasicSheet.AllTSetEqnVariable)
                {
                    foreach (KeyValuePair<string, double> variable in item.DictVariable)
                    {
                        if (items.Exists(x => x.Name.Equals(variable.Key)))
                        {
                            if (items.Find(x => x.Name.Equals(variable.Key)).Typ != variable.Value.ToString())
                            {
                                flag = false;
                            }
                        }
                    }
                    if (flag)
                    {
                        return cat;
                    }
                }
            }
            return "";
        }

        public static void AddTimingRows(this ComTimeSetBasic comTimeSetBasic, List<TimingRow> timingRows)
        {
            comTimeSetBasic.TimingRows.AddRange(timingRows);
        }

        public static void AddRows(this PatSetSubSheet patSetSubSheet, List<PatSetSubRow> patSetSubRows)
        {
            patSetSubSheet.Rows.AddRange(patSetSubRows);
        }

     

        public static List<string> GetEnables(this ReadFlowSheet readFlowSheet, Stream stream, string sheetName)
        {
            var enables = new List<string>();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    if (line != null)
                    {
                        string[] arr = line.Split(new[] { '\t' }, StringSplitOptions.None);
                        if (arr.Length > 3)
                        {
                            string enable = arr[2];
                            if (!string.IsNullOrEmpty(enable))
                            {
                                enables.Add(enable);
                            }
                        }

                    }
                }
            }
            return enables.Distinct().ToList();
        }
    }
}
