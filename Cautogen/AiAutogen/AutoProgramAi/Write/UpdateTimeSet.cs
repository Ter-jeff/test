using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Cautogen.common.IgxlDataExtension;

using DebugPlanReaderLib.DebugPlan;

using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

namespace Cautogen.AiAutogen.AutoProgramAi.Write
{
    public class UpdateTimeSet
    {
        public List<ComTimeSetBasicSheet> Work(DebugPlanMain debugTestPlan, List<TimeSetBasicSheet> timeSetBasicSheets,
            PortMapSheet portMapSheet,
            string patternFolder)
        {
            if (!Directory.Exists(Path.Combine(patternFolder, @"Timeset\")))
                return new List<ComTimeSetBasicSheet>();

            var timeSets = debugTestPlan.AiTestPlanSheets.SelectMany(x => x.Rows)
                .Where(x => x.TimesetMapping.Any()).Select(x => x.TimesetMapping.First()).ToList();

            var emptyTimeSets = new List<string>();
            var timeSetEmptyRows = debugTestPlan.AiTestPlanSheets.SelectMany(x => x.Rows)
                .Where(x => !x.TimesetMapping.Any()).ToList();
            foreach (var timeSetEmptyRow in timeSetEmptyRows)
            {
                foreach (var pattern in timeSetEmptyRow.Payloads)
                {
                    if (debugTestPlan.PatternListSheet.Rows.Exists(x =>
                            x.Pattern.Equals(pattern.OriName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var timeset = debugTestPlan.PatternListSheet.Rows.Find(x =>
                            x.Pattern.Equals(pattern.OriName, StringComparison.CurrentCultureIgnoreCase)).TimeSet;
                        emptyTimeSets.Add(timeset);
                    }
                }
            }

            timeSets.AddRange(emptyTimeSets); //timesets of plan instances
            timeSets = timeSets.Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();

            //var existedTimeSets = Directory.GetFiles(Path.Combine(patternFolder, @"Timeset\")).Select(x=>Path.GetFileNameWithoutExtension(x)); //timesets in timeset folder
            //timeSetBasicSheets.Where(x=> !existedTimeSets.Any(y => string.Equals(y, x.Name, StringComparison.OrdinalIgnoreCase))).ToList()
            //.ForEach(x => x.Write(Path.Combine(patternFolder, @"Timeset\" + x.Name + ".TXT")));

            var timeSetPaths = timeSets.Except(timeSetBasicSheets.Select(x => x.Name))
                .Select(x => Path.Combine(patternFolder, @"Timeset\" + x + ".txt")).ToList();

            MultiTimeSetSheets comTimeSetBasicSheets = MultiTimeSetSheetReader.ReadTimeSetTxt1P4(timeSetPaths);

            #region add port set

            var allTsets = timeSetBasicSheets.SelectMany(x => x.Rows).ToList();
            var tsets = new List<ComTimeSetBasic>();
            foreach (var portSets in portMapSheet.Rows)
            {
                foreach (var tset in allTsets)
                {
                    if (portSets.PortName.Equals(tset.Name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        var comTimeSetBasic = new ComTimeSetBasic();
                        comTimeSetBasic.Name = tset.Name;
                        comTimeSetBasic.CyclePeriod = tset.CyclePeriod;
                        comTimeSetBasic.AddTimingRows(tset.TimingRows);
                        tsets.Add(comTimeSetBasic);
                        break;
                    }
                }
            }

            if (tsets.Count > 0)
            {
                foreach (var comTimeSetBasicSheet in comTimeSetBasicSheets.TimeSetBasicSheetsList)
                {
                    foreach (var tset in tsets)
                    {
                        if (!comTimeSetBasicSheet.Rows.Exists(x =>
                                x.Name.Equals(tset.Name, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            comTimeSetBasicSheet.Rows.Add(tset);
                        }
                    }
                }
            }

            #endregion

            return comTimeSetBasicSheets.TimeSetBasicSheetsList;
        }
    }
}
