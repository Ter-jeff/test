using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.common.IgxlDataExtension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class UpdateTimeSetSheet
    {
        public static void WorkFlow()
        {
            //var timeSetFiles = Directory.GetFiles(string.Format(@"K:\{0}\{1}", LocalSpecs.Project, "TimeSet"));
            if (!Directory.Exists(LocalSpecs.TimeSetFolder))
            {
                return;
            }

            string[] timeSetFiles = Directory.GetFiles(LocalSpecs.TimeSetFolder);

            IEnumerable<string> inputProgramFiles = Directory.GetFiles(LocalSpecs.OutputFolder, "*.*", SearchOption.AllDirectories)
                               .Where(x => x.EndsWith("txt", StringComparison.OrdinalIgnoreCase) ||
                                           x.EndsWith("bas", StringComparison.OrdinalIgnoreCase) ||
                                           x.EndsWith("cls", StringComparison.OrdinalIgnoreCase));

            #region find port row in program timeset
            var allTsets = LocalSpecs.TestProgram.TimeSetSheets.SelectMany(x => x.Rows).ToList();
            var portTsets = new List<ComTimeSetBasic>();
            foreach (PortRow portSets in LocalSpecs.TestProgram.PortRows)
            {
                foreach (TSet tset in allTsets)
                {
                    if (!string.IsNullOrEmpty(tset.Name) &&
                        portSets.PortName.Equals(tset.Name, StringComparison.CurrentCultureIgnoreCase) &&
                        tset.TimingRows.Any(x => string.Equals(x.PinGrpName, portSets.FunctionPin, StringComparison.OrdinalIgnoreCase)))
                    {
                        var comTimeSetBasic = new ComTimeSetBasic { Name = tset.Name, CyclePeriod = tset.CyclePeriod };
                        comTimeSetBasic.AddTimingRows(tset.TimingRows);
                        portTsets.Add(comTimeSetBasic);
                        break;
                    }
                }
            }
            #endregion

            foreach (string tFile in timeSetFiles)
            {
                string tFileName = Path.GetFileName(tFile);
                if (tFileName.StartsWith("TIMESET") &&
                    !inputProgramFiles.Any(x => string.Equals(Path.GetFileName(x), tFileName, StringComparison.OrdinalIgnoreCase)))
                {
                    ComTimeSetBasicSheet nTFile = MultiTimeSetSheetReader
                        .ReadTimeSetTxt1P4(new List<string>() { tFile })
                        .TimeSetBasicSheetsList.FirstOrDefault();
                    foreach (ComTimeSetBasic portTset in portTsets)
                    {
                        TSet nTFilePortRow =
                            nTFile.Rows
                            .FirstOrDefault(x => x.Name.Equals(portTset.Name, StringComparison.OrdinalIgnoreCase) &&
                                                 x.TimingRows.Any(y =>
                                                                    portTset.TimingRows.Any(z =>
                                                                                            z.PinGrpName.Equals(y.PinGrpName, StringComparison.OrdinalIgnoreCase))));
                        if (nTFilePortRow == null)
                        {
                            nTFile.Rows.Add(portTset);
                        }
                    }
                    nTFile.Write(Path.Combine(LocalSpecs.OutputFolder, ConstData.TimeSetFolder, Path.GetFileName(tFile)));
                    LocalSpecs.GenSheets.Add(nTFile);
                }
            }
        }
    }
}
